"""Package a complete binder build and record verifiable release evidence."""
import argparse
import csv
from datetime import datetime, timezone
import hashlib
import importlib.metadata
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import zipfile

PACKAGER_VERSION = "1.0"
REQUIRED_FILES = (
    "Complete Till Binder - Updated - Photos.xlsx",
    "Complete Till Binder - Updated - Photos.pdf",
    "Pocket Guide.pdf",
    "OldTimers_Produce.pdf",
    "BINDER_BOOKLET_IMPOSITION.csv",
    "VALIDATION.json",
    "AUDIT_REPORT.json",
    "PRODUCTION_BLOCKERS.csv",
    "OWNER_REVIEW_FINDINGS.csv",
    "reconciliation/PRODUCTION_READINESS.json",
    "reconciliation/PICKLIST_SYNC.csv",
    "reconciliation/WORKING_LIST_RECONCILIATION.csv",
    "reconciliation/PICKLIST_RECORD_CHANGES.csv",
    "reconciliation/PICKLIST_SOURCE_RECORDS.xml",
)
GENERATOR_FILES = (
    "tools/build_till_binders.py",
    "tools/audit_till_binder.py",
    "tools/audit_production.py",
    "tools/build_oldtimers_produce.py",
    "tools/impose_binder_booklet.py",
    "tools/package_publication.py",
    ".github/workflows/package-publication.yml",
)
MANDATORY_COUNTERS = (
    "unresolved_source_record_reconciliation",
    "unresolved_product_identity",
    "unresolved_image_selection",
    "classification_section_reviews",
)
MANDATORY_GATES = ("data_reconciliation", "image_selection", "classification_placement")


class PackageError(ValueError):
    """Raised when a package is incomplete or fails its release gate."""


def sha256(path):
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_json(path, label):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise PackageError(f"{label} is missing or invalid: {error}") from error


def production_gate(readiness, validation, audit):
    blockers = []
    expected_status = "AUTOMATED_CHECKS_PASSED_REQUIRES_RELEASE_REVIEW"
    if readiness.get("status") != expected_status:
        blockers.append(
            f"production readiness report status is {readiness.get('status', 'NOT_REPORTED')}")
    release_blockers = readiness.get("release_blockers")
    if not isinstance(release_blockers, dict):
        raise PackageError("production readiness report has no release_blockers mapping")
    for name, count in sorted(release_blockers.items()):
        if int(count):
            blockers.append(f"{name}: {count}")
    audit_blockers = audit.get("open_blockers", {})
    if not isinstance(audit_blockers, dict):
        raise PackageError("technical audit report has no open_blockers mapping")
    for name, count in sorted(audit_blockers.items()):
        if int(count):
            blockers.append(f"binder audit {name}: {count}")
    counters = readiness.get("counters")
    if not isinstance(counters, dict):
        raise PackageError("production readiness report has no counters mapping")
    for name in MANDATORY_COUNTERS:
        if int(counters.get(name, 0)):
            blockers.append(f"{name}: {counters[name]}")
    gates = readiness.get("gate_statuses")
    if not isinstance(gates, dict):
        raise PackageError("production readiness report has no gate_statuses mapping")
    for name in MANDATORY_GATES:
        if gates.get(name) != "PASS":
            blockers.append(f"{name} gate is {gates.get(name, 'NOT_REPORTED')}")
    if validation.get("validations", {}).get("photos", {}).get("status") != "PASS":
        blockers.append("workbook build validation is missing or did not pass")
    if validation.get("print_validation", {}).get("photos", {}).get("status") != "PASS":
        blockers.append("print validation is missing or did not pass")
    if audit.get("technical_checks") != "PASS":
        blockers.append("technical audit did not pass")
    return not blockers, blockers


def input_checksums(root):
    paths = []
    for directory in ("sources", "data"):
        base = root / directory
        if base.exists():
            if base.is_symlink():
                raise PackageError(f"refusing symlinked input directory: {directory}")
            for path in base.rglob("*"):
                if path.is_symlink():
                    raise PackageError(f"refusing symlinked input: {path.relative_to(root)}")
                if path.is_file():
                    paths.append(path)
    paths.extend(root / path for path in ("requirements.txt", *GENERATOR_FILES)
                 if (root / path).is_file())
    checksums = {}
    for path in sorted(paths):
        if path.is_symlink():
            raise PackageError(f"refusing symlinked input: {path.relative_to(root)}")
        checksums[path.relative_to(root).as_posix()] = sha256(path)
    return checksums


def generator_commit(root, requested):
    if requested:
        return requested
    if os.environ.get("GITHUB_SHA"):
        return os.environ["GITHUB_SHA"]
    try:
        return subprocess.run(
            ["git", "-C", str(root), "rev-parse", "HEAD"],
            check=True, capture_output=True, text=True,
        ).stdout.strip()
    except (OSError, subprocess.CalledProcessError):
        return "UNKNOWN"


def dependency_versions(root):
    versions = {"python": os.sys.version.split()[0]}
    requirements = root / "requirements.txt"
    if not requirements.is_file():
        return versions
    for line in requirements.read_text(encoding="utf-8").splitlines():
        match = re.match(r"\s*([A-Za-z0-9_.-]+)==", line)
        if not match:
            continue
        name = match.group(1)
        try:
            versions[name] = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            versions[name] = "NOT_INSTALLED"
    try:
        versions["libreoffice"] = subprocess.run(
            ["soffice", "--version"], check=True, capture_output=True, text=True,
        ).stdout.strip()
    except (OSError, subprocess.CalledProcessError):
        versions["libreoffice"] = "NOT_INSTALLED"
    for package in ("libreoffice-calc", "libreoffice-writer",
                    "fonts-dejavu-core", "fonts-dejavu-extra"):
        try:
            versions[package] = subprocess.run(
                ["dpkg-query", "-W", "-f=${Version}", package],
                check=True, capture_output=True, text=True,
            ).stdout.strip()
        except (OSError, subprocess.CalledProcessError):
            versions[package] = "NOT_INSTALLED"
    return versions


def approval_basis_counts(csv_path):
    counts = {}
    with csv_path.open(encoding="utf-8-sig", newline="") as stream:
        for row in csv.DictReader(stream):
            basis = row.get("image_approval_basis", "").strip()
            if basis:
                counts[basis] = counts.get(basis, 0) + 1
    return dict(sorted(counts.items()))


def build_manifest(root, build_dir, mode, readiness, validation, audit, gate_pass, gate_reasons, commit):
    approvals = approval_basis_counts(build_dir / "reconciliation/PICKLIST_SYNC.csv")
    files = sorted(
        path for path in build_dir.rglob("*")
        if path.is_file() and path.relative_to(build_dir).as_posix() != "RELEASE_MANIFEST.json"
    )
    if any(path.is_symlink() for path in files):
        raise PackageError("build contains symlinked files; package was not created")
    return {
        "schema_version": 1,
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "package_mode": mode,
        "publication_status": (
            "BLOCKED" if not gate_pass
            else "OWNER_AUTHORIZATION_REQUIRED" if mode == "review"
            else "REQUIRES_PROTECTED_ENVIRONMENT_APPROVAL"
        ),
        "production_gate_pass": gate_pass,
        "production_gate_findings": gate_reasons,
        "generator": {
            "version": PACKAGER_VERSION,
            "repository_commit": commit,
            "tool_sha256": {
                name: sha256(root / name) for name in GENERATOR_FILES if (root / name).is_file()
            },
        },
        "dependency_versions": dependency_versions(root),
        "input_sha256": input_checksums(root),
        "validation": {
            "build_status": validation.get("status", "NOT_REPORTED"),
            "production_readiness_status": readiness.get("status", "NOT_REPORTED"),
            "review_build_publication_status": validation.get(
                "production_publication", "NOT_REPORTED"),
            "release_blockers": readiness.get("release_blockers", {}),
            "owner_review_findings": readiness.get("owner_review_findings", {}),
            "independent_counters": readiness.get("counters", {}),
            "gate_statuses": readiness.get("gate_statuses", {}),
            "technical_audit": audit.get("technical_checks", "NOT_REPORTED"),
            "technical_audit_failures": audit.get("failures", []),
            "build_validations": validation.get("validations", {}),
            "print_validation": validation.get("print_validation", {}),
        },
        "approval_metadata_available": {
            "image_approval_basis_counts": approvals,
            "source": "image_approval_basis values in reconciliation/PICKLIST_SYNC.csv",
            "note": "Recorded basis metadata is evidence only; this manifest does not grant approval.",
        },
        "output_sha256": {
            path.relative_to(build_dir).as_posix(): sha256(path) for path in files
        },
    }


def create_package(build_dir, archive_path, root, mode="review", commit=""):
    root = Path(root).resolve()
    build_dir = Path(build_dir).resolve()
    archive_path = Path(archive_path).resolve()
    if mode not in {"review", "production"}:
        raise PackageError(f"unsupported package mode: {mode}")
    if not build_dir.is_dir():
        raise PackageError(f"build directory does not exist: {build_dir}")
    if archive_path.is_relative_to(build_dir):
        raise PackageError("package ZIP must be outside the build directory")
    if any(path.is_symlink() for path in build_dir.rglob("*")):
        raise PackageError("build contains symlinked paths; package was not created")
    missing = [name for name in REQUIRED_FILES if not (build_dir / name).is_file()]
    if missing:
        raise PackageError("build is missing required deliverables: " + ", ".join(missing))

    readiness_path = build_dir / "reconciliation/PRODUCTION_READINESS.json"
    validation = read_json(build_dir / "VALIDATION.json", "build validation")
    audit = read_json(build_dir / "AUDIT_REPORT.json", "technical audit")
    readiness = read_json(readiness_path, "production readiness")
    gate_pass, gate_reasons = production_gate(readiness, validation, audit)
    if mode == "production" and not gate_pass:
        raise PackageError("production package blocked: " + "; ".join(gate_reasons))
    if archive_path.exists():
        raise PackageError(f"refusing to overwrite existing package: {archive_path}")
    manifest = build_manifest(
        root, build_dir, mode, readiness, validation, audit, gate_pass, gate_reasons,
        generator_commit(root, commit))
    archive_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="plu-package-") as temp_name:
        stage = Path(temp_name) / "package"
        shutil.copytree(build_dir, stage)
        (stage / "RELEASE_MANIFEST.json").write_text(
            json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        temporary_fd, temporary_name = tempfile.mkstemp(
            prefix=f".{archive_path.name}.", suffix=".tmp", dir=archive_path.parent)
        os.close(temporary_fd)
        temporary_archive = Path(temporary_name)
        try:
            with zipfile.ZipFile(temporary_archive, "w", compression=zipfile.ZIP_DEFLATED,
                                 compresslevel=6) as archive:
                for path in sorted(stage.rglob("*")):
                    if path.is_file():
                        archive.write(path, path.relative_to(stage).as_posix())
            with zipfile.ZipFile(temporary_archive) as archive:
                bad_member = archive.testzip()
                if bad_member:
                    raise PackageError(f"package ZIP integrity failed at {bad_member}")
            os.link(temporary_archive, archive_path)
        finally:
            temporary_archive.unlink(missing_ok=True)
    return manifest


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--build-dir", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True, help="new ZIP path; existing files are never replaced")
    parser.add_argument("--mode", choices=("review", "production"), default="review")
    parser.add_argument("--commit", default="")
    args = parser.parse_args(argv)
    try:
        manifest = create_package(args.build_dir, args.output, args.root, args.mode, args.commit)
    except (OSError, PackageError, zipfile.BadZipFile) as error:
        parser.exit(1, f"ERROR: {error}\n")
    print(json.dumps({
        "package": str(args.output),
        "package_mode": manifest["package_mode"],
        "production_gate_pass": manifest["production_gate_pass"],
        "publication_status": manifest["publication_status"],
        "production_gate_findings": manifest["production_gate_findings"],
    }, indent=2))


if __name__ == "__main__":
    main()
