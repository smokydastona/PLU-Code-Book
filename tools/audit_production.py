"""Audit legacy approvals and current release blockers without altering source evidence."""
import argparse
import csv
import hashlib
import json
from collections import Counter
from pathlib import Path

from PIL import Image, ImageOps

from check_image_uniqueness import audit_entries
from organic_watermarks import is_organic, read_csv, verify_booklets
from release_policy import explicit_image_rejection, known_permission_restriction
from sync_picklist import reconcile as reconcile_picklist

ROOT = Path(__file__).resolve().parents[1]


def image_findings(row, review, root):
    path = root / "data" / row["image"]
    errors = []
    with Image.open(path) as image:
        image.load()
        width, height = ImageOps.exif_transpose(image).size
    if min(width, height) < 500:
        errors.append("IMAGE_QUALITY:BELOW_500_PX")
    if (str(width), str(height)) != (row["width"], row["height"]):
        errors.append("PROVENANCE:RECORDED_DIMENSIONS_MISMATCH")
    if review["verdict"].strip().casefold() == "rejected":
        errors.append("EXPLICIT_IMAGE_REJECTION")
    elif review["verdict"] != "Approved":
        errors.append("HISTORICAL_VISUAL_REVIEW:" + review["verdict"])
    if row["source"] == "AI Generated":
        errors.append("PROVENANCE:AI_REFERENCE_NOT_PHOTOGRAPH")
    sidecar = path.with_suffix(path.suffix + ".json")
    if not sidecar.exists():
        errors.append("PROVENANCE:MISSING_IMAGE_METADATA")
    else:
        metadata = json.loads(sidecar.read_text(encoding="utf-8"))
        if not isinstance(metadata, dict):
            raise ValueError(f"{sidecar}: expected an object")
        if metadata.get("sha256") != hashlib.sha256(path.read_bytes()).hexdigest():
            errors.append("PROVENANCE:CHECKSUM_MISMATCH")
        if (metadata.get("plu"), metadata.get("product")) != (row["plu"], row["name"]):
            errors.append("PROVENANCE:ASSIGNMENT_MISMATCH")
        if metadata.get("rights_verification_status") != "VERIFIED":
            errors.append("RIGHTS:VERIFICATION_NOT_RECORDED")
        if known_permission_restriction(dict(
                rights_verification_status=metadata.get("rights_verification_status"),
                rights_status=metadata.get("rights_status"), provenance=metadata)):
            errors.append("RIGHTS:REQUIRED_PERMISSION_NOT_GRANTED")
        if metadata.get("visual_approval_status") != "APPROVED":
            errors.append("PROVENANCE:VISUAL_APPROVAL_NOT_RECORDED")
        owner_permission = metadata.get("rights_status") == "OWNER_CONFIRMED_PERMISSION"
        required = ["creator", "attribution"]
        if owner_permission:
            required.extend(("source_reference", "rights_evidence"))
            if metadata.get("license") != "OWNER_CONFIRMED_PERMISSION":
                errors.append("RIGHTS:OWNER_PERMISSION_LICENCE_CONFLICT")
        else:
            required.extend(("license", "source_page_url", "license_url"))
        for key in required:
            if not isinstance(metadata.get(key), str) or not metadata[key].strip():
                errors.append("PROVENANCE:MISSING_" + key.upper())
    return width, height, errors


def inventory(root):
    rows = read_csv(root / "data/entries.csv")
    if not rows:
        raise ValueError("entry CSV is empty")
    reviews = read_csv(root / "data/image_review.csv")
    by_idx = {row["idx"]: row for row in reviews}
    if len(by_idx) != len(reviews) or set(by_idx) != {row["idx"] for row in rows}:
        raise ValueError("review and entry IDs must match exactly, without duplicates")
    verify_booklets(root, rows)
    groups = audit_entries(rows, root / "data")
    duplicate_ids = {row["idx"] for group in groups for row in group}
    result = []
    for row in rows:
        review = by_idx[row["idx"]]
        if any(row[key] != review[key] for key in ("book", "plu", "name", "image")):
            raise ValueError(f"entry {row['idx']}: review assignment mismatch")
        width, height, errors = image_findings(row, review, root)
        rejected = explicit_image_rejection({}, {"historical_review": review})
        if row["idx"] in duplicate_ids:
            errors.append("DUPLICATE_IMAGE_REUSE_REVIEW")
        result.append(dict(
            idx=row["idx"], book=row["book"], plu=row["plu"], product=row["name"],
            image="data/" + row["image"], width=width, height=height,
            organic=is_organic(row), approval_status="NOT_APPROVED" if rejected else "APPROVED_LEGACY",
            approval_basis="" if rejected else "OWNER_LEGACY_PREAPPROVED_POLICY",
            visual_verdict=review["verdict"], historical_visual_verdict=review["verdict"],
            historical_findings="; ".join(errors),
            blockers="EXPLICIT_IMAGE_REJECTION" if rejected else "NONE",
        ))
    findings = Counter(error for row in result
                       for error in row["historical_findings"].split("; ") if error)
    summary = dict(
        status="LEGACY_APPROVALS_WITH_EXPLICIT_REJECTIONS" if any(
            row["blockers"] != "NONE" for row in result) else "ALL_LEGACY_ENTRIES_APPROVED",
        entries=len(rows), organic_entries=sum(is_organic(row) for row in rows),
        duplicate_image_groups=len(groups),
        entries_with_release_blockers=sum(row["blockers"] != "NONE" for row in result),
        owner_review_finding_counts=dict(sorted(findings.items())),
        approval_basis="Existing legacy policy applies only to exact legacy entry/product/image "
                       "matches. Historical findings remain separate; explicit rejection overrides it.",
        scope="Legacy image approval only. Non-produce remains a code-only supplement.",
    )
    return result, summary


def combined_summary(picklist_summary, image_summary=None):
    status = (
        "BLOCKED"
        if picklist_summary["status"] != "READY_TO_BUILD"
        else "AUTOMATED_CHECKS_PASSED_REQUIRES_RELEASE_REVIEW"
    )
    result = dict(
        status=status, picklist_sync=picklist_summary,
        release_blockers=picklist_summary.get("release_blocker_counts", {}),
        owner_review_findings=picklist_summary.get("owner_review_finding_counts", {}),
        counters=picklist_summary.get("counters", {}),
        counter_definitions=picklist_summary.get("counter_definitions", {}),
        gate_statuses=dict(
            build="NOT_RUN",
            technical_audit="NOT_RUN",
            data_reconciliation="BLOCKED" if any(
                picklist_summary.get("counters", {}).get(key, 0)
                for key in ("unresolved_source_record_reconciliation", "unresolved_product_identity")
            ) else "PASS",
            image_selection="BLOCKED" if picklist_summary.get(
                "counters", {}).get("unresolved_image_selection", 0) else "PASS",
            image_quality="WARNING" if picklist_summary.get(
                "counters", {}).get("image_quality_warnings", 0) else "PASS",
            rights_permission="OWNER_REVIEW_REQUIRED" if picklist_summary.get(
                "counters", {}).get("rights_permission_findings", 0) else "PASS",
            provenance="DOCUMENTATION_GAPS" if picklist_summary.get(
                "counters", {}).get("provenance_documentation_gaps", 0) else "PASS",
            duplicate_reuse="OWNER_REVIEW_REQUIRED" if picklist_summary.get(
                "counters", {}).get("duplicate_image_reuse_reviews", 0) else "PASS",
            classification_placement="BLOCKED" if picklist_summary.get(
                "counters", {}).get("classification_section_reviews", 0) else "PASS",
            production_release="BLOCKED" if status == "BLOCKED"
            else "OWNER_AUTHORIZATION_REQUIRED",
        ),
        scope="Current WorkingList reconciliation and image readiness. Technical success and "
              "non-blocking owner-review findings do not authorize production publication.",
    )
    if image_summary is not None:
        result["legacy_image_audit"] = image_summary
    return result


def write_release_blockers(rows, picklist_summary, output):
    report_rows = [
        dict(source_id=row["source_id"], source_ids=";".join(row["source_ids"]),
             plu=row["plu"], product=row["name"],
             blockers="; ".join(row["blockers"]))
        for row in rows if row["blockers"]
    ]
    reconciliation_count = picklist_summary.get("working_list_review_records", 0)
    if reconciliation_count or picklist_summary.get("reconciliation_review_decisions", 0):
        report_rows.append(dict(
            source_id="", source_ids="", plu="", product="WorkingList source reconciliation",
            blockers=f"WORKING_LIST_RECONCILIATION: {reconciliation_count} records require review; "
                     f"{picklist_summary.get('reconciliation_review_decisions', 0)} decision(s) unresolved"))
    with (output / "PRODUCTION_BLOCKERS.csv").open(
            "w", newline="", encoding="utf-8-sig") as file:
        fields = ["source_id", "source_ids", "plu", "product", "blockers"]
        writer = csv.DictWriter(file, fieldnames=fields, lineterminator="\n")
        writer.writeheader()
        writer.writerows(report_rows)
    return report_rows


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--output", type=Path, help="Report directory")
    parser.add_argument("--legacy", action="store_true",
                        help="Also inventory historical booklet images; not a production gate")
    args = parser.parse_args()
    try:
        output = args.output or args.root / "outputs/reports"
        output.mkdir(parents=True, exist_ok=True)
        rows, _, _, picklist_summary = reconcile_picklist(args.root, output)
        if hashlib.sha256((args.root / "sources/plu/PickList.xml").read_bytes()).hexdigest() != picklist_summary["source_sha256"]:
            raise ValueError("source changed during reconciliation; discard output and retry")
        release_blockers = write_release_blockers(rows, picklist_summary, output)
        image_summary = None
        if args.legacy:
            rows, image_summary = inventory(args.root)
            with (output / "LEGACY_IMAGE_INVENTORY.csv").open(
                    "w", newline="", encoding="utf-8-sig") as file:
                writer = csv.DictWriter(file, fieldnames=list(rows[0]), lineterminator="\n")
                writer.writeheader()
                writer.writerows(rows)
        if hashlib.sha256((args.root / "sources/plu/PickList.xml").read_bytes()).hexdigest() != picklist_summary["source_sha256"]:
            raise ValueError("source changed during audit; discard output and retry")
        summary = combined_summary(picklist_summary, image_summary)
        summary["release_blocker_rows"] = len(release_blockers)
        (output / "PRODUCTION_READINESS.json").write_text(
            json.dumps(summary, indent=2) + "\n", encoding="utf-8")
    except (OSError, ValueError, KeyError) as exc:
        parser.exit(1, f"ERROR: {exc}\n")
    print(json.dumps(summary, indent=2))
    raise SystemExit(1 if summary["status"] == "BLOCKED" else 0)


if __name__ == "__main__":
    main()
