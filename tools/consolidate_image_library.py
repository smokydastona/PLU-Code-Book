"""Store identical library JPEGs once while retaining every logical asset sidecar."""
import argparse
import csv
import hashlib
import json
from collections import defaultdict
from pathlib import Path

from library_assets import ALIASES, check_aliases, resolve_image
from used_image_library import check_used_library, save_if_changed

ROOT = Path(__file__).resolve().parents[1]


def consolidate(root):
    root = Path(root)
    library = root / "sources/image-library"
    check_used_library(root)
    previous = check_aliases(library)
    logical_names = {path.name for path in library.glob("*.jpg")} | set(previous)
    groups = defaultdict(list)
    contents = {}
    for name in sorted(logical_names):
        path = resolve_image(library / name, previous)
        content = path.read_bytes()
        checksum = hashlib.sha256(content).hexdigest()
        metadata = library / (name + ".json")
        if metadata.is_symlink() or not metadata.is_file():
            raise ValueError(f"library metadata missing: {name}")
        if json.loads(metadata.read_text())["sha256"] != checksum:
            raise ValueError(f"library metadata checksum mismatch: {name}")
        groups[checksum].append(name)
        contents[name] = content
    aliases, duplicates, report = {}, [], []
    for checksum, names in sorted(groups.items()):
        def rank(name):
            metadata = json.loads((library / (name + ".json")).read_text())
            return (not bool(metadata.get("owner_candidate_id")),
                    "rendering" not in metadata, name)

        existing = [name for name in names if (library / name).is_file()]
        canonical = min(existing, key=rank)
        for name in names:
            if contents[name] != contents[canonical]:
                raise ValueError("SHA256 collision: refusing to consolidate different images")
            if name == canonical:
                continue
            aliases[name] = dict(image=canonical, sha256=checksum)
            report.append(dict(logical_image=name, stored_image=canonical,
                               sha256=checksum, metadata=name + ".json",
                               approval_effect="NONE"))
            if (library / name).is_file():
                duplicates.append(library / name)
    removed_bytes = sum(path.stat().st_size for path in duplicates)
    save_if_changed(library / ALIASES,
                    (json.dumps(aliases, indent=2, sort_keys=True) + "\n").encode())
    for path in duplicates:
        path.unlink()
    check_used_library(root)
    check_aliases(library)
    result = dict(logical_images=len(logical_names), stored_images=len(groups),
                  aliases=len(aliases), removed_files=len(duplicates),
                  removed_bytes=removed_bytes,
                  duplicate_bytes_avoided=sum(len(contents[name]) for name in aliases),
                  metadata_records_preserved=len(logical_names),
                  approval_effect="NONE", original_sources="UNCHANGED")
    reports = root / "outputs/reports"
    reports.mkdir(parents=True, exist_ok=True)
    with (reports / "IMAGE_LIBRARY_ALIASES.csv").open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=[
            "logical_image", "stored_image", "sha256", "metadata", "approval_effect"],
            lineterminator="\n")
        writer.writeheader()
        writer.writerows(report)
    save_if_changed(reports / "IMAGE_LIBRARY_CONSOLIDATION.json",
                    (json.dumps(result, indent=2) + "\n").encode())
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args()
    try:
        result = consolidate(args.root)
    except (OSError, ValueError, KeyError) as exc:
        parser.exit(1, f"ERROR: {exc}\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
