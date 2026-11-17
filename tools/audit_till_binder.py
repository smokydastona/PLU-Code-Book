"""Post-build audit of the complete till binder: package, print layout, images and blockers.

Reads a finished build directory (workbook, PDF, VALIDATION.json and CSV reports) plus the
untouched original workbook, writes audit reports into the same directory, and exits non-zero
when a technical check fails. A passing audit never approves operational use.
"""
import argparse
import csv
import hashlib
import json
import sys
import zipfile
from collections import Counter, defaultdict
from pathlib import Path

import pymupdf
from openpyxl import load_workbook

from impose_binder_booklet import (MAP_NAME as BOOKLET_MAP, OUTPUT_NAME as BOOKLET_PDF,
                                   ROTATED_PAGES as BOOKLET_ROTATED_SECTIONS, audit_imposed_booklet)

WORKBOOK = "Complete Till Binder - Updated - Photos.xlsx"
PDF = "Complete Till Binder - Updated - Photos.pdf"
MARGIN_POINTS = 14
MIN_IMAGE_EDGE = 200
BLOCKER_ACTIONS = {
    "APPROVED_IMAGE_REQUIRED": "Owner supplies or approves an exact-identity image for each listed PLU (MISSING_IMAGES.csv).",
    "PRODUCT_IMAGE_REVIEW_REQUIRED": "Fresh product-specific image review after identity change (MISSING_IMAGES.csv).",
    "PROHIBITED_IMAGE_REUSE": "Owner resolves image reuse between different product identities (reconciliation/IMAGE_REQUESTS.csv).",
    "CLASSIFICATION_REQUIRED": "Owner classifies the listed products (PLACEMENT_REVIEW.csv).",
    "SECTION_REQUIRED": "Owner assigns the listed products to a display section (PLACEMENT_REVIEW.csv).",
}


def sha(content):
    return hashlib.sha256(content).hexdigest()


def read_csv(path):
    with Path(path).open(encoding="utf-8-sig", newline="") as file:
        return list(csv.DictReader(file))


def write_csv(path, rows, fields):
    with Path(path).open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fields, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def audit_package(original_path, built_path):
    with zipfile.ZipFile(original_path) as before, zipfile.ZipFile(built_path) as after:
        old = {name: sha(before.read(name)) for name in before.namelist()}
        new = {name: sha(after.read(name)) for name in after.namelist()}
        corrupt = after.testzip()
    unchanged = sorted(name for name in old if new.get(name) == old[name])
    changed = sorted(name for name in old if name in new and new[name] != old[name])
    removed = sorted(name for name in old if name not in new)
    added = sorted(name for name in new if name not in old)
    original_media = [name for name in old if "/media/" in name]
    sheets_before = load_workbook(original_path, read_only=True).sheetnames
    sheets_after = load_workbook(built_path, read_only=True).sheetnames
    return dict(
        zip_integrity=corrupt is None,
        original_parts=len(old), built_parts=len(new),
        unchanged_parts=len(unchanged), changed_parts=changed, removed_parts=removed,
        added_part_count=len(added),
        original_media_parts=len(original_media),
        original_media_byte_identical=sum(new.get(name) == old[name] for name in original_media),
        original_media_removed=[name for name in original_media if name not in new],
        original_media_changed=[name for name in original_media if name in new and new[name] != old[name]],
        sheets_original=sheets_before, sheets_built=sheets_after,
        sheets_removed=[name for name in sheets_before if name not in sheets_after],
    )


def section_plan(validation):
    produce = validation["planned_produce_print_pages"]
    return ([("Time To Stretch", 1, (612, 792)), ("Conversions", 1, (612, 792)),
             ("Weighed Items Over $100", 2, (612, 792)), ("Misc Codes", 1, (612, 792)),
             ("Soil", 1, (612, 792)), ("Bakery", 1, (612, 792)), ("Deli", 1, (612, 792)),
             ("Produce", produce, (792, 612)), ("Bulk", 6, (792, 612)),
             ("Delivery Charges", 2, (612, 792)), ("Store Directory", 2, (792, 612))])


def audit_pdf(pdf_path, validation, output):
    plan = section_plan(validation)
    document = pymupdf.open(pdf_path)
    expected_pages = sum(count for _, count, _ in plan)
    findings = []
    if len(document) != expected_pages:
        findings.append(f"page count {len(document)} differs from the expected {expected_pages}")
    labels = [(name, size) for name, count, size in plan for _ in range(count)]
    rows = []
    for number, page in enumerate(document, 1):
        name, size = labels[number - 1] if number <= len(labels) else ("UNEXPECTED", None)
        width, height = round(page.rect.width), round(page.rect.height)
        pixmap = page.get_pixmap(dpi=40, colorspace=pymupdf.csGRAY)
        inked = sum(1 for value in pixmap.samples if value < 240) / len(pixmap.samples)
        words = page.get_text("words")
        spill = [w for w in words if w[0] < MARGIN_POINTS - 8 or w[1] < MARGIN_POINTS - 8
                 or w[2] > width - MARGIN_POINTS + 8 or w[3] > height - MARGIN_POINTS + 8]
        problems = []
        if size and (width, height) != size:
            problems.append(f"size {width}x{height}, expected {size[0]}x{size[1]}")
        if inked < 0.005:
            problems.append("page is effectively blank")
        if spill:
            problems.append(f"{len(spill)} words within the print-margin guard")
        if name not in {"Time To Stretch", "Conversions", "Weighed Items Over $100", "Delivery Charges",
                        "Store Directory", "Soil"} and not words:
            problems.append("no extractable text")
        rows.append(dict(page=number, section=name, width_pt=width, height_pt=height,
                         orientation="landscape" if width > height else "portrait",
                         words=len(words), ink_percent=round(inked * 100, 2),
                         images=len(page.get_images()), flags="; ".join(problems) or "none"))
        findings.extend(f"page {number} ({name}): {problem}" for problem in problems)
    write_csv(output / "PRINT_LAYOUT_REPORT.csv", rows, list(rows[0]))
    return dict(pages=len(document), expected_pages=expected_pages,
                sections={name: count for name, count, _ in plan}, findings=findings), rows


def audit_images(output):
    mapping = read_csv(output / "PHOTO_MAPPING.csv")
    missing = read_csv(output / "MISSING_IMAGES.csv")
    by_hash = defaultdict(list)
    for row in mapping:
        if row["embedded"] == "True":
            by_hash[row["embedded_sha256"]].append(row)
    shared = {digest: rows for digest, rows in by_hash.items()
              if len({row["plu"] for row in rows}) > 1}
    small = [row for row in mapping if row["embedded"] == "True" and
             min(int(row["width"] or 0), int(row["height"] or 0)) < MIN_IMAGE_EDGE]
    return dict(
        records=len(mapping), embedded=sum(row["embedded"] == "True" for row in mapping),
        missing_or_uncertain=len(missing),
        status_counts=dict(Counter(row["image_status"] for row in mapping)),
        sources=dict(Counter(row["image_source"] for row in mapping if row["embedded"] == "True")),
        organic_badge_embedded=sum(row["organic_badge"] == "True" for row in mapping),
        images_shared_by_multiple_plus=len(shared),
        embedded_below_200px=len(small),
        smallest_edge=min((min(int(r["width"]), int(r["height"])) for r in mapping if r["embedded"] == "True"),
                          default=0),
    ), small, shared


def audit_oldtimers(output, validation):
    info = validation["oldtimers_produce"]
    path = output / info["file"]
    if not path.is_file() or sha(path.read_bytes()) != info["sha256"]:
        raise ValueError("OldTimers PDF is missing or differs from the verified build")
    if info["records"] != len(read_csv(output / "PRODUCE_LAYOUT.csv")):
        raise ValueError("OldTimers record count differs from the binder produce layout")
    with pymupdf.open(path) as document:
        if len(document) != info["pages"] or any(page.get_images() for page in document):
            raise ValueError("OldTimers PDF page count changed or contains images")
        if any(span["size"] < 12 for page in document
               for block in page.get_text("dict")["blocks"]
               for line in block.get("lines", []) for span in line["spans"]):
            raise ValueError("OldTimers text must be at least 12 points")
    return info


def write_blockers(validation, output):
    reconciliation = validation["source_reconciliation"]
    rows = [dict(blocker=name, records=count, required_owner_action=BLOCKER_ACTIONS.get(
        name, "Review the reconciliation reports."), status="OPEN")
        for name, count in sorted(reconciliation["blocker_counts"].items())]
    rows.append(dict(blocker="WORKING_LIST_RECONCILIATION", records=reconciliation["working_list_review_records"],
                     required_owner_action="Owner records reviewed decisions for each flagged WorkingList record "
                     "(reconciliation/WORKING_LIST_RECONCILIATION.csv); decisions recorded: "
                     f"{reconciliation['reconciliation_review_decisions']}.", status="OPEN"))
    rows.append(dict(blocker="ORGANIC_PLACEMENT", records=reconciliation["organic_placements_needing_review"],
                     required_owner_action="Owner confirms organic counterpart placement (PLACEMENT_REVIEW.csv).",
                     status="OPEN"))
    write_csv(output / "PRODUCTION_BLOCKERS.csv", rows, ["blocker", "records", "required_owner_action", "status"])
    return rows


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True, type=Path, help="finished build directory")
    parser.add_argument("--binder", required=True, type=Path, help="untouched original workbook")
    args = parser.parse_args(argv)
    output = args.output
    validation = json.loads((output / "VALIDATION.json").read_text(encoding="utf-8"))
    original_sha = sha(args.binder.read_bytes())
    failures = []
    if original_sha != validation["original_binder_sha256"]:
        failures.append("original workbook checksum differs from the build record")
    package = audit_package(args.binder, output / WORKBOOK)
    if not package["zip_integrity"]:
        failures.append("workbook package is corrupt")
    if package["original_media_removed"] or package["original_media_changed"]:
        failures.append("original media parts were removed or changed")
    layout, layout_rows = audit_pdf(output / PDF, validation, output)
    failures.extend(layout["findings"])
    booklet = {}
    try:
        booklet = audit_imposed_booklet(
            output / PDF, output / BOOKLET_PDF, output / BOOKLET_MAP,
            rotate_pages={int(row["page"]) for row in layout_rows
                          if row["section"] in BOOKLET_ROTATED_SECTIONS})
    except (OSError, ValueError) as error:
        failures.append(f"binder booklet imposition: {error}")
    images, small, shared = audit_images(output)
    oldtimers = {}
    try:
        oldtimers = audit_oldtimers(output, validation)
    except (OSError, ValueError, KeyError) as error:
        failures.append(f"OldTimers produce list: {error}")
    blockers = write_blockers(validation, output)
    if validation.get("production_publication") != "BLOCKED" and blockers:
        failures.append("open blockers exist but production publication is not BLOCKED")
    report = dict(
        technical_checks="PASS" if not failures else "FAIL", failures=failures,
        operational_status="REVIEW ONLY — NOT APPROVED FOR OPERATIONAL USE" if blockers
        else "TECHNICAL CHECKS ONLY",
        original_binder_sha256=original_sha, package=package, print_layout=layout,
        booklet_imposition=booklet, images=images, oldtimers_produce=oldtimers,
        open_blockers={row["blocker"]: row["records"] for row in blockers})
    (output / "AUDIT_REPORT.json").write_text(json.dumps(report, indent=2, ensure_ascii=False) + "\n",
                                              encoding="utf-8")
    validation["till_binder_audit"] = dict(
        technical_checks=report["technical_checks"], operational_status=report["operational_status"],
        report="AUDIT_REPORT.json")
    (output / "VALIDATION.json").write_text(json.dumps(validation, indent=2, ensure_ascii=False) + "\n",
                                            encoding="utf-8")
    print(json.dumps({key: report[key] for key in ("technical_checks", "operational_status", "failures")}, indent=2))
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
