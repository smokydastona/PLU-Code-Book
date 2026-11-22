"""Refresh the single temporary produce review using exact screened candidate bytes."""
import argparse
import csv
import hashlib
import json
import math
import tempfile
from pathlib import Path

import pymupdf
from PIL import Image

from library_assets import resolve_image
from release_policy import explicit_image_rejection
from used_image_library import check_used_library, save_if_changed
from working_list import read_working_list

ROOT = Path(__file__).resolve().parents[1]
PDF_NAME = "Produce_Image_Review_Temporary.pdf"
INDEX_NAME = "TEMPORARY_PRODUCE_IMAGE_REVIEW.csv"
REPORT_NAME = "TEMPORARY_PRODUCE_IMAGE_REVIEW.json"
SCREENING_NAME = "PRODUCE_IMAGE_AUTOMATIC_REVIEW.csv"


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def candidates(root, output):
    items, source_sha = read_working_list(root)
    current = {row["source_id"]: row for row in items if row["included"]}
    assignments = {row["source_id"]: row for row in json.loads(
        (root / "data/picklist_assignments.json").read_text())["assignments"]}
    check_used_library(root)
    library = root / "sources/image-library"
    by_sha = {}
    for path in sorted(library.glob("*.jpg")):
        by_sha.setdefault(sha(path), path)
    with (output / SCREENING_NAME).open(encoding="utf-8-sig", newline="") as file:
        screening = list(csv.DictReader(file))
    pending, skipped, seen = [], [], set()
    for row in screening:
        identity = row["source_id"]
        item = current.get(identity)
        if (identity in seen or not item or item.get("source_record_changed")
                or (item["plu"], item["name"]) != (row["plu"], row["product"])
                or set(item["source_ids"]) != set(row["source_identity_ids"].split(";"))):
            raise ValueError(f"stale or ambiguous screening identity: {row['plu']} / {row['product']}")
        seen.add(identity)
        declared = root / row["candidate_path"]
        if not declared.resolve().is_relative_to(root.resolve()):
            raise ValueError(f"candidate escapes repository: {declared}")
        path = resolve_image(declared)
        expected_sha = row["candidate_sha256"]
        if not path.is_file() or sha(path) != expected_sha:
            path = by_sha.get(expected_sha)
            if path is None:
                raise ValueError(f"exact screened candidate missing: {row['plu']} / {row['product']}")
        with Image.open(path) as image:
            image.load()
            if image.size != (int(row["width"]), int(row["height"])):
                raise ValueError(f"screened candidate dimensions changed: {path}")
        record = dict(row, resolved_candidate_path=str(path.relative_to(root)),
                      candidate_path_resolution=("UNCHANGED" if declared == path else
                                                 "EXACT_SHA256_RELOCATED"),
                      candidate_path_current_sha256=sha(resolve_image(declared))
                      if resolve_image(declared).is_file() else "MISSING",
                      working_record_sha256=item["record_sha256"])
        if row["owner_image_approval"] == "OWNER_APPROVED_EXACT_IMAGE_FOR_PRODUCTION":
            evidence = root / row["owner_review_evidence"]
            assignment = assignments.get(identity, {})
            if (not evidence.is_file() or sha(evidence) != row["owner_review_pdf_sha256"]
                    or row["owner_review_picklist_sha256"] != source_sha
                    or assignment.get("visual_approval_status") != "APPROVED"
                    or explicit_image_rejection(assignment)
                    or (assignment.get("reviewed_plu"), assignment.get("reviewed_product")) !=
                    (row["plu"], row["product"])):
                raise ValueError(f"prior approval evidence needs review: {row['plu']}")
            skipped.append(dict(record, disposition="ALREADY_OWNER_APPROVED_EXACT_CANDIDATE"))
            continue
        record["checkbox_allowed"] = (
            row["visual_screening"] in {"VISUAL_SCREEN_PASS", "EDITED_REVIEW_CANDIDATE"}
            and not explicit_image_rejection(assignments.get(identity, {})))
        record["disposition"] = ("OWNER_REVIEW_REQUIRED" if record["checkbox_allowed"] else
                                 "REPLACEMENT_REQUIRED")
        pending.append(record)
    return pending, skipped, source_sha


def text(page, rect, value, size=10, color=(0, 0, 0)):
    if page.insert_textbox(rect, value, fontsize=size, fontname="review", color=color) < 0:
        raise ValueError(f"review text does not fit: {value}")


def render(path, rows, source_sha):
    document = pymupdf.open()
    pages = max(1, math.ceil(len(rows) / 4))
    try:
        document.set_metadata(dict(
            title="Temporary Produce Image Review", author="PLU Code Book review",
            subject="Exact-image owner review; production publication gates remain separate",
            keywords="picklist-sha256:" + source_sha))
        for number in range(pages):
            page = document.new_page(width=792, height=612)
            page.insert_font(fontname="review", fontbuffer=pymupdf.Font("helv").buffer)
            text(page, (24, 10, 768, 34),
                 "PRODUCE IMAGE REVIEW - TEMPORARY / NOT A PRODUCTION RELEASE", 12)
            text(page, (24, 32, 768, 60),
                 "Checking a box approves only the pictured image for that exact product. "
                 "Red replacement-required entries have no checkbox.", 10)
            for offset, row in enumerate(rows[number * 4:number * 4 + 4]):
                x, y = 24 + (offset % 2) * 378, 66 + (offset // 2) * 252
                page.draw_rect((x, y, x + 366, y + 242), color=(.7, .7, .7))
                text(page, (x + 8, y + 8, x + 324, y + 49),
                     f"PLU {row['plu']}  |  {row['product']}", 11)
                page.insert_image(
                    (x + 8, y + 55, x + 184, y + 231),
                    stream=(path.parent / row["_image_file"]).read_bytes(),
                    keep_proportion=True)
                color = (0, .35, 0) if row["checkbox_allowed"] else (.7, 0, 0)
                label = ("OWNER REVIEW REQUIRED" if row["checkbox_allowed"] else
                         "REPLACEMENT REQUIRED")
                text(page, (x + 196, y + 56, x + 358, y + 87), label, 10, color)
                text(page, (x + 196, y + 90, x + 358, y + 188),
                     row["visual_note"], 8.5)
                text(page, (x + 196, y + 189, x + 358, y + 225),
                     "Rights, provenance, image quality and publication gates remain separate.", 8.5)
                text(page, (x + 196, y + 226, x + 358, y + 240),
                     "SHA-256: " + row["candidate_sha256"][:16] + "...", 7)
                if row["checkbox_allowed"]:
                    widget = pymupdf.Widget()
                    widget.field_name = "approve_" + row["source_id"]
                    widget.field_type = pymupdf.PDF_WIDGET_TYPE_CHECKBOX
                    widget.rect = pymupdf.Rect(x + 338, y + 12, x + 356, y + 30)
                    widget.field_value = "Off"
                    page.add_widget(widget)
                row["review_page"] = number + 1
            if not rows:
                text(page, (24, 80, 768, 120), "No pending screened candidates.", 12)
            text(page, (720, 587, 768, 605), f"{number + 1}/{pages}", 9)
        document.save(path, garbage=4, deflate=True)
    finally:
        document.close()


def verify(path, rows):
    with pymupdf.open(path) as document:
        if len(document) != max(1, math.ceil(len(rows) / 4)):
            raise ValueError("incorrect temporary review page count")
        for number, page in enumerate(document):
            expected = rows[number * 4:number * 4 + 4]
            content = " ".join(page.get_text().split())
            if page.rect != pymupdf.Rect(0, 0, 792, 612):
                raise ValueError("temporary review must use landscape US Letter")
            for row in expected:
                if " ".join(f"PLU {row['plu']} | {row['product']}".split()) not in content:
                    raise ValueError(f"missing or changed review product: {row['plu']}")
            widgets = list(page.widgets())
            if {widget.field_name for widget in widgets} != {
                    "approve_" + row["source_id"] for row in expected if row["checkbox_allowed"]}:
                raise ValueError("incorrect review approval checkbox identities")
            if any(widget.field_value != "Off" for widget in widgets):
                raise ValueError("review checkboxes must start unchecked")
            drawn = page.get_image_info(xrefs=True)
            if len(drawn) != len(expected):
                raise ValueError("incorrect review image count")
            for row, image in zip(expected, drawn):
                embedded = document.extract_image(image["xref"])["image"]
                if hashlib.sha256(embedded).hexdigest() != row["candidate_sha256"]:
                    raise ValueError(f"review image bytes differ: {row['plu']}")


def build(root):
    root = Path(root)
    output = root / "outputs/review-booklets"
    pending, skipped, source_sha = candidates(root, output)
    duplicates = [path for path in output.glob("Produce_Image_Review_Temporary*.pdf")
                  if path.name != PDF_NAME]
    if any(path.is_symlink() or not path.is_file() for path in duplicates):
        raise ValueError("temporary review duplicates must be regular files")
    with tempfile.TemporaryDirectory(prefix="produce-review-", dir=output.parent) as folder:
        staged = Path(folder)
        for number, row in enumerate(pending):
            row["_image_file"] = f"candidate-{number}.jpg"
            (staged / row["_image_file"]).write_bytes((root / row["resolved_candidate_path"]).read_bytes())
        render(staged / PDF_NAME, pending, source_sha)
        verify(staged / PDF_NAME, pending)
        for row in pending:
            row.pop("_image_file")
        report = dict(
            status="TEMPORARY_REVIEW_NOT_A_PRODUCTION_RELEASE",
            pdf=PDF_NAME, pdf_sha256=sha(staged / PDF_NAME), source_sha256=source_sha,
            screening_sha256=sha(output / SCREENING_NAME),
            entries=len(pending), pages=max(1, math.ceil(len(pending) / 4)),
            approval_checkboxes=sum(row["checkbox_allowed"] for row in pending),
            replacement_required=sum(not row["checkbox_allowed"] for row in pending),
            excluded_prior_approval_count=len(skipped),
            excluded_prior_approvals=skipped,
            production_publication="NOT_AUTHORIZED",
            validation="EXACT_PRODUCT_TEXT_IMAGE_BYTES_AND_UNCHECKED_IDENTITY_LINKED_WIDGETS",
        )
        with (staged / INDEX_NAME).open("w", newline="", encoding="utf-8-sig") as file:
            fields = list(pending[0]) if pending else ["source_id", "plu", "product", "disposition"]
            writer = csv.DictWriter(file, fieldnames=fields, lineterminator="\n")
            writer.writeheader()
            writer.writerows(pending)
        (staged / REPORT_NAME).write_text(json.dumps(report, indent=2) + "\n")
        for name in (PDF_NAME, INDEX_NAME, REPORT_NAME):
            save_if_changed(output / name, (staged / name).read_bytes())
        for duplicate in duplicates:
            duplicate.unlink()
    checksums = output / "SHA256SUMS.txt"
    if checksums.exists():
        lines = checksums.read_text().splitlines()
        names = {PDF_NAME, INDEX_NAME, REPORT_NAME} | {path.name for path in duplicates}
        lines = [line for line in lines if line.split("  ", 1)[-1] not in names]
        lines.extend(f"{sha(output / name)}  {name}" for name in (PDF_NAME, INDEX_NAME, REPORT_NAME))
        save_if_changed(checksums, ("\n".join(lines) + "\n").encode())
    return {key: value for key, value in report.items() if key != "excluded_prior_approvals"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args()
    try:
        result = build(args.root)
    except (OSError, ValueError, KeyError) as exc:
        parser.exit(1, f"ERROR: {exc}\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
