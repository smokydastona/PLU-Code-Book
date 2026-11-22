"""Audit historical PLU references against the library without approving new identities."""
import argparse
import csv
import hashlib
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

import pymupdf

from library_assets import resolve_image
from used_image_library import OWNER_CANDIDATES, local_inventory, save_if_changed

ROOT = Path(__file__).resolve().parents[1]
REFERENCE = Path("data/review-references/2011-PLU-Listing1.pdf")
RANGE_DASHES = "-\u2010\u2011\u2012\u2013\u2014"
CODE_TOKEN = re.compile(rf"(?<!\d)\d{{4,5}}(?:\s*[{RANGE_DASHES}]\s*\d{{4,5}})?(?!\d)")


def normalize(text):
    return " ".join(text.split())


def reference_lines(page):
    lines = []
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            spans = line["spans"]
            lines.append(dict(
                text=normalize("".join(span["text"] for span in spans if not span["flags"] & 1)),
                raw="".join(span["text"] for span in spans),
                left=line["bbox"][0],
                superscripts=[span["text"] for span in spans if span["flags"] & 1],
            ))
    return lines


def extract_references(pages):
    entries, current, section = [], None, ""
    for page_number, lines in enumerate(pages, 1):
        for line in lines:
            text = line["text"]
            if not text:
                continue
            if text in {"\u2022", "\u25cb"}:
                if current is not None:
                    entries.append(current)
                current = dict(section=section, page_start=page_number, page_end=page_number,
                               lines=[], raw_lines=[], superscripts=[])
            elif line["left"] < 90:
                if current is not None:
                    entries.append(current)
                    current = None
                section = text
            elif current is not None:
                current["lines"].append(text)
                current["raw_lines"].append(line["raw"])
                current["superscripts"].extend(line["superscripts"])
                current["page_end"] = page_number
    if current is not None:
        entries.append(current)
    result = []
    for position, entry in enumerate(entries, 1):
        text = normalize(" ".join(entry["lines"]))
        groups = list(re.finditer(r"\(([^()]*)\)", text))
        tokens = [(match.group(), group.group(1)) for group in groups
                  for match in CODE_TOKEN.finditer(group.group(1))]
        if text.casefold().startswith("no listing"):
            tokens = [(match.group(), text) for match in CODE_TOKEN.finditer(text)]
        if not tokens:
            raise ValueError(f"reference entry has no parseable code on page {entry['page_start']}: {text}")
        kind = ("UNLISTED" if text.casefold().startswith("no listing") else
                "RETAILER_ASSIGNED" if "retailer assigned" in text.casefold() else
                "UNSPECIFIED_COMMODITY" if entry["section"] == "Other" else
                "NAMED_PRODUCT")
        for code, group in tokens:
            is_range = bool(re.search(f"[{RANGE_DASHES}]", code))
            if not is_range and len(code) != 4:
                raise ValueError(f"unexpected PLU length; review PDF extraction: {code}")
            result.append(dict(
                reference_position=position, section=entry["section"],
                page_start=entry["page_start"], page_end=entry["page_end"],
                plu=code if not is_range else "", code_range=code if is_range else "",
                reference_kind=kind, reference_text=text,
                source_text="\n".join(line.rstrip() for line in entry["raw_lines"]),
                source_code_group=group,
                superscript_footnotes=";".join(entry["superscripts"]),
            ))
    return result


def coverage_rows(references, options, source_sha256):
    by_plu = defaultdict(list)
    for row in options:
        by_plu[row["plu"]].append(row)
    records = []
    for reference in references:
        matches = by_plu.get(reference["plu"], []) if reference["reference_kind"] == "NAMED_PRODUCT" else []
        identities = {}
        for match in matches:
            path = match["path"]
            checksum = hashlib.sha256(path.read_bytes()).hexdigest()
            key = (match.get("source_id", match.get("candidate_id", "")),
                   match.get("product", match.get("name", "")), checksum, match["tier"])
            identities[key] = dict(
                source_id=key[0], product=key[1], image=str(path),
                sha256=checksum, approval_tier=match["tier"],
            )
        status = ("CODE_RANGE_OR_PLACEHOLDER_REVIEW" if reference["reference_kind"] != "NAMED_PRODUCT"
                  or reference["code_range"] else
                  "POTENTIAL_IMAGES_REQUIRE_IDENTITY_REVIEW" if identities else "MISSING_IMAGE")
        identity = hashlib.sha256(json.dumps(
            [source_sha256, reference["reference_position"], reference["plu"], reference["code_range"],
             reference["reference_text"]], ensure_ascii=False).encode()).hexdigest()
        records.append(dict(
            reference, reference_id=identity, reference_sha256=source_sha256,
            coverage_status=status, distinct_candidate_images=len({key[2] for key in identities}),
            candidate_identities=json.dumps(list(identities.values()), ensure_ascii=False, sort_keys=True),
            identity_review_status="OWNER_REVIEW_REQUIRED",
            image_approval_for_reference="NOT_GRANTED",
            picklist_action="NONE",
        ))
    return records


def build(root):
    root = Path(root)
    source = root / REFERENCE
    checksum = hashlib.sha256(source.read_bytes()).hexdigest()
    metadata = json.loads(source.with_suffix(".pdf.json").read_text())
    if metadata.get("sha256") != checksum or metadata.get("authorization") != "REFERENCE_ONLY":
        raise ValueError("historical PLU reference checksum or reference-only authorization mismatch")
    source_before = (root / "sources/plu/PickList.xml").read_bytes()
    with pymupdf.open(source) as pdf:
        lines = [reference_lines(page) for page in pdf]
        raw_pages = [page.get_text() for page in pdf]
    references = extract_references(lines)
    options = local_inventory(root)
    owner_index = root / "sources/image-library" / OWNER_CANDIDATES
    if owner_index.exists():
        for row in json.loads(owner_index.read_text()):
            options.append(dict(row, path=resolve_image(root / row["library_image"]),
                                tier="PREAPPROVED" if row.get("preapproval_status") ==
                                "APPROVED_FOR_EXACT_PLU_PRODUCT" else "REVIEWED"))
    rows = coverage_rows(references, options, checksum)
    for row in rows:
        identities = json.loads(row["candidate_identities"])
        for identity in identities:
            identity["image"] = str(Path(identity["image"]).relative_to(root))
        row["candidate_identities"] = json.dumps(identities, ensure_ascii=False, sort_keys=True)
    report = root / "outputs/reports"
    report.mkdir(parents=True, exist_ok=True)
    with (report / "FUTURE_PLU_IMAGE_COVERAGE.csv").open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=list(rows[0]), lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
    evidence = dict(reference=str(REFERENCE), sha256=checksum, pages=raw_pages)
    save_if_changed(report / "PLU_REFERENCE_SOURCE_TEXT.json",
                    (json.dumps(evidence, ensure_ascii=False, indent=2) + "\n").encode())
    summary = dict(
        reference=str(REFERENCE), reference_sha256=checksum, pages=len(lines),
        source_entries=len({row["reference_position"] for row in rows}),
        code_mentions=len(rows),
        distinct_named_plu_codes=len({row["plu"] for row in rows
                                     if row["reference_kind"] == "NAMED_PRODUCT" and row["plu"]}),
        code_ranges=sum(bool(row["code_range"]) for row in rows),
        statuses=dict(Counter(row["coverage_status"] for row in rows)),
        reference_kinds=dict(Counter(row["reference_kind"] for row in rows)),
        superscript_policy="Excluded using PDF span superscript flags; retained in source evidence.",
        range_policy="Ranges are preserved, never expanded into invented products.",
        approval_policy="Same PLU is a search lead only; each future identity and image requires owner review.",
        scope="Historical listing only, not an exhaustive current list of possible produce PLUs.",
        picklist_sha256=hashlib.sha256(source_before).hexdigest(),
        working_list_sha256=hashlib.sha256((root / "sources/plu/WorkingList.xml").read_bytes()).hexdigest(),
        image_library_indexes_sha256={
            path.name: hashlib.sha256(path.read_bytes()).hexdigest()
            for name in ("IMAGE_LIBRARY.csv", "WORKING_LIST_IMAGES.json", "LEGACY_BOOKLET_IMAGES.json",
                         OWNER_CANDIDATES, "IMAGE_ALIASES.json")
            if (path := root / "sources/image-library" / name).exists()
        },
        picklist_status="UNCHANGED", working_list_status="UNCHANGED",
        publication_authorization="NOT_GRANTED",
    )
    if (root / "sources/plu/PickList.xml").read_bytes() != source_before:
        raise ValueError("PickList changed during coverage generation; discard reports and retry")
    save_if_changed(report / "FUTURE_PLU_IMAGE_COVERAGE.json",
                    (json.dumps(summary, indent=2) + "\n").encode())
    return summary


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
