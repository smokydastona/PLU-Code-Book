"""Impose a complete binder PDF as a half-fold, saddle-stitched booklet."""
import csv
import math
from collections import Counter
from pathlib import Path

import pymupdf

SHEET_WIDTH = 792
SHEET_HEIGHT = 612
PANEL_WIDTH = SHEET_WIDTH / 2
OUTPUT_NAME = "Pocket Guide.pdf"
MAP_NAME = "BINDER_BOOKLET_IMPOSITION.csv"
ROTATED_PAGES = ("Produce", "Bulk", "Store Directory")


def _slot_text(page, slot):
    rect = pymupdf.Rect(slot * PANEL_WIDTH, 0, (slot + 1) * PANEL_WIDTH, SHEET_HEIGHT)
    return " ".join(page.get_text(clip=rect).split())


def _same_text(first, second):
    return Counter(first.split()) == Counter(second.split())


def _sheet_sides(page_count):
    padded_count = math.ceil(page_count / 4) * 4
    for index in range(padded_count // 4):
        yield (
            (index + 1, "front", padded_count - 2 * index, 1 + 2 * index),
            (index + 1, "back", 2 + 2 * index, padded_count - 1 - 2 * index),
        )


def impose_binder_pdf(source_path, output_path, map_path, rotate_pages=()):
    source = pymupdf.open(source_path)
    if source.is_closed or len(source) == 0:
        raise ValueError("the complete binder PDF has no pages")
    page_count = len(source)
    padded_count = math.ceil(page_count / 4) * 4
    sheet_count = padded_count // 4
    imposed = pymupdf.open()
    mapping = []
    output_page_number = 0

    def add_side(sheet_number, side, left, right):
        nonlocal output_page_number
        page = imposed.new_page(width=SHEET_WIDTH, height=SHEET_HEIGHT)
        output_page_number += 1
        scales = []
        rotations = []
        for slot, source_number in enumerate((left, right)):
            scale = ""
            if source_number <= page_count:
                original = source[source_number - 1]
                rect = pymupdf.Rect(slot * PANEL_WIDTH, 0, (slot + 1) * PANEL_WIDTH, SHEET_HEIGHT)
                rotation = 90 if source_number in rotate_pages else 0
                width, height = ((original.rect.height, original.rect.width) if rotation else
                                 (original.rect.width, original.rect.height))
                scale = min(rect.width / width, rect.height / height)
                page.show_pdf_page(rect, source, source_number - 1, keep_proportion=True,
                                   rotate=rotation)
                if not _same_text(_slot_text(page, slot), original.get_text()):
                    raise ValueError(
                        f"imposed sheet {sheet_number} {side} clips text from binder page {source_number}")
            else:
                rotation = ""
            scales.append(f"{scale:.6f}" if scale else "")
            rotations.append(rotation)
        mapping.append(dict(
            sheet=sheet_number, side=side, imposed_page=output_page_number,
            left_source_page=left if left <= page_count else "",
            right_source_page=right if right <= page_count else "",
            left_scale=scales[0], right_scale=scales[1],
            left_rotation="CCW_90" if rotations[0] == 90 else
            ("0" if rotations[0] == 0 else ""),
            right_rotation="CCW_90" if rotations[1] == 90 else
            ("0" if rotations[1] == 0 else ""),
        ))

    try:
        for front, back in _sheet_sides(page_count):
            add_side(*front)
            add_side(*back)
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        imposed.set_metadata({
            "title": "Pocket Guide",
            "subject": "2-up duplex imposition for half-fold saddle-stitch printing",
        })
        imposed.save(output, garbage=4, deflate=True)
    finally:
        imposed.close()
        source.close()
    with Path(map_path).open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=[
            "sheet", "side", "imposed_page", "left_source_page", "right_source_page",
            "left_scale", "right_scale", "left_rotation", "right_rotation",
        ], lineterminator="\n")
        writer.writeheader()
        writer.writerows(mapping)
    return dict(
        source_pages=page_count, padded_pages=padded_count,
        blank_padding_pages=padded_count - page_count, physical_sheets=sheet_count,
        imposed_pdf_pages=output_page_number, paper="US Letter landscape",
        print_mode="2-up duplex, flip on short edge, 100% scale",
        fold="half-fold, saddle-stitch",
    )


def audit_imposed_booklet(source_path, booklet_path, map_path, rotate_pages=()):
    source = pymupdf.open(source_path)
    booklet = pymupdf.open(booklet_path)
    try:
        if not source or not booklet:
            raise ValueError("binder or imposed booklet PDF has no pages")
        source_count = len(source)
        padded_count = math.ceil(source_count / 4) * 4
        sheets = padded_count // 4
        if len(booklet) != sheets * 2:
            raise ValueError("imposed booklet has an incorrect physical-side count")
        with Path(map_path).open(encoding="utf-8-sig", newline="") as file:
            rows = list(csv.DictReader(file))
        if len(rows) != len(booklet):
            raise ValueError("imposition map does not cover every output page")
        expected_sides = [side for pair in _sheet_sides(source_count) for side in pair]
        for row, (sheet, side, left, right) in zip(rows, expected_sides):
            expected = (str(sheet), side,
                        str(left) if left <= source_count else "",
                        str(right) if right <= source_count else "")
            observed = (row["sheet"], row["side"], row["left_source_page"], row["right_source_page"])
            if (observed != expected or
                    any(rotation not in {"", "0", "CCW_90"}
                        for rotation in (row["left_rotation"], row["right_rotation"]))):
                raise ValueError("imposition map does not follow saddle-stitch page order")
        observed_source_pages = []
        for index, row in enumerate(rows):
            page = booklet[index]
            if (abs(page.rect.width - SHEET_WIDTH) > .1 or
                    abs(page.rect.height - SHEET_HEIGHT) > .1):
                raise ValueError(f"imposed page {index + 1} is not landscape US Letter")
            for slot, column in enumerate(("left_source_page", "right_source_page")):
                value = row[column]
                rotation_column = f"{'left' if slot == 0 else 'right'}_rotation"
                if not value:
                    if _slot_text(page, slot) or row[rotation_column]:
                        raise ValueError(f"imposed page {index + 1} has content in its padding panel")
                    continue
                source_number = int(value)
                if not 1 <= source_number <= source_count:
                    raise ValueError(f"imposition map references invalid source page {source_number}")
                expected_rotation = "CCW_90" if source_number in rotate_pages else "0"
                if row[rotation_column] != expected_rotation:
                    raise ValueError(f"source page {source_number} has an incorrect booklet rotation")
                source_images = source[source_number - 1].get_image_info()
                panel_rect = pymupdf.Rect(slot * PANEL_WIDTH, 0,
                                          (slot + 1) * PANEL_WIDTH, SHEET_HEIGHT)
                panel_images = [
                    image for image in page.get_image_info()
                    if pymupdf.Rect(image["bbox"]).intersects(panel_rect)
                ]
                if (len(panel_images) != len(source_images) or
                        any(not panel_rect.contains(pymupdf.Rect(image["bbox"]))
                            for image in panel_images)):
                    raise ValueError(f"imposed page {index + 1} does not preserve source page "
                                     f"{source_number} images within its panel")
                observed_source_pages.append(source_number)
                if not _same_text(_slot_text(page, slot),
                                  source[source_number - 1].get_text()):
                    raise ValueError(f"imposed page {index + 1} does not preserve source page {source_number}")
        if sorted(observed_source_pages) != list(range(1, source_count + 1)):
            raise ValueError("imposed booklet does not contain every source page exactly once")
        return dict(
            technical_checks="PASS", source_pages=source_count, padded_pages=padded_count,
            blank_padding_pages=padded_count - source_count, physical_sheets=sheets,
            imposed_pdf_pages=len(booklet), all_source_pages_preserved_once=True,
            all_panels_fit_without_text_clipping=True,
            rotated_source_pages=sorted(rotate_pages),
        )
    finally:
        source.close()
        booklet.close()
