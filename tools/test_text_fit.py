import unittest
from pathlib import Path

from build_bakery_sheet import (PRODUCT_WIDTH as BAKERY_WIDTH, BODY_POINTS, ROW_HEIGHT,
                                read_source, workbook_bytes)
from build_bulk_sheet import PAGE_CATEGORIES, bulk_sheet
from text_fit import line_count, text_points
from openpyxl import load_workbook
import io

ROOT = Path(__file__).resolve().parents[1]


class TextFitChecks(unittest.TestCase):
    def test_wrapping_is_measured_not_estimated_from_character_count(self):
        self.assertEqual(line_count("Coconut Medium Unsweetened", 11, 26), 1)
        self.assertEqual(line_count("Coconut Shredded Unsweetened", 11, 26), 2)
        self.assertEqual(line_count("Unbroken", 11, 26), 1)
        self.assertGreater(text_points("WWWW", 11), text_points("iiii", 11))

    def test_bakery_rows_are_uniform_and_names_never_wrap(self):
        rows = read_source(ROOT)[2]
        for row in rows:
            self.assertEqual(line_count(row["product"], BODY_POINTS, BAKERY_WIDTH), 1, row["product"])
        content, coordinates = workbook_bytes(rows)
        sheet = load_workbook(io.BytesIO(content))["Bakery"]
        heights = {sheet.row_dimensions[int(entry["product_cell"][1:])].height for entry in coordinates}
        self.assertEqual(heights, {ROW_HEIGHT})

    def test_bulk_rows_are_tall_enough_for_every_wrapped_name(self):
        names = ("Coconut Shredded Unsweetened", "Trail Mix Mountain Man Unsalted", "Plain", "Walnuts")
        entries = [dict(category=category, product=names[index % 4], plu=f"{index:05d}")
                   for page in PAGE_CATEGORIES for category in page for index in range(4)]
        content, coordinates = bulk_sheet(entries)
        sheet = load_workbook(io.BytesIO(content))["Bulk"]
        self.assertEqual(len(coordinates), len(entries))
        for entry in coordinates:
            top = sheet[entry["product_cell"]]
            merged = next((r for r in sheet.merged_cells.ranges if top.coordinate in r), None)
            rows = range(top.row, (merged.max_row if merged else top.row) + 1)
            height = sum(sheet.row_dimensions[row].height for row in rows)
            self.assertGreaterEqual(height, 16 * line_count(entry["product"], 11, 26), entry["product"])


if __name__ == "__main__":
    unittest.main()
