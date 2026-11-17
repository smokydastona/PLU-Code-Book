import csv
import tempfile
import unittest
from pathlib import Path

import pymupdf

from impose_binder_booklet import audit_imposed_booklet, impose_binder_pdf


class BinderBookletImpositionChecks(unittest.TestCase):
    def test_all_pages_are_imposed_in_saddle_stitch_order_with_padding(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source_path = root / "binder.pdf"
            booklet_path = root / "Pocket Guide.pdf"
            map_path = root / "imposition.csv"
            source = pymupdf.open()
            sizes = [(612, 792), (792, 612), (612, 792), (792, 612), (612, 792)]
            for number, (width, height) in enumerate(sizes, 1):
                page = source.new_page(width=width, height=height)
                page.insert_text((40, 50), f"Binder page {number}")
            source.save(source_path)
            source.close()

            summary = impose_binder_pdf(source_path, booklet_path, map_path)
            self.assertEqual(summary, dict(
                source_pages=5, padded_pages=8, blank_padding_pages=3, physical_sheets=2,
                imposed_pdf_pages=4, paper="US Letter landscape",
                print_mode="2-up duplex, flip on short edge, 100% scale",
                fold="half-fold, saddle-stitch",
            ))
            with map_path.open(encoding="utf-8-sig", newline="") as file:
                rows = list(csv.DictReader(file))
            self.assertEqual(
                [(row["left_source_page"], row["right_source_page"]) for row in rows],
                [("", "1"), ("2", ""), ("", "3"), ("4", "5")],
            )
            audit = audit_imposed_booklet(source_path, booklet_path, map_path)
            self.assertTrue(audit["all_source_pages_preserved_once"])
            self.assertTrue(audit["all_panels_fit_without_text_clipping"])
            rows[0]["right_source_page"], rows[-1]["right_source_page"] = (
                rows[-1]["right_source_page"], rows[0]["right_source_page"])
            with map_path.open("w", newline="", encoding="utf-8-sig") as file:
                writer = csv.DictWriter(file, fieldnames=rows[0].keys(), lineterminator="\n")
                writer.writeheader()
                writer.writerows(rows)
            with self.assertRaisesRegex(ValueError, "saddle-stitch page order"):
                audit_imposed_booklet(source_path, booklet_path, map_path)

    def test_four_page_signature_places_first_and_last_on_outer_sheet(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source_path, booklet_path, map_path = (
                root / "source.pdf", root / "Pocket Guide.pdf", root / "map.csv")
            source = pymupdf.open()
            for number in range(1, 5):
                page = source.new_page(width=612, height=792)
                page.insert_text((40, 50), f"Page {number}")
            source.save(source_path)
            source.close()
            impose_binder_pdf(source_path, booklet_path, map_path)
            with map_path.open(encoding="utf-8-sig", newline="") as file:
                rows = list(csv.DictReader(file))
            self.assertEqual(
                [(row["left_source_page"], row["right_source_page"]) for row in rows],
                [("4", "1"), ("2", "3")],
            )


if __name__ == "__main__":
    unittest.main()
