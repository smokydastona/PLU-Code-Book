import tempfile
import unittest
from pathlib import Path

from audit_till_binder import audit_images, audit_package, read_csv, write_blockers, write_csv


class TillBinderAuditChecks(unittest.TestCase):
    def test_blockers_stay_open_and_images_are_not_upscaled_or_hidden(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            fields = ["embedded", "embedded_sha256", "plu", "width", "height", "image_status",
                      "image_source", "organic_badge"]
            write_csv(output / "PHOTO_MAPPING.csv", [
                dict(embedded="True", embedded_sha256="a", plu="4011", width=192, height=192,
                     image_status="LEGACY_PREAPPROVED_IMAGE", image_source="Code Booklet", organic_badge="False"),
                dict(embedded="True", embedded_sha256="a", plu="4012", width=500, height=500,
                     image_status="LEGACY_PREAPPROVED_IMAGE", image_source="Code Booklet", organic_badge="False"),
                dict(embedded="False", embedded_sha256="", plu="4013", width=0, height=0,
                     image_status="MISSING", image_source="", organic_badge="False")], fields)
            write_csv(output / "MISSING_IMAGES.csv", [dict(plu="4013")], ["plu"])
            summary, small, shared = audit_images(output)
            self.assertEqual((summary["embedded"], summary["missing_or_uncertain"]), (2, 1))
            self.assertEqual(summary["images_shared_by_multiple_plus"], 1)
            self.assertEqual(len(small), 1)
            rows = write_blockers(dict(source_reconciliation=dict(
                blocker_counts={"PROHIBITED_IMAGE_REUSE": 3}, working_list_review_records=2,
                reconciliation_review_decisions=0, organic_placements_needing_review=1)), output)
            self.assertTrue(rows and all(row["status"] == "OPEN" for row in rows))
            self.assertEqual(len(read_csv(output / "PRODUCTION_BLOCKERS.csv")), 3)

    def test_package_audit_flags_removed_media(self):
        from openpyxl import Workbook
        from openpyxl.drawing.image import Image as SheetImage
        from PIL import Image
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            Image.new("RGB", (4, 4), "black").save(base / "barcode.png")
            original, built = Workbook(), Workbook()
            original.active.add_image(SheetImage(base / "barcode.png"), "A1")
            original.save(base / "old.xlsx")
            built.save(base / "new.xlsx")
            result = audit_package(base / "old.xlsx", base / "new.xlsx")
            self.assertEqual(result["original_media_parts"], 1)
            self.assertEqual(result["original_media_removed"], ["xl/media/image1.png"])
            result = audit_package(base / "old.xlsx", base / "old.xlsx")
            self.assertEqual(result["original_media_byte_identical"], 1)


if __name__ == "__main__":
    unittest.main()
