import csv
import hashlib
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pymupdf
from PIL import Image

from build_temporary_produce_review import PDF_NAME, SCREENING_NAME, build, candidates


class TemporaryReviewChecks(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.output = self.root / "outputs/review-booklets"
        self.output.mkdir(parents=True)
        (self.root / "data").mkdir()
        (self.root / "data/picklist_assignments.json").write_text('{"assignments":[]}')
        self.image = self.root / "data/candidate.jpg"
        Image.new("RGB", (500, 500), "green").save(self.image)
        self.item = dict(source_id="exact-id", source_ids=["original-id"], plu="00042",
                         name="Organic Cantaloupe – Small", included=True,
                         record_sha256="record-sha")
        self.row = dict(
            source_id="exact-id", source_identity_ids="original-id", plu="00042",
            product=self.item["name"], candidate_path="data/candidate.jpg",
            candidate_sha256=hashlib.sha256(self.image.read_bytes()).hexdigest(),
            width=500, height=500, visual_screening="VISUAL_SCREEN_PASS",
            visual_note="Product intact; owner review required.", owner_image_approval="NOT_APPROVED",
        )
        self.write_rows([self.row])
        mock = patch("build_temporary_produce_review.read_working_list",
                     return_value=([self.item], "source-sha"))
        mock.start()
        self.addCleanup(mock.stop)

    def write_rows(self, rows):
        with (self.output / SCREENING_NAME).open("w", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=list(rows[0]), lineterminator="\n")
            writer.writeheader()
            writer.writerows(rows)

    def test_refresh_replaces_duplicate_and_preserves_exact_text_and_unchecked_widget(self):
        old = self.output / "Produce_Image_Review_Temporary_old.pdf"
        old.write_bytes(b"old temporary copy")
        report = build(self.root)
        self.assertEqual(report["entries"], 1)
        self.assertFalse(old.exists())
        self.assertEqual(list(self.output.glob("Produce_Image_Review_Temporary*.pdf")),
                         [self.output / PDF_NAME])
        with pymupdf.open(self.output / PDF_NAME) as document:
            self.assertIn("00042", document[0].get_text())
            self.assertIn(self.item["name"], document[0].get_text())
            widgets = list(document[0].widgets())
            self.assertEqual(widgets[0].field_name, "approve_exact-id")
            self.assertEqual(widgets[0].field_value, "Off")

    def test_replacement_required_candidate_has_no_checkbox(self):
        self.row["visual_screening"] = "AUTO_FAIL_PRODUCT_MISMATCH"
        self.write_rows([self.row])
        report = build(self.root)
        self.assertEqual(report["replacement_required"], 1)
        with pymupdf.open(self.output / PDF_NAME) as document:
            self.assertEqual(list(document[0].widgets()), [])

    def test_changed_candidate_fails_without_replacing_current_pdf(self):
        build(self.root)
        original = (self.output / PDF_NAME).read_bytes()
        Image.new("RGB", (500, 500), "red").save(self.image)
        with self.assertRaisesRegex(ValueError, "exact screened candidate missing"):
            build(self.root)
        self.assertEqual((self.output / PDF_NAME).read_bytes(), original)

    def test_exact_checksum_relocation_is_reported_without_reusing_new_asset(self):
        library = self.root / "sources/image-library"
        library.mkdir(parents=True)
        original = self.image.read_bytes()
        retained = library / "retained.jpg"
        retained.write_bytes(original)
        Image.new("RGB", (500, 500), "red").save(self.image)
        with patch("build_temporary_produce_review.check_used_library"):
            pending, _, _ = candidates(self.root, self.output)
        self.assertEqual(pending[0]["resolved_candidate_path"], "sources/image-library/retained.jpg")
        self.assertEqual(pending[0]["candidate_path_resolution"], "EXACT_SHA256_RELOCATED")
        self.assertEqual(retained.read_bytes(), original)

    def test_changed_product_identity_cannot_reuse_screening(self):
        self.item["name"] = "Different Melon"
        with self.assertRaisesRegex(ValueError, "stale or ambiguous screening identity"):
            build(self.root)


if __name__ == "__main__":
    unittest.main()
