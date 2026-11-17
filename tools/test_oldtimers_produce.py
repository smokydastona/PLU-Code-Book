import tempfile
import unittest
from pathlib import Path
import hashlib

import pymupdf

from build_oldtimers_produce import generate
from audit_till_binder import audit_oldtimers, write_csv


def row(source_id, plu, name, product, book, variety=""):
    return dict(source_id=source_id, plu=plu, display_name=name, sort_product=product,
                sort_variety=variety, organic=False, book=book)


class OldTimersProduceTests(unittest.TestCase):
    def test_exact_entries_legacy_split_and_continuation(self):
        rows = [
            row("1", "00123", "MELON – SMALL", "MELON", "Vegetable"),
            row("2", "90001", "Org. TOMATO ON THE VINE", "TOMATO", "Vegetable"),
        ]
        rows.extend(row(f"apple-{i}", str(4000 + i), f"APPLES VARIETY {i}", "APPLE", "Fruit", str(i))
                    for i in range(160))
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "oldtimers.pdf"
            summary = generate(path, rows)
            self.assertEqual((summary["records"], summary["fruit_records"],
                              summary["vegetable_records"]), (162, 161, 1))
            with pymupdf.open(path) as document:
                texts = [page.get_text() for page in document]
                self.assertIn("00123", " ".join(texts))
                self.assertIn("MELON – SMALL", " ".join(texts))
                self.assertIn("TOMATO", texts[-1])
                self.assertIn("VEGETABLES", texts[-1])
                self.assertTrue(any("APPLE (CONT.)" in text for text in texts))
                self.assertTrue(all(not page.get_images() for page in document))

    def test_rejects_non_text_plu_and_invalid_classification(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "oldtimers.pdf"
            with self.assertRaisesRegex(ValueError, "PLU"):
                generate(path, [row("1", 123, "APPLE", "APPLE", "Fruit")])
            with self.assertRaisesRegex(ValueError, "classification"):
                generate(path, [row("1", "123", "APPLE", "APPLE", "Other")])

    def test_audit_rejects_modified_pdf_or_mismatched_binder(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            path = output / "OldTimers_Produce.pdf"
            summary = generate(path, [row("1", "00123", "APPLE", "APPLE", "Fruit")])
            summary["sha256"] = hashlib.sha256(path.read_bytes()).hexdigest()
            write_csv(output / "PRODUCE_LAYOUT.csv", [dict(plu="00123")], ["plu"])
            audit_oldtimers(output, {"oldtimers_produce": summary})
            write_csv(output / "PRODUCE_LAYOUT.csv", [], ["plu"])
            with self.assertRaisesRegex(ValueError, "record count"):
                audit_oldtimers(output, {"oldtimers_produce": summary})
            path.write_bytes(b"modified")
            with self.assertRaisesRegex(ValueError, "differs"):
                audit_oldtimers(output, {"oldtimers_produce": summary})


if __name__ == "__main__":
    unittest.main()
