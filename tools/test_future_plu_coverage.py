import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock

from PIL import Image

from build_future_plu_coverage import CODE_TOKEN, coverage_rows, extract_references, reference_lines


def line(text, left=108):
    return dict(text=text, raw=text, left=left, superscripts=[])


class FutureCoverageChecks(unittest.TestCase):
    def test_superscripts_do_not_become_five_digit_codes(self):
        page = Mock()
        page.get_text.return_value = {"blocks": [{"lines": [{
            "bbox": [108, 0, 400, 12],
            "spans": [{"text": "Apple (4108", "flags": 0},
                      {"text": "1,2", "flags": 1},
                      {"text": ", 4109)", "flags": 0}],
        }]}]}
        lines = reference_lines(page)
        self.assertEqual(lines[0]["superscripts"], ["1,2"])
        rows = extract_references([[line("Apple", 72), line("\u2022"), *lines]])
        self.assertEqual([row["plu"] for row in rows], ["4108", "4109"])
        self.assertEqual(rows[0]["source_text"], "Apple (41081,2, 4109)")
        self.assertEqual(rows[0]["superscript_footnotes"], "1,2")

    def test_leading_zeros_ranges_and_cross_page_continuations_are_preserved(self):
        rows = extract_references([
            [line("Apple", 72), line("\u2022"), line("Small (0042),")],
            [line("large (0043)"), line("\u2022"), line("retailer assigned (4500\u20104513)")],
        ])
        self.assertEqual([row["plu"] for row in rows], ["0042", "0043", ""])
        self.assertEqual(rows[0]["page_end"], 2)
        self.assertEqual(rows[-1]["code_range"], "4500\u20104513")
        self.assertEqual(rows[-1]["reference_kind"], "RETAILER_ASSIGNED")

    def test_unparseable_entry_is_reported_not_silently_dropped(self):
        with self.assertRaisesRegex(ValueError, "no parseable code"):
            extract_references([[line("Apple", 72), line("\u2022"), line("New variety")]])

    def test_same_code_candidates_do_not_approve_future_identity(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "apple.jpg"
            buffer = io.BytesIO()
            Image.new("RGB", (1, 1), "red").save(buffer, format="JPEG")
            path.write_bytes(buffer.getvalue())
            references = extract_references([[line("Apple", 72), line("\u2022"),
                                               line("Different variety (0042)"),
                                               line("\u2022"), line("Missing (0043)")]])
            rows = coverage_rows(references, [
                dict(plu="0042", product="Approved variety", source_id="approved-id",
                     path=path, tier="APPROVED"),
            ], "source-sha")
            self.assertEqual(rows[0]["coverage_status"], "POTENTIAL_IMAGES_REQUIRE_IDENTITY_REVIEW")
            self.assertEqual(rows[0]["image_approval_for_reference"], "NOT_GRANTED")
            self.assertEqual(rows[1]["coverage_status"], "MISSING_IMAGE")
            self.assertEqual(rows[0]["picklist_action"], "NONE")

    def test_uploaded_reference_accounts_for_every_code_token(self):
        import pymupdf
        from collections import Counter

        source = Path(__file__).resolve().parents[1] / "data/review-references/2011-PLU-Listing1.pdf"
        with pymupdf.open(source) as pdf:
            pages = [reference_lines(page) for page in pdf]
        tokens = Counter(match.group() for page in pages for item in page
                         for match in CODE_TOKEN.finditer(item["text"]))
        rows = extract_references(pages)
        self.assertEqual(Counter(row["plu"] or row["code_range"] for row in rows), tokens)
        self.assertFalse(any(row["plu"] == "41081" for row in rows))


if __name__ == "__main__":
    unittest.main()
