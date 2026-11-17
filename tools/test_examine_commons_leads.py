import csv
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from examine_commons_leads import examine


class ExamineCommonsLeadsTests(unittest.TestCase):
    def test_exact_file_and_search_leads_are_not_approved(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            leads = root / "leads.csv"
            with leads.open("w", newline="", encoding="utf-8") as file:
                writer = csv.DictWriter(file, fieldnames=["plu", "product", "source_url", "source_type"])
                writer.writeheader()
                writer.writerows([
                    dict(plu="0001", product="Melon", source_type="FILE_PAGE",
                         source_url="https://commons.wikimedia.org/wiki/File:Melon.jpg"),
                    dict(plu="0002", product="Melon", source_type="SEARCH_QUERY",
                         source_url="https://commons.wikimedia.org/w/index.php?search=melon"),
                ])
            page = dict(title="File:Melon.jpg", imageinfo=[dict(
                width=1000, height=800, extmetadata={"Artist": {"value": "Creator"}})])
            with patch("examine_commons_leads.category_pages", return_value=([page], None)), \
                    patch("examine_commons_leads.search", return_value=[page]) as find, \
                    patch("examine_commons_leads.download", return_value=False):
                self.assertFalse(examine(leads, root / "out", "test"))
            self.assertEqual(find.call_args_list[0].kwargs["title"], "File:Melon.jpg")
            self.assertEqual(find.call_args_list[1].args[0], "melon")
            results = json.loads((root / "out/lead-results.json").read_text())
            self.assertEqual([r["status"] for r in results], ["UNRESOLVED", "UNRESOLVED"])
            self.assertEqual(results[0]["metadata"][0]["creator"], "Creator")
            self.assertEqual(len(results[0]["category_matches"]), 2)

    def test_category_failure_is_visible_and_fails_run(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            leads = root / "leads.csv"
            leads.write_text("plu,product,source_url,source_type\n"
                             "1,Melon,https://commons.wikimedia.org/wiki/File:Melon.jpg,FILE_PAGE\n")
            with patch("examine_commons_leads.category_pages", side_effect=OSError("offline")), \
                    patch("examine_commons_leads.search", side_effect=OSError("offline")):
                self.assertTrue(examine(leads, root / "out", "test"))
            categories = json.loads((root / "out/categories.json").read_text())
            self.assertEqual(len(categories["errors"]), 2)
            result = json.loads((root / "out/lead-results.json").read_text())[0]
            self.assertEqual(result["status"], "ERROR")
            self.assertIn("offline", result["details"])


if __name__ == "__main__":
    unittest.main()
