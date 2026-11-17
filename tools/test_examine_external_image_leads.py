import json
import socket
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from urllib.error import URLError

from examine_external_image_leads import PageMetadata, examine, read_leads, validate_url


class ExamineExternalImageLeadsTests(unittest.TestCase):
    def test_html_metadata_parser_keeps_license_claim_unverified(self):
        parser = PageMetadata()
        parser.feed(
            '<title>Romaine Lettuce</title>'
            '<meta property="og:image" content="https://example.test/preview.jpg">'
            '<meta property="og:image:width" content="1200">'
            '<meta name="license" content="Commercial license">'
            '<link rel="canonical" href="https://example.test/lettuce">'
        )
        self.assertEqual(parser.title, "Romaine Lettuce")
        self.assertEqual(parser.metadata["license"], "Commercial license")
        self.assertEqual(parser.metadata["og:image:width"], "1200")
        self.assertEqual(parser.metadata["canonical"], "https://example.test/lettuce")

    def test_manifest_rejects_non_https_and_unapproved_hosts(self):
        for url in ("http://pixabay.com/page", "https://example.org/page",
                    "https://user@pixabay.com/page"):
            with self.subTest(url=url), self.assertRaises(ValueError):
                validate_url(url)

    def test_examine_records_identity_and_never_approves_source(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest = root / "leads.json"
            manifest.write_text(json.dumps([{
                "plu": "4640", "product": "Romaine Lettuce",
                "source_url": "https://pixabay.com/photos/lettuce-romaine-greens-vegetable-2468495/"
            }]))
            output = root / "out"
            item = dict(
                plu="4640", name="Romaine Lettuce", source_id="working-id",
                source_ids=["picklist-id"], record_sha256="record-hash",
                included=True,
            )
            response = dict(
                status="PAGE_METADATA_FETCHED", http_status=200,
                final_url="https://pixabay.com/page", content_type="text/html",
                content_length="100", title="Romaine Lettuce", description="Page description",
                canonical_url="", preview_url="https://pixabay.com/preview.jpg",
                preview_width="1200", preview_height="800",
                license_claim="Pixabay License", license_url="", details="HTML metadata only",
            )
            with patch("examine_external_image_leads.read_working_list",
                       return_value=([item], "picklist-hash")), \
                    patch("examine_external_image_leads.fetch_page", return_value=response):
                rows = examine(manifest, output, root)
            self.assertEqual(rows[0]["working_list_id"], "working-id")
            self.assertEqual(rows[0]["working_source_ids"], "picklist-id")
            self.assertEqual(rows[0]["license_verification_status"], "NOT_VERIFIED")
            self.assertEqual(rows[0]["rights_verification_status"], "NOT_VERIFIED")
            self.assertEqual(rows[0]["photographic_origin_verification_status"], "NOT_VERIFIED")
            self.assertEqual(rows[0]["visual_approval_status"], "NOT_APPROVED")
            self.assertEqual(rows[0]["resolution_verification_status"], "NOT_VERIFIED")
            self.assertEqual(rows[0]["production_status"], "NOT_APPROVED")
            saved = json.loads((output / "source-findings.json").read_text())
            self.assertEqual(saved[0]["license_claim"], "Pixabay License")

    def test_search_only_url_is_not_fetched(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest = root / "leads.json"
            manifest.write_text(json.dumps([{
                "plu": "4425", "product": "Shinko (Shingo) Asian Pear",
                "source_url": "https://www.google.com/search?tbm=isch&q=Shinko+Asian+pear",
                "search_only": True,
            }]))
            item = dict(
                plu="4425", name="Shinko (Shingo) Asian Pear", source_id="working-id",
                source_ids=["picklist-id"], record_sha256="record-hash", included=True,
            )
            with patch("examine_external_image_leads.read_working_list",
                       return_value=([item], "picklist-hash")), \
                    patch("examine_external_image_leads.fetch_page") as fetch:
                rows = examine(manifest, root / "out", root)
            fetch.assert_not_called()
            self.assertEqual(rows[0]["status"], "SEARCH_URL_ONLY")

    def test_dns_failure_is_reported_separately_from_http_errors(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest = root / "leads.json"
            manifest.write_text(json.dumps([{
                "plu": "4640", "product": "Romaine Lettuce",
                "source_url": "https://pixabay.com/photos/lettuce-romaine-greens-vegetable-2468495/"
            }]))
            item = dict(
                plu="4640", name="Romaine Lettuce", source_id="working-id",
                source_ids=["picklist-id"], record_sha256="record-hash",
                included=True,
            )
            with patch("examine_external_image_leads.read_working_list",
                       return_value=([item], "picklist-hash")), \
                    patch("examine_external_image_leads.fetch_page",
                          side_effect=URLError(socket.gaierror(socket.EAI_NONAME, "no address"))):
                rows = examine(manifest, root / "out", root)
            self.assertEqual(rows[0]["status"], "DNS_ERROR")
            self.assertEqual(rows[0]["production_status"], "NOT_APPROVED")


if __name__ == "__main__":
    unittest.main()
