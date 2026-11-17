import csv
import hashlib
import json
import tempfile
import unittest
import zipfile
from pathlib import Path

from package_publication import PackageError, REQUIRED_FILES, create_package


class PackagePublicationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.build = self.root / "build"
        self.build.mkdir()
        for name in REQUIRED_FILES:
            path = self.build / name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text("fixture\n", encoding="utf-8")
        (self.build / "reconciliation/PICKLIST_SYNC.csv").write_text(
            "source_id,image_approval_basis\nabc,EXACT_OWNER_APPROVAL\n", encoding="utf-8")
        (self.build / "VALIDATION.json").write_text(json.dumps({
            "status": "REVIEW_ONLY_NOT_RELEASED",
            "production_publication": "NOT_AUTHORIZED_REVIEW_BUILD",
            "validations": {"photos": {"status": "PASS"}},
            "print_validation": {"photos": {"status": "PASS"}},
        }), encoding="utf-8")
        (self.build / "AUDIT_REPORT.json").write_text(json.dumps({
            "technical_checks": "PASS", "failures": [], "open_blockers": {},
        }), encoding="utf-8")
        self.readiness_path = self.build / "reconciliation/PRODUCTION_READINESS.json"
        self.set_readiness()

    def tearDown(self):
        self.temp.cleanup()

    def set_readiness(self, blockers=None, status="AUTOMATED_CHECKS_PASSED_REQUIRES_RELEASE_REVIEW"):
        self.readiness_path.write_text(json.dumps({
            "status": status,
            "release_blockers": blockers or {},
            "owner_review_findings": {"DUPLICATE_IMAGE_REUSE_REVIEW": 3},
            "counters": {
                "unresolved_source_record_reconciliation": 0,
                "unresolved_product_identity": 0,
                "unresolved_image_selection": sum((blockers or {}).values()),
                "classification_section_reviews": 0,
                "image_quality_warnings": 1,
            },
            "gate_statuses": {
                "data_reconciliation": "PASS",
                "image_selection": "PASS",
                "classification_placement": "PASS",
                "production_release": "OWNER_AUTHORIZATION_REQUIRED",
            },
        }), encoding="utf-8")

    def test_passing_fixture_packages_all_deliverables_and_consistent_hashes(self):
        archive_path = self.root / "publication.zip"
        manifest = create_package(self.build, archive_path, self.root, mode="production",
                                  commit="a" * 40)
        with zipfile.ZipFile(archive_path) as archive:
            self.assertIsNone(archive.testzip())
            names = set(archive.namelist())
            self.assertTrue(set(REQUIRED_FILES).issubset(names))
            self.assertIn("RELEASE_MANIFEST.json", names)
            packaged_manifest = json.loads(archive.read("RELEASE_MANIFEST.json"))
            for name, digest in packaged_manifest["output_sha256"].items():
                self.assertEqual(hashlib.sha256(archive.read(name)).hexdigest(), digest)
        self.assertTrue(manifest["production_gate_pass"])
        self.assertEqual(manifest["generator"]["repository_commit"], "a" * 40)
        self.assertEqual(
            manifest["validation"]["review_build_publication_status"],
            "NOT_AUTHORIZED_REVIEW_BUILD")
        self.assertEqual(
            manifest["approval_metadata_available"]["image_approval_basis_counts"],
            {"EXACT_OWNER_APPROVAL": 1})

    def test_mandatory_blocker_prevents_production_and_preserves_existing_package(self):
        self.set_readiness({"EXPLICIT_IMAGE_REJECTION": 1}, status="BLOCKED")
        archive_path = self.root / "approved-existing.zip"
        archive_path.write_bytes(b"previous-approved-publication")
        with self.assertRaisesRegex(PackageError, "EXPLICIT_IMAGE_REJECTION"):
            create_package(self.build, archive_path, self.root, mode="production")
        self.assertEqual(archive_path.read_bytes(), b"previous-approved-publication")

    def test_blocked_review_build_is_packaged_without_becoming_production(self):
        self.set_readiness({"EXPLICIT_IMAGE_REJECTION": 1}, status="BLOCKED")
        archive_path = self.root / "review.zip"
        manifest = create_package(self.build, archive_path, self.root, mode="review")
        self.assertFalse(manifest["production_gate_pass"])
        self.assertEqual(manifest["publication_status"], "BLOCKED")
        with zipfile.ZipFile(archive_path) as archive:
            self.assertIn("RELEASE_MANIFEST.json", archive.namelist())

    def test_binder_audit_blocker_prevents_production_package(self):
        (self.build / "AUDIT_REPORT.json").write_text(json.dumps({
            "technical_checks": "PASS", "failures": [],
            "open_blockers": {"EXPLICIT_IMAGE_REJECTION": 1},
        }), encoding="utf-8")
        archive_path = self.root / "blocked.zip"
        with self.assertRaisesRegex(PackageError, "binder audit EXPLICIT_IMAGE_REJECTION"):
            create_package(self.build, archive_path, self.root, mode="production")
        self.assertFalse(archive_path.exists())

    def test_missing_deliverable_fails_before_creating_package(self):
        (self.build / REQUIRED_FILES[0]).unlink()
        archive_path = self.root / "incomplete.zip"
        with self.assertRaisesRegex(PackageError, "missing required deliverables"):
            create_package(self.build, archive_path, self.root)
        self.assertFalse(archive_path.exists())

    def test_existing_package_is_never_overwritten(self):
        archive_path = self.root / "existing.zip"
        archive_path.write_bytes(b"keep")
        with self.assertRaisesRegex(PackageError, "refusing to overwrite"):
            create_package(self.build, archive_path, self.root, mode="review")
        self.assertEqual(archive_path.read_bytes(), b"keep")


if __name__ == "__main__":
    unittest.main()
