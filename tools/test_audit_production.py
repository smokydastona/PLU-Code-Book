import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image

from audit_production import combined_summary, image_findings, inventory, main


class ProductionAuditChecks(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        (self.root / "data").mkdir()
        self.image = self.root / "data/apple.png"
        Image.new("RGB", (500, 500), "red").save(self.image)
        self.row = dict(idx="1", book="Fruit", plu="03283", name="Apple",
                        image="apple.png", width="500", height="500", source="Photo")
        self.review = dict(self.row, verdict="Approved")

    def test_missing_optional_metadata_is_reported_without_revoking_approval(self):
        _, _, errors = image_findings(self.row, self.review, self.root)
        self.assertIn("PROVENANCE:MISSING_IMAGE_METADATA", errors)
        self.assertFalse(any(error.startswith("EXPLICIT_IMAGE_REJECTION") for error in errors))

    def test_reports_resolution_ai_and_visual_problems_separately(self):
        Image.new("RGB", (1000, 499)).save(self.image)
        self.row["source"] = "AI Generated"
        self.review["verdict"] = "Rejected"
        width, height, errors = image_findings(self.row, self.review, self.root)
        self.assertEqual((width, height), (1000, 499))
        for error in ("IMAGE_QUALITY:BELOW_500_PX", "PROVENANCE:RECORDED_DIMENSIONS_MISMATCH",
                      "EXPLICIT_IMAGE_REJECTION", "PROVENANCE:AI_REFERENCE_NOT_PHOTOGRAPH"):
            self.assertIn(error, errors)

    def test_candidate_metadata_is_not_rights_or_visual_approval(self):
        metadata = dict(
            sha256="stale", plu="3283", product="Apple", approval_status="CANDIDATE",
            creator="Photographer", license="CC0", attribution="Photographer",
            source_page_url="https://example.org/apple", license_url="https://example.org/cc0",
        )
        self.image.with_suffix(".png.json").write_text(json.dumps(metadata))
        _, _, errors = image_findings(self.row, self.review, self.root)
        self.assertIn("RIGHTS:VERIFICATION_NOT_RECORDED", errors)
        self.assertIn("PROVENANCE:VISUAL_APPROVAL_NOT_RECORDED", errors)
        self.assertIn("PROVENANCE:CHECKSUM_MISMATCH", errors)
        self.assertIn("PROVENANCE:ASSIGNMENT_MISMATCH", errors)

    def test_all_current_conditions_required_for_automated_checks(self):
        metadata = dict(
            sha256=hashlib.sha256(self.image.read_bytes()).hexdigest(),
            plu="03283", product="Apple", creator="Photographer", license="CC0",
            attribution="Photographer", source_page_url="https://example.org/apple",
            license_url="https://example.org/cc0", rights_verification_status="VERIFIED",
            visual_approval_status="APPROVED",
        )
        self.image.with_suffix(".png.json").write_text(json.dumps(metadata))
        self.assertEqual(image_findings(self.row, self.review, self.root)[2], [])

    def test_owner_confirmed_permission_is_not_mislabeled_as_a_public_licence(self):
        metadata = dict(
            sha256=hashlib.sha256(self.image.read_bytes()).hexdigest(),
            plu="03283", product="Apple", creator="Not separately identified by the owner",
            license="OWNER_CONFIRMED_PERMISSION", rights_status="OWNER_CONFIRMED_PERMISSION",
            attribution="Owner-provided image", source_reference="Owner-uploaded image attachment",
            rights_evidence="Owner confirmed permission for this repository",
            rights_verification_status="VERIFIED", visual_approval_status="APPROVED",
        )
        self.image.with_suffix(".png.json").write_text(json.dumps(metadata))
        self.assertEqual(image_findings(self.row, self.review, self.root)[2], [])

    def test_reviews_cannot_be_applied_to_another_assignment(self):
        with patch("audit_production.read_csv", side_effect=[
                [self.row], [dict(self.review, plu="other")]]), \
                patch("audit_production.verify_booklets"), \
                patch("audit_production.audit_entries", return_value=[]):
            with self.assertRaisesRegex(ValueError, "review assignment mismatch"):
                inventory(self.root)

    def test_legacy_policy_approves_entries_and_retains_duplicate_findings_as_history(self):
        with patch("audit_production.read_csv", side_effect=[[self.row], [self.review]]), \
                patch("audit_production.verify_booklets"), \
                patch("audit_production.audit_entries", return_value=[[self.row]]):
            rows, summary = inventory(self.root)
        self.assertEqual(summary["status"], "ALL_LEGACY_ENTRIES_APPROVED")
        self.assertEqual(summary["duplicate_image_groups"], 1)
        self.assertEqual(rows[0]["approval_status"], "APPROVED_LEGACY")
        self.assertEqual(rows[0]["blockers"], "NONE")
        self.assertIn("DUPLICATE_IMAGE_REUSE_REVIEW", rows[0]["historical_findings"])

    def test_legacy_approval_survives_low_resolution_missing_metadata_and_duplicate_flag(self):
        Image.new("RGB", (499, 420), "red").save(self.image)
        self.row.update(width="499", height="420")
        historical_review = dict(self.review, verdict="Approximate")
        _, _, findings = image_findings(self.row, historical_review, self.root)
        with patch("audit_production.read_csv", side_effect=[
                [self.row], [historical_review]]), patch("audit_production.verify_booklets"), \
                patch("audit_production.audit_entries", return_value=[[self.row]]):
            rows, summary = inventory(self.root)
        self.assertIn("IMAGE_QUALITY:BELOW_500_PX", findings)
        self.assertIn("PROVENANCE:MISSING_IMAGE_METADATA", findings)
        self.assertEqual(rows[0]["approval_status"], "APPROVED_LEGACY")
        self.assertEqual(rows[0]["approval_basis"], "OWNER_LEGACY_PREAPPROVED_POLICY")
        self.assertEqual(rows[0]["blockers"], "NONE")
        self.assertIn("DUPLICATE_IMAGE_REUSE_REVIEW", rows[0]["historical_findings"])
        self.assertEqual(summary["entries_with_release_blockers"], 0)

    def test_explicit_rejection_overrides_legacy_policy_but_review_labels_do_not(self):
        reviewed = dict(self.review, verdict="Review")
        with patch("audit_production.read_csv", side_effect=[
                [self.row], [reviewed]]), patch("audit_production.verify_booklets"), \
                patch("audit_production.audit_entries", return_value=[]):
            rows, _ = inventory(self.root)
        self.assertEqual(rows[0]["approval_status"], "APPROVED_LEGACY")
        self.assertEqual(rows[0]["visual_verdict"], "Review")

        rejected = dict(self.review, verdict="Rejected")
        with patch("audit_production.read_csv", side_effect=[
                [self.row], [rejected]]), patch("audit_production.verify_booklets"), \
                patch("audit_production.audit_entries", return_value=[]):
            rows, summary = inventory(self.root)
        self.assertEqual(rows[0]["approval_status"], "NOT_APPROVED")
        self.assertEqual(rows[0]["blockers"], "EXPLICIT_IMAGE_REJECTION")
        self.assertEqual(summary["entries_with_release_blockers"], 1)

    def test_picklist_reconciliation_blockers_block_production(self):
        image_summary = dict(
            status="AUTOMATED_CHECKS_PASSED_REQUIRES_RELEASE_REVIEW",
            entries_with_blockers=0,
        )
        picklist_summary = dict(status="BLOCKED", working_list_review_records=2)

        result = combined_summary(picklist_summary, image_summary)

        self.assertEqual(result["status"], "BLOCKED")
        self.assertEqual(result["picklist_sync"]["working_list_review_records"], 2)

    def test_passing_automated_checks_still_require_release_review(self):
        image_summary = dict(
            status="AUTOMATED_CHECKS_PASSED_REQUIRES_RELEASE_REVIEW",
            entries_with_blockers=0,
        )
        picklist_summary = dict(status="READY_TO_BUILD", working_list_review_records=0)

        self.assertEqual(
            combined_summary(picklist_summary, image_summary)["status"],
            "AUTOMATED_CHECKS_PASSED_REQUIRES_RELEASE_REVIEW",
        )

    def test_quality_rights_provenance_and_duplicate_findings_do_not_become_release_blockers(self):
        result = combined_summary(dict(
            status="READY_TO_BUILD",
            counters=dict(
                image_quality_warnings=1, rights_permission_findings=2,
                provenance_documentation_gaps=3, duplicate_image_reuse_reviews=4,
            ),
        ))
        self.assertEqual(result["status"], "AUTOMATED_CHECKS_PASSED_REQUIRES_RELEASE_REVIEW")
        self.assertEqual(result["gate_statuses"]["image_quality"], "WARNING")
        self.assertEqual(result["gate_statuses"]["rights_permission"], "OWNER_REVIEW_REQUIRED")
        self.assertEqual(result["gate_statuses"]["provenance"], "DOCUMENTATION_GAPS")
        self.assertEqual(result["gate_statuses"]["duplicate_reuse"], "OWNER_REVIEW_REQUIRED")
        self.assertEqual(result["gate_statuses"]["production_release"],
                         "OWNER_AUTHORIZATION_REQUIRED")
        self.assertEqual(result["gate_statuses"]["build"], "NOT_RUN")

    def test_mandatory_identity_or_image_findings_block_independently(self):
        result = combined_summary(dict(
            status="BLOCKED",
            counters=dict(unresolved_product_identity=1, unresolved_image_selection=1),
        ))
        self.assertEqual(result["gate_statuses"]["data_reconciliation"], "BLOCKED")
        self.assertEqual(result["gate_statuses"]["image_selection"], "BLOCKED")
        self.assertEqual(result["gate_statuses"]["production_release"], "BLOCKED")

    def test_unknown_audit_status_fails_closed(self):
        picklist_summary = dict(status="UNEXPECTED")

        self.assertEqual(combined_summary(picklist_summary)["status"], "BLOCKED")

    def test_legacy_image_blockers_do_not_block_current_working_list(self):
        result = combined_summary(
            dict(status="READY_TO_BUILD"), dict(status="BLOCKED", entries_with_blockers=597))
        self.assertEqual(result["status"], "AUTOMATED_CHECKS_PASSED_REQUIRES_RELEASE_REVIEW")
        self.assertEqual(result["legacy_image_audit"]["status"], "BLOCKED")

    def test_default_audit_does_not_depend_on_legacy_booklets(self):
        output = self.root / "reports"
        source = self.root / "sources/plu/PickList.xml"
        source.parent.mkdir(parents=True)
        source.write_bytes(b"fixture")
        digest = hashlib.sha256(source.read_bytes()).hexdigest()
        with patch.object(sys, "argv", ["audit_production.py", "--root", str(self.root),
                                       "--output", str(output)]), \
                patch("audit_production.reconcile_picklist",
                      return_value=([], {}, {}, dict(status="READY_TO_BUILD", source_sha256=digest))), \
                patch("audit_production.inventory", side_effect=AssertionError("legacy audit called")):
            with self.assertRaises(SystemExit) as result:
                main()
        self.assertEqual(result.exception.code, 0)
        self.assertEqual(json.loads((output / "PRODUCTION_READINESS.json").read_text())["status"],
                         "AUTOMATED_CHECKS_PASSED_REQUIRES_RELEASE_REVIEW")

    def test_invalid_source_produces_no_stale_report_in_fresh_output(self):
        source = self.root / "sources/plu/PickList.xml"
        source.parent.mkdir(parents=True)
        source.write_text("<broken")
        old = self.root / "outputs/reports/PRODUCTION_READINESS.json"
        old.parent.mkdir(parents=True)
        old.write_text('{"status": "OLD"}')
        output = self.root / "fresh"
        result = subprocess.run(
            [sys.executable, str(Path(__file__).with_name("audit_production.py")),
             "--root", str(self.root), "--output", str(output)],
            capture_output=True, text=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("cannot parse owner-supplied PickList.xml", result.stderr)
        self.assertEqual(old.read_text(), '{"status": "OLD"}')
        self.assertEqual(list(output.iterdir()), [])

    def test_changed_source_cannot_produce_a_readiness_report(self):
        source = self.root / "sources/plu/PickList.xml"
        source.parent.mkdir(parents=True)
        source.write_bytes(b"changed")
        output = self.root / "reports"
        with patch.object(sys, "argv", ["audit_production.py", "--root", str(self.root),
                                       "--output", str(output)]), \
                patch("audit_production.reconcile_picklist",
                      return_value=([], {}, {}, dict(status="READY_TO_BUILD",
                                                    source_sha256="stale"))):
            with self.assertRaises(SystemExit) as result:
                main()
        self.assertEqual(result.exception.code, 1)
        self.assertFalse((output / "PRODUCTION_READINESS.json").exists())


if __name__ == "__main__":
    unittest.main()
