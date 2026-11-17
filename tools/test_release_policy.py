import unittest

from release_policy import (
    COUNTER_DEFINITIONS,
    explicit_image_rejection,
    finding_category,
    known_permission_restriction,
    release_status,
)


class ReleasePolicyChecks(unittest.TestCase):
    def test_only_explicit_rejection_and_supersession_override_approval(self):
        self.assertFalse(explicit_image_rejection({"visual_approval_status": "APPROVED"}))
        self.assertFalse(explicit_image_rejection({"visual_approval_status": "REVIEW"}))
        self.assertFalse(explicit_image_rejection({"visual_approval_status": "NOT_APPROVED"}))
        self.assertTrue(explicit_image_rejection({"visual_approval_status": "REJECTED"}))
        self.assertTrue(explicit_image_rejection({"production_status": "SUPERSEDED"}))
        self.assertTrue(explicit_image_rejection({}, {
            "historical_review": {"verdict": "Rejected"}}))
        self.assertFalse(explicit_image_rejection({}, {
            "historical_review": {"verdict": "Review"}}))

    def test_blank_or_missing_licence_metadata_is_not_a_known_restriction(self):
        self.assertFalse(known_permission_restriction({"provenance": {"license_url": ""}}))
        self.assertFalse(known_permission_restriction({"rights_verification_status": "NOT_VERIFIED"}))
        self.assertTrue(known_permission_restriction({
            "rights_verification_status": "REQUIRED_PERMISSION_NOT_GRANTED"}))

    def test_release_status_uses_mandatory_gates_only(self):
        self.assertEqual(release_status(False), "READY_TO_BUILD")
        self.assertEqual(release_status(True), "BLOCKED")
        self.assertEqual(release_status(False, reconciliation_review_records=1), "BLOCKED")
        self.assertEqual(release_status(False, review_decisions=1), "BLOCKED")

    def test_report_categories_have_independently_documented_meanings(self):
        expected = {
            "unresolved_source_record_reconciliation",
            "unresolved_product_identity",
            "unresolved_image_selection",
            "image_quality_warnings",
            "rights_permission_findings",
            "provenance_documentation_gaps",
            "duplicate_image_reuse_reviews",
            "classification_section_reviews",
        }
        self.assertEqual(set(COUNTER_DEFINITIONS), expected)
        self.assertIn("does not revoke approval",
                      COUNTER_DEFINITIONS["duplicate_image_reuse_reviews"])
        self.assertEqual(finding_category("MISSING_PROVENANCE_LICENSE_URL"),
                         "provenance_documentation_gaps")
        self.assertEqual(finding_category("BELOW_500_PX"), "image_quality_warnings")


if __name__ == "__main__":
    unittest.main()
