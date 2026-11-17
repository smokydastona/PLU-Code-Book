"""Shared approval, finding and release-gate rules for catalogue outputs."""

IMAGE_FINDING_CATEGORIES = {
    "BELOW_500_PX": "image_quality_warnings",
    "MULTIFRAME_IMAGE": "image_quality_warnings",
    "PHOTOGRAPHIC_ORIGIN_NOT_VERIFIED": "provenance_documentation_gaps",
    "IMAGE_RIGHTS_NOT_VERIFIED": "rights_permission_findings",
    "SUPPORTED_IMAGE_LICENCE_REQUIRED": "rights_permission_findings",
    "MISSING_IMAGE_PROVENANCE_AND_RIGHTS_VERIFICATION": "provenance_documentation_gaps",
    "PROVENANCE_CHECKSUM_MISMATCH": "provenance_documentation_gaps",
    "PROVENANCE_ASSIGNMENT_MISMATCH": "provenance_documentation_gaps",
    "RECORDED_DIMENSIONS_MISMATCH": "provenance_documentation_gaps",
}

COUNTER_DEFINITIONS = {
    "unresolved_source_record_reconciliation":
        "Current source identities requiring reconciliation plus removed baseline identities.",
    "unresolved_product_identity":
        "Included WorkingList entries flagged for a source/product identity decision.",
    "unresolved_image_selection":
        "Included entries without an approved image or approved code-only selection for the exact identity.",
    "image_quality_warnings":
        "Non-blocking image-quality finding occurrences, including low resolution and multiframe files.",
    "rights_permission_findings":
        "Unresolved rights/licence findings; missing optional URLs alone are not release blockers.",
    "provenance_documentation_gaps":
        "Missing or inconsistent provenance evidence, counted independently from approval status.",
    "duplicate_image_reuse_reviews":
        "Fingerprint groups requiring owner review; detection does not revoke approval or block release.",
    "classification_section_reviews":
        "Included entries missing a category or output section required to build the catalogue.",
}

IMAGE_RELEASE_BLOCKERS = {
    "IMAGE_APPROVAL_REQUIRED",
    "IMAGE_ASSET_CHANGED_REVIEW",
    "IMAGE_FILE_UNUSABLE",
    "EXPLICIT_IMAGE_REJECTION",
    "PRODUCT_IMAGE_REVIEW_REQUIRED",
    "RIGHTS_PERMISSION_REVIEW_REQUIRED",
}

LIBRARY_EXCLUSION_BLOCKERS = IMAGE_RELEASE_BLOCKERS | {
    "WORKING_LIST_REVIEW_REQUIRED",
    "SOURCE_RECORD_CHANGED_REVIEW",
    "APPROVED_IMAGE_REQUIRED",
}

EXPLICITLY_UNAPPROVED_STATUSES = {
    "REJECTED",
    "SUPERSEDED",
    "REVOKED",
}

KNOWN_PERMISSION_RESTRICTIONS = {
    "REQUIRED_PERMISSION_NOT_GRANTED",
    "PERMISSION_DENIED",
    "USAGE_PROHIBITED",
    "RESTRICTED",
}


def explicit_image_rejection(assignment, legacy_record=None):
    if any(str(assignment.get(field, "")).strip().upper() in EXPLICITLY_UNAPPROVED_STATUSES
           for field in ("visual_approval_status", "production_status", "approval_status")):
        return True
    return bool(legacy_record and
                str(legacy_record["historical_review"].get("verdict", "")).strip().casefold()
                == "rejected")


def known_permission_restriction(assignment):
    provenance = assignment.get("provenance") or {}
    return any(str(value).strip().upper() in KNOWN_PERMISSION_RESTRICTIONS for value in (
        assignment.get("rights_verification_status"),
        assignment.get("rights_status"),
        provenance.get("rights_verification_status"),
        provenance.get("rights_status"),
    ))


def finding_category(finding):
    if finding.startswith(("MISSING_PROVENANCE_", "PROVENANCE_")):
        return "provenance_documentation_gaps"
    return IMAGE_FINDING_CATEGORIES.get(finding)


def release_status(blockers, reconciliation_review_records=0, review_decisions=0):
    return "BLOCKED" if blockers or reconciliation_review_records or review_decisions else "READY_TO_BUILD"
