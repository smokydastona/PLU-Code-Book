import test from "node:test";
import assert from "node:assert/strict";

import {
  normalizePlu,
  parseCsv,
  readinessPresentation,
  reviewDetails,
  searchCatalogue,
  validateCatalogue,
  validateSearchQuery,
} from "./app-model.mjs";

const catalogue = [
  { source_id: "apple", plu: "00328", name: "Honeycrisp Apple", book: "Fruit",
    section: "Apples", image_approval_status: "APPROVED",
    image_approval_basis: "EXACT_WORKINGLIST_APPROVAL",
    review_findings: "BELOW_500_PX; DUPLICATE_IMAGE_REUSE_REVIEW; MISSING_PROVENANCE_LICENSE_URL; IMAGE_RIGHTS_NOT_VERIFIED; SUPPORTED_IMAGE_LICENCE_REQUIRED",
    width: "640", height: "480", blockers: "" },
  { source_id: "pear", plu: "00401", name: "Green Pear", book: "Fruit",
    section: "Pears", image_approval_status: "APPROVED_LEGACY", blockers: "" },
];

test("PLU strings retain leading zeros through parsing and normalization", () => {
  assert.equal(parseCsv("plu,name\n00328,Apple\n")[0].plu, "00328");
  assert.equal(normalizePlu("00328"), "00328");
});

test("catalogue search covers PLU, description, book and section filters", () => {
  assert.deepEqual(searchCatalogue(catalogue, "00328").rows.map((row) => row.source_id), ["apple"]);
  assert.deepEqual(searchCatalogue(catalogue, "honey").rows.map((row) => row.source_id), ["apple"]);
  assert.deepEqual(searchCatalogue(catalogue, "", "Pears").rows.map((row) => row.source_id), ["pear"]);
  assert.deepEqual(searchCatalogue(catalogue, "fruit").rows.length, 2);
});

test("CSV parser handles quoted commas, escaped quotes and embedded newlines", () => {
  const rows = parseCsv('plu,name,note\n0001,"Apple, Red","line 1\nline ""2"""\n');
  assert.deepEqual(rows[0], { plu: "0001", name: "Apple, Red", note: 'line 1\nline "2"' });
  assert.throws(() => parseCsv('plu,name\n1,"unfinished'), /unterminated quoted field/);
});

test("catalogue and search validation return actionable errors", () => {
  assert.deepEqual(validateCatalogue([{ source_id: "x", plu: "", name: "" }]), [
    "Row 1: PLU is required.",
    "Row 1: product description is required.",
  ]);
  assert.deepEqual(validateCatalogue([{ plu: "0001", name: "Apple" }]), [
    "Row 1: source identity is required.",
  ]);
  assert.match(validateSearchQuery("x".repeat(101)), /100 characters/);
  assert.match(validateSearchQuery("bad\u0000input"), /control characters/);
  assert.deepEqual(searchCatalogue(catalogue, "x".repeat(101)).rows, []);
});

test("approval, quality, rights, provenance and reuse remain separate", () => {
  const details = reviewDetails(catalogue[0], {
    image_sha256: "asset-hash",
    rights_verification_status: "NOT_VERIFIED",
    provenance: { license: "", rights_status: "NOT_VERIFIED" },
    reuse_from_source_id: "pear",
  });
  assert.equal(details.approval.status, "APPROVED");
  assert.equal(details.approval.basis, "EXACT_WORKINGLIST_APPROVAL");
  assert.equal(details.approval.dimensions, "640 × 480 px");
  assert.deepEqual(details.quality, ["BELOW_500_PX"]);
  assert.equal(details.rights.status, "NOT_VERIFIED");
  assert.equal(details.rights.license, "Not recorded");
  assert.deepEqual(details.rights.findings,
    ["IMAGE_RIGHTS_NOT_VERIFIED", "SUPPORTED_IMAGE_LICENCE_REQUIRED"]);
  assert.deepEqual(details.provenance.completenessFindings, ["MISSING_PROVENANCE_LICENSE_URL"]);
  assert.equal(details.reuse.sourceId, "pear");
  assert.deepEqual(details.blockers, []);
});

test("blocked release presentation never reports production authorization", () => {
  const view = readinessPresentation({
    status: "BLOCKED",
    release_blockers: { EXPLICIT_IMAGE_REJECTION: 1 },
    owner_review_findings: { DUPLICATE_IMAGE_REUSE_REVIEW: 2 },
  });
  assert.equal(view.blocked, true);
  assert.deepEqual(view.blockers, { EXPLICIT_IMAGE_REJECTION: 1 });
  assert.deepEqual(view.findings, { DUPLICATE_IMAGE_REUSE_REVIEW: 2 });
  assert.equal(readinessPresentation({ status: "AUTOMATED_CHECKS_PASSED_REQUIRES_RELEASE_REVIEW" }).blocked,
    false);
});
