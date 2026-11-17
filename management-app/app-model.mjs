const CONTROL_CHARACTERS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/;

export function normalizePlu(value) {
  return value == null ? "" : String(value).trim();
}

export function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (quoted) {
      if (character === '"' && text[index + 1] === '"') {
        field += '"';
        index += 1;
      } else if (character === '"') {
        quoted = false;
      } else {
        field += character;
      }
    } else if (character === '"' && field === "") {
      quoted = true;
    } else if (character === ",") {
      row.push(field);
      field = "";
    } else if (character === "\n" || character === "\r") {
      if (character === "\r" && text[index + 1] === "\n") index += 1;
      row.push(field);
      if (row.some((cell) => cell !== "")) rows.push(row);
      row = [];
      field = "";
    } else {
      field += character;
    }
  }
  if (quoted) throw new Error("CSV contains an unterminated quoted field.");
  if (field !== "" || row.length) {
    row.push(field);
    if (row.some((cell) => cell !== "")) rows.push(row);
  }
  if (!rows.length) return [];

  const headers = rows.shift().map((header) => header.replace(/^\uFEFF/, ""));
  return rows.map((cells, rowIndex) => {
    if (cells.length !== headers.length) {
      throw new Error(`CSV row ${rowIndex + 2} has ${cells.length} cells; expected ${headers.length}.`);
    }
    return Object.fromEntries(headers.map((header, index) => [header, cells[index]]));
  });
}

export function validateSearchQuery(query) {
  if (query.length > 100) return "Search text must be 100 characters or fewer.";
  if (CONTROL_CHARACTERS.test(query)) return "Search text contains unsupported control characters.";
  return "";
}

export function validateCatalogue(rows) {
  const errors = [];
  const identities = new Set();
  for (const [index, row] of rows.entries()) {
    if (!normalizePlu(row.plu)) errors.push(`Row ${index + 1}: PLU is required.`);
    if (typeof row.name !== "string" || !row.name.trim()) {
      errors.push(`Row ${index + 1}: product description is required.`);
    }
    if (typeof row.source_id !== "string" || !row.source_id.trim()) {
      errors.push(`Row ${index + 1}: source identity is required.`);
    } else {
      if (identities.has(row.source_id)) errors.push(`Row ${index + 1}: duplicate source identity.`);
      identities.add(row.source_id);
    }
  }
  return errors;
}

export function searchCatalogue(rows, query, category = "") {
  const error = validateSearchQuery(query);
  if (error) return { error, rows: [] };
  const needle = query.trim().toLocaleLowerCase();
  return {
    error: "",
    rows: rows.filter((row) => {
      if (category && row.book !== category && row.section !== category) return false;
      if (!needle) return true;
      return [
        normalizePlu(row.plu),
        row.name,
        row.book,
        row.section,
        row.category,
      ].some((value) => String(value ?? "").toLocaleLowerCase().includes(needle));
    }),
  };
}

export function reviewDetails(row, assignment = {}) {
  const findings = String(row.review_findings || "")
    .split(";").map((finding) => finding.trim()).filter(Boolean);
  return {
    approval: {
      status: row.image_approval_status || "NOT_RECORDED",
      basis: row.image_approval_basis || "Not recorded",
      asset: assignment.image_sha256 || "Not recorded",
      dimensions: row.width && row.height ? `${row.width} × ${row.height} px` : "Not recorded",
    },
    identity: {
      workingList: row.working_list_status || "Not recorded",
      sourceRecord: row.source_record_status || "Not recorded",
      sourceIds: row.source_ids || row.source_id || "",
    },
    quality: findings.filter((finding) =>
      ["BELOW_500_PX", "MULTIFRAME_IMAGE"].includes(finding)),
    rights: {
      status: assignment.rights_verification_status || "Not recorded",
      permission: assignment.provenance?.rights_status || "Not recorded",
      license: assignment.provenance?.license || "Not recorded",
      evidence: assignment.provenance?.rights_evidence || "Not recorded",
      findings: findings.filter((finding) =>
        !finding.startsWith("MISSING_PROVENANCE_")
        && (finding.includes("RIGHTS") || finding.includes("LICENCE")
          || finding.includes("LICENSE") || finding.includes("PERMISSION"))),
    },
    provenance: {
      creator: assignment.provenance?.creator || "Not recorded",
      attribution: assignment.provenance?.attribution || "Not recorded",
      source: assignment.provenance?.source_page_url
        || assignment.provenance?.source_reference || "Not recorded",
      completenessFindings: findings.filter((finding) =>
        finding.startsWith("MISSING_PROVENANCE_")
        || finding.startsWith("PROVENANCE_")
        || finding === "PHOTOGRAPHIC_ORIGIN_NOT_VERIFIED"),
    },
    reuse: {
      sourceId: assignment.reuse_from_source_id || "",
      rationale: assignment.reuse_rationale || "",
      reviewFindings: findings.filter((finding) =>
        finding.includes("REUSE") || finding.includes("DUPLICATE")),
    },
    blockers: String(row.blockers || "").split(";").map((item) => item.trim()).filter(Boolean),
  };
}

export function readinessPresentation(readiness = {}) {
  const blocked = readiness.status === "BLOCKED"
    || readiness.gate_statuses?.production_release === "BLOCKED";
  return {
    blocked,
    status: readiness.status || "NOT_AVAILABLE",
    blockers: readiness.release_blockers || {},
    findings: readiness.owner_review_findings || {},
    counters: readiness.counters || {},
    gates: readiness.gate_statuses || {},
  };
}
