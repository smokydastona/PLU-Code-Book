import {
  parseCsv,
  readinessPresentation,
  reviewDetails,
  searchCatalogue,
  validateCatalogue,
} from "./app-model.mjs";

const byId = (id) => document.getElementById(id);
const status = byId("app-status");
const results = byId("catalogue-results");
const details = byId("product-details");
const queryInput = byId("catalogue-search");
const categoryInput = byId("category-filter");

let catalogue = [];
let reconciliation = [];
let assignments = new Map();
let selectedId = "";

function csvUrl(file) {
  return new URL(`../outputs/reports/${file}`, import.meta.url);
}

async function fetchText(url, label) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`${label} could not be loaded (HTTP ${response.status}).`);
  return response.text();
}

async function fetchJson(url, label) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`${label} could not be loaded (HTTP ${response.status}).`);
  return response.json();
}

function setMessage(element, message, kind = "info") {
  element.textContent = message;
  element.className = `message message-${kind}`;
}

function appendValue(parent, label, value) {
  const row = document.createElement("div");
  row.className = "detail-row";
  const term = document.createElement("dt");
  term.textContent = label;
  const description = document.createElement("dd");
  description.textContent = value || "Not recorded";
  row.append(term, description);
  parent.append(row);
}

function appendSection(parent, heading, entries, emptyMessage = "None recorded.") {
  const section = document.createElement("section");
  section.className = "detail-section";
  const title = document.createElement("h3");
  title.textContent = heading;
  section.append(title);
  if (!entries.length) {
    const empty = document.createElement("p");
    empty.className = "muted";
    empty.textContent = emptyMessage;
    section.append(empty);
  } else {
    const list = document.createElement("ul");
    for (const entry of entries) {
      const item = document.createElement("li");
      item.textContent = entry;
      list.append(item);
    }
    section.append(list);
  }
  parent.append(section);
}

function matchingSourceRecords(row) {
  const workingId = row.source_id;
  const sourceIds = new Set(String(row.source_ids || row.source_id).split(";"));
  return reconciliation.filter((source) =>
    source.working_id === workingId || sourceIds.has(source.source_id));
}

function imageSource(assignment) {
  const path = assignment?.image;
  if (typeof path !== "string" || !path.startsWith("data/") || path.includes("..")) return "";
  return new URL(`../${path}`, import.meta.url).href;
}

function renderImagePanel(row, assignment, view) {
  const panel = document.createElement("section");
  panel.className = "detail-section";
  const heading = document.createElement("h3");
  heading.textContent = "Selected image";
  panel.append(heading);
  const source = imageSource(assignment);
  if (source && row.image) {
    const image = document.createElement("img");
    image.src = source;
    image.alt = `Approved image for PLU ${row.plu}: ${row.name}`;
    image.loading = "lazy";
    image.addEventListener("error", () => {
      image.remove();
      const message = document.createElement("p");
      message.className = "message message-error";
      message.textContent = "The selected image asset could not be loaded.";
      panel.append(message);
    }, { once: true });
    panel.append(image);
  } else {
    const empty = document.createElement("p");
    empty.className = "muted";
    empty.textContent = "No image is selected for this entry.";
    panel.append(empty);
  }
  const approval = document.createElement("dl");
  appendValue(approval, "Approval", view.approval.status);
  appendValue(approval, "Approval basis", view.approval.basis);
  appendValue(approval, "Dimensions", view.approval.dimensions);
  appendValue(approval, "Asset SHA-256", view.approval.asset);
  panel.append(approval);
  return panel;
}

function renderReconciliation(row) {
  const section = document.createElement("section");
  section.className = "detail-section";
  const heading = document.createElement("h3");
  heading.textContent = "Original source records";
  section.append(heading);
  const sourceRows = matchingSourceRecords(row);
  if (!sourceRows.length) {
    const empty = document.createElement("p");
    empty.className = "message message-warning";
    empty.textContent = "No matching reconciliation rows were found.";
    section.append(empty);
    return section;
  }
  const table = document.createElement("table");
  const head = document.createElement("thead");
  const header = document.createElement("tr");
  for (const label of ["Original PLU", "Original description", "Department",
    "Working description", "Disposition", "Review status", "Source identity"]) {
    const cell = document.createElement("th");
    cell.scope = "col";
    cell.textContent = label;
    header.append(cell);
  }
  head.append(header);
  table.append(head);
  const body = document.createElement("tbody");
  for (const source of sourceRows) {
    const line = document.createElement("tr");
    for (const value of [
      source.plu,
      source.original_name,
      source.department,
      source.working_name,
      source.disposition,
      source.review_status,
      source.source_id,
    ]) {
      const cell = document.createElement("td");
      cell.textContent = value || "Not recorded";
      line.append(cell);
    }
    body.append(line);
  }
  table.append(body);
  section.append(table);
  return section;
}

function renderDetails(row) {
  details.replaceChildren();
  const heading = document.createElement("h2");
  heading.textContent = `${row.name} — PLU ${row.plu}`;
  details.append(heading);
  const assignment = assignments.get(row.source_id) || {};
  const view = reviewDetails(row, assignment);
  const statusGrid = document.createElement("dl");
  statusGrid.className = "status-grid";
  appendValue(statusGrid, "Product identity", view.identity.workingList);
  appendValue(statusGrid, "Source record", view.identity.sourceRecord);
  appendValue(statusGrid, "Source identities", view.identity.sourceIds);
  appendValue(statusGrid, "Category", `${row.book || "Unclassified"} / ${row.section || "No section"}`);
  details.append(statusGrid);

  const imagePanel = renderImagePanel(row, assignment, view);
  const columns = document.createElement("div");
  columns.className = "detail-columns";
  columns.append(imagePanel);

  const evidence = document.createElement("div");
  evidence.className = "evidence-columns";
  const rights = document.createElement("section");
  rights.className = "detail-section";
  const rightsHeading = document.createElement("h3");
  rightsHeading.textContent = "Rights and permission";
  rights.append(rightsHeading);
  const rightsDetails = document.createElement("dl");
  appendValue(rightsDetails, "Verification status", view.rights.status);
  appendValue(rightsDetails, "Permission status", view.rights.permission);
  appendValue(rightsDetails, "Recorded licence", view.rights.license);
  appendValue(rightsDetails, "Evidence", view.rights.evidence);
  rights.append(rightsDetails);
  appendSection(rights, "Rights findings", view.rights.findings);
  evidence.append(rights);

  const provenance = document.createElement("section");
  provenance.className = "detail-section";
  const provenanceHeading = document.createElement("h3");
  provenanceHeading.textContent = "Provenance and documentation";
  provenance.append(provenanceHeading);
  const provenanceDetails = document.createElement("dl");
  appendValue(provenanceDetails, "Creator", view.provenance.creator);
  appendValue(provenanceDetails, "Attribution", view.provenance.attribution);
  appendValue(provenanceDetails, "Source", view.provenance.source);
  provenance.append(provenanceDetails);
  appendSection(provenance, "Documentation gaps", view.provenance.completenessFindings);
  evidence.append(provenance);
  columns.append(evidence);
  details.append(columns);

  const findings = document.createElement("div");
  findings.className = "evidence-columns";
  appendSection(findings, "Image quality findings", view.quality);
  appendSection(findings, "Duplicate and reuse review", view.reuse.reviewFindings);
  appendSection(findings, "Release blockers", view.blockers,
    "No mandatory blockers are recorded for this entry.");
  if (view.reuse.sourceId) {
    const reuse = document.createElement("section");
    reuse.className = "detail-section";
    const reuseHeading = document.createElement("h3");
    reuseHeading.textContent = "Reuse relationship";
    reuse.append(reuseHeading);
    const reuseDetails = document.createElement("dl");
    appendValue(reuseDetails, "Reused from source identity", view.reuse.sourceId);
    appendValue(reuseDetails, "Recorded rationale", view.reuse.rationale);
    reuse.append(reuseDetails);
    findings.append(reuse);
  }
  details.append(findings, renderReconciliation(row));
}

function renderResults() {
  const result = searchCatalogue(catalogue, queryInput.value, categoryInput.value);
  results.replaceChildren();
  if (result.error) {
    setMessage(status, result.error, "error");
    setMessage(results, "Correct the search input to continue.", "warning");
    details.replaceChildren();
    return;
  }
  setMessage(status, `${result.rows.length} catalogue entries found.`, "info");
  if (!result.rows.length) {
    setMessage(results, "No entries match this search.", "empty");
    details.replaceChildren();
    return;
  }
  const list = document.createElement("ul");
  list.className = "result-list";
  for (const row of result.rows.slice(0, 100)) {
    const item = document.createElement("li");
    const button = document.createElement("button");
    button.type = "button";
    button.className = "result-button";
    button.setAttribute("aria-pressed", String(row.source_id === selectedId));
    button.textContent = `PLU ${row.plu} — ${row.name} (${row.book || "Unclassified"})`;
    button.addEventListener("click", () => {
      selectedId = row.source_id;
      renderResults();
      renderDetails(row);
    });
    item.append(button);
    list.append(item);
  }
  results.append(list);
  if (result.rows.length > 100) {
    const note = document.createElement("p");
    note.className = "muted";
    note.textContent = "Showing the first 100 matches. Refine the search to narrow results.";
    results.append(note);
  }
  const selected = result.rows.find((row) => row.source_id === selectedId);
  if (selected) renderDetails(selected);
  else renderDetails(result.rows[0]);
}

function renderDashboard(readiness) {
  const panel = byId("release-dashboard");
  panel.replaceChildren();
  const view = readinessPresentation(readiness);
  const heading = document.createElement("h2");
  heading.id = "release-heading";
  heading.textContent = "Production release status";
  panel.append(heading);
  const state = document.createElement("p");
  state.className = view.blocked ? "status status-blocked" : "status status-review";
  state.setAttribute("role", "status");
  state.textContent = view.blocked
    ? `BLOCKED — ${view.status}`
    : `${view.status} — owner authorization is still required`;
  panel.append(state);

  const gates = document.createElement("dl");
  gates.className = "status-grid";
  for (const [name, value] of Object.entries(view.gates)) appendValue(gates, name, value);
  panel.append(gates);
  appendSection(panel, "Mandatory release blockers",
    Object.entries(view.blockers).map(([name, count]) => `${name}: ${count}`),
    "No mandatory blockers are listed. This is not a production approval.");
  appendSection(panel, "Owner-review findings",
    Object.entries(view.findings).map(([name, count]) => `${name}: ${count}`));
  appendSection(panel, "Independent counters",
    Object.entries(view.counters).map(([name, count]) => `${name}: ${count}`));
}

async function start() {
  setMessage(status, "Loading catalogue and audit reports…");
  try {
    const [catalogueText, reconciliationText, assignmentData, readiness] = await Promise.all([
      fetchText(csvUrl("PICKLIST_SYNC.csv"), "Catalogue"),
      fetchText(csvUrl("WORKING_LIST_RECONCILIATION.csv"), "Reconciliation report"),
      fetchJson(new URL("../data/picklist_assignments.json", import.meta.url), "Image assignments"),
      fetchJson(new URL("../outputs/reports/PRODUCTION_READINESS.json", import.meta.url),
        "Production readiness"),
    ]);
    catalogue = parseCsv(catalogueText);
    reconciliation = parseCsv(reconciliationText);
    const errors = validateCatalogue(catalogue);
    if (errors.length) throw new Error(`Catalogue validation failed: ${errors.slice(0, 5).join(" ")}`);
    assignments = new Map((assignmentData.assignments || []).map((assignment) =>
      [assignment.source_id, assignment]));
    const categories = [...new Set(catalogue.flatMap((row) =>
      [row.book, row.section]).filter(Boolean))].sort();
    for (const category of categories) {
      const option = document.createElement("option");
      option.value = category;
      option.textContent = category;
      categoryInput.append(option);
    }
    renderDashboard(readiness);
    renderResults();
  } catch (error) {
    setMessage(status, `Application data could not be validated: ${error.message}`, "error");
    setMessage(results, "Catalogue is unavailable until its reports can be loaded.", "error");
    details.replaceChildren();
    byId("release-dashboard").replaceChildren();
  }
}

queryInput.addEventListener("input", renderResults);
categoryInput.addEventListener("change", renderResults);
start();
