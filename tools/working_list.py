"""Rebuild the reviewed WorkingList from the immutable owner PickList."""
import copy
import csv
import hashlib
import json
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path

from picklist_catalog import (
    read_assignments, read_picklist, read_record_baseline, source_identity_id,
)


def _sha(data):
    return hashlib.sha256(data).hexdigest()


def _write_csv(path, rows, fields):
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=fields, extrasaction="ignore", lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def _identity(item):
    return dict(source_id=item["source_id"], plu=item["plu"], name=item["name"],
                department=item["department"], occurrence=item["occurrence"],
                visible=item["visible"], deleted=item["deleted"])


def _working_id(source_ids, plu, department, name, decision_id):
    if len(source_ids) == 1 and not decision_id:
        return source_ids[0]
    key = json.dumps([source_ids, plu, department, name, decision_id],
                     ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(("working:" + key).encode("utf-8")).hexdigest()


def _set_description(item, name):
    description = item.find("./Languages/Language[@LCID='0409']/Description")
    if description is None:
        raise ValueError(f"PLU {item.get('UPC')!r}: missing English description")
    description.text = name


def _validate_decision_registry(registry):
    if not isinstance(registry, dict) or registry.get("schema_version") != 1:
        raise ValueError("unsupported WorkingList decision registry schema")
    if not isinstance(registry.get("baseline"), list) or not isinstance(registry.get("decisions"), list):
        raise ValueError("WorkingList decision registry requires baseline and decisions lists")
    baseline_sha = registry.get("baseline_source_sha256")
    if (not isinstance(baseline_sha, str) or len(baseline_sha) != 64 or
            any(char not in "0123456789abcdef" for char in baseline_sha)):
        raise ValueError("WorkingList decision registry has an invalid baseline source SHA-256")
    baseline_by_id = {}
    for row in registry["baseline"]:
        if not isinstance(row, dict):
            raise ValueError("WorkingList baseline records must be objects")
        source_id = row.get("source_id")
        expected_id = source_identity_id([
            row.get("plu"), row.get("name"), row.get("department"), row.get("occurrence")
        ]) if all(key in row for key in ("plu", "name", "department", "occurrence")) else ""
        if not isinstance(source_id, str) or len(source_id) != 64 or source_id != expected_id:
            raise ValueError("WorkingList baseline contains a missing or invalid source_id")
        if source_id in baseline_by_id:
            raise ValueError(f"duplicate source identity in WorkingList baseline: {source_id}")
        if (not isinstance(row.get("plu"), str) or not row["plu"] or
                not isinstance(row.get("name"), str) or
                not isinstance(row.get("department"), str) or
                type(row.get("occurrence")) is not int or row["occurrence"] < 1 or
                not isinstance(row.get("visible"), bool) or not isinstance(row.get("deleted"), bool)):
            raise ValueError(f"invalid WorkingList baseline record: {source_id}")
        baseline_by_id[source_id] = row
    baseline_ids = set(baseline_by_id)

    decision_ids, claimed_source_ids = set(), set()
    for decision in registry["decisions"]:
        if not isinstance(decision, dict):
            raise ValueError("WorkingList decisions must be objects")
        plu, mode = decision.get("plu"), decision.get("mode")
        canonical_name = decision.get("canonical_name")
        if not isinstance(plu, str) or not plu:
            raise ValueError("WorkingList decision requires a text PLU")
        if plu in decision_ids:
            raise ValueError(f"conflicting duplicate WorkingList decision for PLU {plu!r}")
        decision_ids.add(plu)
        if not isinstance(mode, str) or mode not in {
                "MERGE", "SELECT", "STANDARDIZE", "PENDING"}:
            raise ValueError(f"PLU {plu}: invalid decision mode {mode!r}")
        if (mode != "PENDING" and
                (not isinstance(canonical_name, str) or not canonical_name.strip())):
            raise ValueError(f"PLU {plu}: canonical_name must be a non-empty string")
        source_ids = decision.get("source_ids")
        selected_ids = decision.get("selected_source_ids")
        if (not isinstance(source_ids, list) or not source_ids or
                any(not isinstance(value, str) or len(value) != 64 for value in source_ids) or
                len(source_ids) != len(set(source_ids))):
            raise ValueError(f"PLU {plu}: source_ids must be a non-empty list of unique source IDs")
        if selected_ids is None:
            selected_ids = source_ids
        if (not isinstance(selected_ids, list) or (mode != "PENDING" and not selected_ids) or
                any(not isinstance(value, str) for value in selected_ids) or
                len(selected_ids) != len(set(selected_ids)) or
                not set(selected_ids).issubset(source_ids)):
            raise ValueError(f"PLU {plu}: selected_source_ids must be a unique subset of source_ids")
        if not set(source_ids).issubset(baseline_ids):
            raise ValueError(f"PLU {plu}: source_ids contain identities missing from the registry baseline")
        if any(baseline_by_id[source_id]["plu"] != plu for source_id in selected_ids):
            raise ValueError(f"PLU {plu}: selected source IDs must match the decision PLU")
        condition = decision.get("condition")
        if not isinstance(condition, str) or not condition.strip():
            raise ValueError(f"PLU {plu}: decision requires an auditable condition")
        classification = decision.get("classification")
        if classification is not None and (
                not isinstance(classification, str) or classification not in {
                    "Merge", "Standardize", "Investigate", "Verify"}):
            raise ValueError(f"PLU {plu}: unsupported classification {classification!r}")
        excluded_disposition = decision.get("excluded_source_disposition")
        if excluded_disposition is not None and (
                not isinstance(excluded_disposition, str) or excluded_disposition not in {
                    "EXCLUDED_CONFLICTING_ASSOCIATION", "EXCLUDED_EQUIVALENCE_NOT_CONFIRMED",
                    "EXCLUDED_DUPLICATE_PLU_REMOVED"}):
            raise ValueError(f"PLU {plu}: unsupported excluded-source disposition")
        conflict = claimed_source_ids.intersection(source_ids)
        if conflict:
            raise ValueError(f"source identity appears in conflicting decisions: {min(conflict)}")
        claimed_source_ids.update(source_ids)


def _validate_source_accounting(source_items, output_groups, source_disposition):
    source_ids = [item["source_id"] for item in source_items]
    source_id_set = set(source_ids)
    if len(source_id_set) != len(source_ids):
        raise ValueError("PickList contains duplicate logical source identities")
    if set(source_disposition) != source_id_set:
        raise ValueError("source audit does not contain exactly one disposition per source identity")

    working_source_ids = set()
    for group in output_groups:
        identities = group["source_ids"]
        if (not identities or len(identities) != len(set(identities)) or
                not set(identities).issubset(source_id_set)):
            raise ValueError("WorkingList entry has missing or invalid source identities")
        overlap = working_source_ids.intersection(identities)
        if overlap:
            raise ValueError(f"source identity belongs to multiple WorkingList entries: {min(overlap)}")
        working_source_ids.update(identities)

    for source_id in source_id_set:
        excluded = source_disposition[source_id]["disposition"].startswith("EXCLUDED_")
        if excluded and source_id in working_source_ids:
            raise ValueError(f"excluded source identity appears in WorkingList: {source_id}")
        if not excluded and source_id not in working_source_ids:
            raise ValueError(f"source identity has inconsistent WorkingList disposition: {source_id}")


def rebuild_working_list(root, audit_output):
    root = Path(root)
    audit_output = Path(audit_output)
    source_path = root / "sources/plu/PickList.xml"
    working_path = root / "sources/plu/WorkingList.xml"
    registry_path = root / "data/working_list_decisions.json"
    source_items, source_digest = read_picklist(source_path)
    record_baseline, record_baseline_digest, _ = read_record_baseline(
        root / "data/picklist_record_baseline.json")
    source_tree = ET.parse(source_path)
    source_root = source_tree.getroot()
    source_nodes = source_root.findall("./Items/Item")
    if len(source_nodes) != len(source_items):
        raise ValueError("PickList record positions differ from parsed records")
    registry = json.loads(registry_path.read_text(encoding="utf-8"))
    _validate_decision_registry(registry)
    assignments = read_assignments(root / "data/picklist_assignments.json")
    baseline = {row["source_id"]: row for row in registry["baseline"]}
    if len(baseline) != len(registry["baseline"]):
        raise ValueError("duplicate source identity in WorkingList baseline")

    current_by_id = {item["source_id"]: item for item in source_items}
    current_by_plu = defaultdict(list)
    for item in source_items:
        current_by_plu[item["plu"]].append(item)

    source_status = {}
    source_record_status = {}
    for item in source_items:
        saved_record = record_baseline.get(item["source_id"])
        if saved_record is None:
            record_status = "NEW_SOURCE_RECORD_REVIEW"
        elif saved_record["record_sha256"] != item["record_sha256"]:
            record_status = "SOURCE_RECORD_CHANGED_REVIEW"
        else:
            record_status = "MATCHED_RECORD_BASELINE"
        source_record_status[item["source_id"]] = record_status
        old = baseline.get(item["source_id"])
        if old is not None:
            status = "UNCHANGED_SOURCE_IDENTITY"
            if (old.get("visible"), old.get("deleted")) != (item["visible"], item["deleted"]):
                status = "SOURCE_VISIBILITY_CHANGED_REVIEW"
        else:
            same_code = [row for row in registry["baseline"]
                         if (row["plu"], row["department"]) == (item["plu"], item["department"])]
            same_description = [row for row in same_code if row["name"] == item["name"]]
            if same_description:
                status = "NEW_DUPLICATE_IDENTITY_REVIEW"
            elif same_code:
                status = "CHANGED_SOURCE_IDENTITY_REVIEW"
            else:
                status = "NEW_SOURCE_IDENTITY_REVIEW"
        if (record_status == "SOURCE_RECORD_CHANGED_REVIEW" or
                old is not None and record_status == "NEW_SOURCE_RECORD_REVIEW"):
            status = record_status
        source_status[item["source_id"]] = status

    source_disposition = {item["source_id"]: dict(
        disposition="UNCHANGED_ORIGINAL_RECORD", working_id="", working_name=item["name"],
        decision_id="", status=source_status[item["source_id"]],
        record_sha256=item["record_sha256"],
        record_baseline_sha256=(record_baseline.get(item["source_id"]) or {}).get(
            "record_sha256", ""),
        record_status=source_record_status[item["source_id"]], details="")
        for item in source_items}
    output_groups = []
    applied_decisions = []
    handled_decision_ids = set()

    for decision in registry["decisions"]:
        decision_id = str(decision["plu"])
        source_ids = decision["source_ids"]
        if decision.get("mode") == "PENDING":
            applied_decisions.append(dict(
                decision_id=decision_id, status="PENDING_OWNER_CONDITION",
                details=decision["condition"]))
            for source_id in source_ids:
                if source_id in source_disposition:
                    source_disposition[source_id]["status"] = "WORKING_LIST_REVIEW_REQUIRED"
                    source_disposition[source_id]["details"] = decision["condition"]
            continue
        matching = [current_by_id.get(source_id) for source_id in source_ids]
        complete = all(item is not None and item["included"] for item in matching)
        records_unchanged = all(
            item is not None and source_record_status.get(item["source_id"]) == "MATCHED_RECORD_BASELINE"
            for item in matching)
        if not complete or not records_unchanged:
            decision_status = ("DECISION_SOURCE_RECORD_CHANGED_REVIEW"
                               if complete and not records_unchanged else
                               "DECISION_SOURCE_MISSING_OR_CHANGED")
            applied_decisions.append(dict(
                decision_id=decision_id, status=decision_status,
                details=("A pinned source record changed; decision not applied."
                         if complete and not records_unchanged else
                         "Expected exact source identities are not all present and active; decision not applied.")))
            for item in current_by_plu.get(str(decision["plu"]), []):
                if source_disposition[item["source_id"]]["status"] == "UNCHANGED_SOURCE_IDENTITY":
                    source_disposition[item["source_id"]]["status"] = "WORKING_LIST_REVIEW_REQUIRED"
                source_disposition[item["source_id"]]["details"] = (
                    "Approved decision no longer matches its complete source identity set.")
            continue

        selected_ids = decision.get("selected_source_ids", source_ids)
        selected = [current_by_id[source_id] for source_id in selected_ids]
        extras = [item for item in current_by_plu.get(str(decision["plu"]), [])
                  if item["source_id"] not in source_ids]
        status = ("APPLIED_WITH_ADDITIONAL_SOURCE_REVIEW" if extras else
                  "APPLIED_TO_EXACT_SOURCE_IDENTITIES")
        applied_decisions.append(dict(
            decision_id=decision_id, status=status,
            details=decision.get("condition", "")))

        working_name = decision["canonical_name"]
        primary_id = selected[0]["source_id"]
        books = {assignments.get(source_id, {}).get("book") for source_id in selected_ids}
        classification_source_id = (primary_id if len(books) == 1 and
                                     None not in books else "")
        group = dict(
            source_ids=selected_ids, source_items=selected, working_name=working_name,
            decision_id=decision_id, classification_source_id=classification_source_id,
            status=status,
            details=decision.get("condition", "Product identity changed; image review required."),
        )
        output_groups.append(group)
        handled_decision_ids.update(source_ids)
        working_id = _working_id(selected_ids, selected[0]["plu"],
                                 selected[0]["department"], working_name, decision_id)
        for item in selected:
            source_disposition[item["source_id"]].update(
                disposition=("MERGED_INTO_WORKING_ENTRY" if len(selected_ids) > 1 else
                             "STANDARDIZED_WORKING_ENTRY"),
                working_id=working_id, working_name=working_name, decision_id=decision_id,
                status=status, details=decision.get("condition", ""))
        for source_id in source_ids:
            if source_id in selected_ids:
                continue
            source_disposition[source_id].update(
                disposition=decision.get("excluded_source_disposition",
                                         "EXCLUDED_CONFLICTING_ASSOCIATION"),
                working_id="",
                working_name=decision["canonical_name"], decision_id=decision_id,
                status=status, details=decision.get("condition", ""))
        if extras:
            for item in extras:
                source_disposition[item["source_id"]]["status"] = "NEW_OR_AMBIGUOUS_PLU_IDENTITY_REVIEW"
                source_disposition[item["source_id"]]["details"] = (
                    "A new identity shares a PLU with an approved WorkingList decision.")

    for item in source_items:
        if item["source_id"] in handled_decision_ids or any(
                item["source_id"] in decision["source_ids"] for decision in registry["decisions"]):
            continue
        output_groups.append(dict(
            source_ids=[item["source_id"]], source_items=[item], working_name=item["name"],
            decision_id="", classification_source_id=item["source_id"],
            status=source_disposition[item["source_id"]]["status"],
            details=source_disposition[item["source_id"]]["details"]))

    # A missing decision group must fall back to unchanged individual source entries.
    assigned_ids = ({source_id for group in output_groups for source_id in group["source_ids"]} |
                    handled_decision_ids)
    for item in source_items:
        if item["source_id"] not in assigned_ids:
            output_groups.append(dict(
                source_ids=[item["source_id"]], source_items=[item], working_name=item["name"],
                decision_id="", classification_source_id=item["source_id"],
                status=source_disposition[item["source_id"]]["status"],
                details=source_disposition[item["source_id"]]["details"] or
                "Decision source set changed; review required."))
            source_disposition[item["source_id"]].update(
                status="WORKING_LIST_REVIEW_REQUIRED",
                details="Decision source set changed; source record kept unchanged.")

    # Preserve source order among merged entries using the first selected source position.
    output_groups.sort(key=lambda group: min(item["source_position"] for item in group["source_items"]))
    _validate_source_accounting(source_items, output_groups, source_disposition)
    working_root = ET.Element("PickList", {
        "SourceSHA256": source_digest,
        "DecisionRegistrySHA256": _sha(registry_path.read_bytes()),
        "RecordBaselineSHA256": record_baseline_digest,
        "WorkingListSchema": "1",
    })
    working_items = ET.SubElement(working_root, "Items")
    working_rows = []
    for group in output_groups:
        representative = group["source_items"][0]
        source_ids = group["source_ids"]
        work_id = _working_id(source_ids, representative["plu"], representative["department"],
                              group["working_name"], group["decision_id"])
        node = copy.deepcopy(source_nodes[representative["source_position"] - 1])
        _set_description(node, group["working_name"])
        node.set("WorkingId", work_id)
        node.set("SourceIdentityIds", ";".join(source_ids))
        node.set("DecisionId", group["decision_id"])
        node.set("ClassificationSourceId", group["classification_source_id"])
        node.set("WorkingStatus", group["status"])
        node.set("WorkingDetails", group["details"])
        working_items.append(node)
        row = dict(
            representative, source_id=work_id, name=group["working_name"],
            source_ids=source_ids,
            source_positions=[item["source_position"] for item in group["source_items"]],
            decision_id=group["decision_id"],
            classification_source_id=group["classification_source_id"],
            identity_changed=(work_id != representative["source_id"]),
            source_record_status=("MATCHED_RECORD_BASELINE" if all(
                source_record_status[source_id] == "MATCHED_RECORD_BASELINE"
                for source_id in source_ids) else "SOURCE_RECORD_REVIEW_REQUIRED"),
            working_list_status=group["status"], working_list_details=group["details"],
        )
        working_rows.append(row)
        for source_id in source_ids:
            source_disposition[source_id]["working_id"] = work_id
            source_disposition[source_id]["working_name"] = group["working_name"]

    working_tree = ET.ElementTree(working_root)
    ET.indent(working_tree, space="  ")
    working_path.parent.mkdir(parents=True, exist_ok=True)
    working_tree.write(working_path, encoding="utf-8", xml_declaration=True)

    audit_output.mkdir(parents=True, exist_ok=True)
    audit_root = ET.Element("PickListSourceAudit", {
        "SourceSHA256": source_digest,
        "DecisionRegistrySHA256": _sha(registry_path.read_bytes()),
        "RecordBaselineSHA256": record_baseline_digest,
    })
    audit_records = ET.SubElement(audit_root, "Records")
    reconciliation_rows = []
    working_id_by_source = {
        source_id: row["source_id"] for row in working_rows for source_id in row["source_ids"]
    }
    for item, original_node in zip(source_items, source_nodes):
        disposition = source_disposition[item["source_id"]]
        wrapper = ET.SubElement(audit_records, "SourceRecord", {
            "SourceIdentityId": item["source_id"],
            "SourcePosition": str(item["source_position"]),
            "LogicalIdentity": json.dumps(item["logical_identity"], ensure_ascii=False,
                                          separators=(",", ":")),
            "RecordSHA256": item["record_sha256"],
            "RecordBaselineSHA256": disposition["record_baseline_sha256"],
            "RecordFingerprintStatus": disposition["record_status"],
            "WorkingId": working_id_by_source.get(item["source_id"], ""),
            "Disposition": disposition["disposition"],
            "DecisionId": disposition["decision_id"],
            "ReviewStatus": disposition["status"],
        })
        wrapper.append(copy.deepcopy(original_node))
        reconciliation_rows.append(dict(
            source_id=item["source_id"], source_position=item["source_position"],
            logical_identity=json.dumps(item["logical_identity"], ensure_ascii=False,
                                        separators=(",", ":")),
            record_sha256=item["record_sha256"],
            record_baseline_sha256=disposition["record_baseline_sha256"],
            record_status=disposition["record_status"],
            plu=item["plu"], original_name=item["name"], department=item["department"],
            visible=item["visible"], deleted=item["deleted"], included=item["included"],
            working_id=working_id_by_source.get(item["source_id"], ""),
            working_name=disposition["working_name"],
            disposition=disposition["disposition"], decision_id=disposition["decision_id"],
            review_status=disposition["status"], details=disposition["details"]))
    orphaned_baseline = [
        row for source_id, row in record_baseline.items()
        if source_id not in {item["source_id"] for item in source_items}
    ]
    orphan_xml = ET.SubElement(audit_root, "OrphanedBaselineIdentities")
    for row in orphaned_baseline:
        identity = row["logical_identity"]
        ET.SubElement(orphan_xml, "SourceIdentity", {
            "SourceIdentityId": row["source_id"],
            "LogicalIdentity": json.dumps(identity, ensure_ascii=False, separators=(",", ":")),
            "RecordSHA256": row["record_sha256"],
            "BaselineSourcePosition": str(row["source_position"]),
            "Disposition": "SOURCE_IDENTITY_REMOVED_REVIEW",
        })
    audit_tree = ET.ElementTree(audit_root)
    ET.indent(audit_tree, space="  ")
    audit_tree.write(audit_output / "PICKLIST_SOURCE_RECORDS.xml",
                     encoding="utf-8", xml_declaration=True)
    fields = ["source_id", "source_position", "logical_identity", "record_sha256",
              "record_baseline_sha256", "record_status", "plu", "original_name", "department",
              "visible", "deleted", "included", "working_id", "working_name",
              "disposition", "decision_id", "review_status", "details"]
    _write_csv(audit_output / "WORKING_LIST_RECONCILIATION.csv", reconciliation_rows, fields)
    record_change_rows = []
    for item in source_items:
        disposition = source_disposition[item["source_id"]]
        if disposition["status"] in {
                "UNCHANGED_SOURCE_IDENTITY", "APPLIED_TO_EXACT_SOURCE_IDENTITIES"}:
            continue
        saved = record_baseline.get(item["source_id"], {})
        record_change_rows.append(dict(
            source_id=item["source_id"], source_position=item["source_position"],
            baseline_source_position=saved.get("source_position", ""),
            logical_identity=json.dumps(item["logical_identity"], ensure_ascii=False,
                                        separators=(",", ":")),
            plu=item["plu"], name=item["name"], department=item["department"],
            record_sha256=item["record_sha256"],
            baseline_record_sha256=saved.get("record_sha256", ""),
            status=disposition["status"], disposition=disposition["disposition"],
            decision_id=disposition["decision_id"], details=disposition["details"],
        ))
    for saved in orphaned_baseline:
        identity = saved["logical_identity"]
        record_change_rows.append(dict(
            source_id=saved["source_id"], source_position="",
            baseline_source_position=saved["source_position"],
            logical_identity=json.dumps(identity, ensure_ascii=False, separators=(",", ":")),
            plu=identity[0], name=identity[1], department=identity[2],
            record_sha256="", baseline_record_sha256=saved["record_sha256"],
            status="SOURCE_IDENTITY_REMOVED_REVIEW",
            disposition="SOURCE_IDENTITY_REMOVED_FROM_CURRENT_PICKLIST",
            decision_id="", details="The baseline identity is absent from the current PickList.",
        ))
    _write_csv(audit_output / "PICKLIST_RECORD_CHANGES.csv", record_change_rows, [
        "source_id", "source_position", "baseline_source_position", "logical_identity",
        "plu", "name", "department", "record_sha256", "baseline_record_sha256",
        "status", "disposition", "decision_id", "details",
    ])
    _write_csv(audit_output / "WORKING_LIST_ENTRIES.csv", [
        dict(source_id=row["source_id"], source_ids=";".join(row["source_ids"]),
             source_positions=";".join(map(str, row["source_positions"])),
             record_sha256=row["record_sha256"],
             source_record_status=row["source_record_status"],
             plu=row["plu"], name=row["name"], department=row["department"],
             decision_id=row["decision_id"], working_list_status=row["working_list_status"],
             image_review_required=row["identity_changed"])
        for row in working_rows
    ], ["source_id", "source_ids", "source_positions", "record_sha256", "source_record_status",
        "plu", "name", "department",
        "decision_id", "working_list_status", "image_review_required"])
    summary = dict(
        source="sources/plu/PickList.xml", working_list="sources/plu/WorkingList.xml",
        source_sha256=source_digest, decision_registry_sha256=_sha(registry_path.read_bytes()),
        record_baseline_sha256=record_baseline_digest,
        source_records_with_changed_fingerprints=sum(
            value != "MATCHED_RECORD_BASELINE" for value in source_record_status.values()),
        orphaned_record_baseline_identities=len(set(record_baseline) -
                                                {item["source_id"] for item in source_items}),
        source_records=len(source_items), working_records=len(working_rows),
        merged_or_standardized_records=sum(row["identity_changed"] for row in working_rows),
        excluded_conflicting_source_records=sum(
            row["disposition"] == "EXCLUDED_CONFLICTING_ASSOCIATION" for row in reconciliation_rows),
        excluded_unconfirmed_equivalence_records=sum(
            row["disposition"] == "EXCLUDED_EQUIVALENCE_NOT_CONFIRMED"
            for row in reconciliation_rows),
        hidden_or_deleted_source_records=sum(not item["included"] for item in source_items),
        review_required_records=sum(row["review_status"] not in {
            "UNCHANGED_SOURCE_IDENTITY", "APPLIED_TO_EXACT_SOURCE_IDENTITIES"
        } for row in reconciliation_rows),
        decisions=applied_decisions,
        source_unchanged=True,
    )
    (audit_output / "WORKING_LIST_RECONCILIATION.json").write_text(
        json.dumps(summary, indent=2) + "\n", encoding="utf-8")
    return working_rows, source_items, summary


def read_working_list(root):
    root = Path(root)
    source_items, source_digest = read_picklist(root / "sources/plu/PickList.xml")
    working_path = root / "sources/plu/WorkingList.xml"
    try:
        tree = ET.parse(working_path)
    except (OSError, ET.ParseError) as exc:
        raise ValueError(f"cannot parse generated WorkingList.xml: {exc}") from exc
    source_hash = tree.getroot().get("SourceSHA256")
    if source_hash != source_digest:
        raise ValueError("WorkingList.xml is stale; rebuild it from the current PickList.xml")
    registry_path = root / "data/working_list_decisions.json"
    expected_registry_hash = _sha(registry_path.read_bytes())
    if tree.getroot().get("DecisionRegistrySHA256") != expected_registry_hash:
        raise ValueError("WorkingList.xml is stale; rebuild it from the current decision registry")
    record_baseline, record_baseline_digest, _ = read_record_baseline(
        root / "data/picklist_record_baseline.json")
    if tree.getroot().get("RecordBaselineSHA256") != record_baseline_digest:
        raise ValueError("WorkingList.xml is stale; rebuild it from the current record baseline")
    parsed, _ = read_picklist(working_path)
    nodes = tree.getroot().findall("./Items/Item")
    if len(parsed) != len(nodes):
        raise ValueError("WorkingList record positions differ from parsed records")
    rows = []
    original_by_id = {item["source_id"]: item for item in source_items}
    seen_working_ids = set()
    seen_source_ids = set()
    for row, node in zip(parsed, nodes):
        identities = [value for value in node.get("SourceIdentityIds", "").split(";") if value]
        working_id = node.get("WorkingId")
        if (not identities or len(identities) != len(set(identities)) or not working_id or
                working_id in seen_working_ids or seen_source_ids.intersection(identities)):
            raise ValueError(f"WorkingList PLU {row['plu']!r}: invalid source identity linkage")
        seen_working_ids.add(working_id)
        seen_source_ids.update(identities)
        row["source_id"] = working_id
        row["source_ids"] = identities
        row["decision_id"] = node.get("DecisionId", "")
        row["classification_source_id"] = node.get("ClassificationSourceId", "")
        row["identity_changed"] = row["source_id"] not in identities
        row["working_list_status"] = node.get("WorkingStatus", "")
        row["working_list_details"] = node.get("WorkingDetails", "")
        originals = [original_by_id.get(identity) for identity in identities]
        if any(original is None for original in originals):
            raise ValueError(f"WorkingList PLU {row['plu']!r}: source identity is not in PickList.xml")
        if any(original["plu"] != row["plu"] for original in originals):
            raise ValueError(f"WorkingList PLU {row['plu']!r}: linked source PLU differs")
        if row["classification_source_id"] and row["classification_source_id"] not in identities:
            raise ValueError(f"WorkingList PLU {row['plu']!r}: classification source is not linked")
        record_states = []
        for original in originals:
            saved = record_baseline.get(original["source_id"])
            record_states.append(
                "NEW_SOURCE_RECORD_REVIEW" if saved is None else
                "SOURCE_RECORD_CHANGED_REVIEW"
                if saved["record_sha256"] != original["record_sha256"] else
                "MATCHED_RECORD_BASELINE")
        row["source_record_status"] = (
            "MATCHED_RECORD_BASELINE" if all(
                value == "MATCHED_RECORD_BASELINE" for value in record_states) else
            next(value for value in record_states if value != "MATCHED_RECORD_BASELINE"))
        row["source_record_changed"] = row["source_record_status"] != "MATCHED_RECORD_BASELINE"
        row["source_positions"] = [original["source_position"] for original in originals]
        rows.append(row)
    return rows, source_digest
