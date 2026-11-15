"""Read the owner-maintained PickList without modifying or correcting its values."""
import hashlib
import json
import re
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path


def _semantic_element(element):
    children = list(element)
    text = element.text or ""
    if children and not text.strip():
        text = ""
    return (
        element.tag,
        tuple(sorted(element.attrib.items())),
        text,
        tuple(
            (_semantic_element(child), child.tail if child.tail and child.tail.strip() else "")
            for child in children
        ),
    )


def record_fingerprint(item):
    semantic = json.dumps(
        _semantic_element(item), ensure_ascii=False, separators=(",", ":")
    ).encode("utf-8")
    return hashlib.sha256(semantic).hexdigest()


def source_identity_id(logical_identity):
    key = json.dumps(logical_identity, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(key.encode("utf-8")).hexdigest()


def record_baseline_payload(rows, source_sha256):
    return {
        "schema_version": 1,
        "source_sha256": source_sha256,
        "records": [
            {
                "source_id": row["source_id"],
                "logical_identity": row["logical_identity"],
                "record_sha256": row["record_sha256"],
                "source_position": row["source_position"],
            }
            for row in rows
        ],
    }


def write_record_baseline(path, rows, source_sha256):
    path = Path(path)
    if path.exists():
        raise ValueError(f"record baseline already exists; refusing to overwrite: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(record_baseline_payload(rows, source_sha256), indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


def read_record_baseline(path):
    path = Path(path)
    try:
        content = path.read_bytes()
    except OSError as exc:
        raise ValueError(f"cannot read PickList record baseline: {exc}") from exc
    try:
        data = json.loads(content)
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"cannot parse PickList record baseline: {exc}") from exc
    if not isinstance(data, dict):
        raise ValueError("record baseline must be a JSON object")
    if data.get("schema_version") != 1 or not isinstance(data.get("records"), list):
        raise ValueError("expected record baseline schema_version 1 and a records list")
    if (not isinstance(data.get("source_sha256"), str) or
            re.fullmatch(r"[0-9a-f]{64}", data["source_sha256"]) is None):
        raise ValueError("record baseline has an invalid source SHA-256")
    records, positions = {}, set()
    for row in data["records"]:
        if not isinstance(row, dict):
            raise ValueError("record baseline entries must be objects")
        identity = row.get("logical_identity")
        if (not isinstance(identity, list) or len(identity) != 4 or
                not isinstance(identity[0], str) or not isinstance(identity[1], str) or
                not isinstance(identity[2], str) or
                type(identity[3]) is not int or identity[3] < 1):
            raise ValueError("record baseline contains an invalid logical identity")
        source_id = row.get("source_id")
        if source_id != source_identity_id(identity):
            raise ValueError(f"record baseline identity checksum mismatch: {source_id!r}")
        record_sha256 = row.get("record_sha256")
        if (not isinstance(record_sha256, str) or
                re.fullmatch(r"[0-9a-f]{64}", record_sha256) is None):
            raise ValueError(f"record baseline has an invalid record SHA-256: {source_id}")
        position = row.get("source_position")
        if type(position) is not int or position < 1:
            raise ValueError(f"record baseline has an invalid source position: {source_id}")
        if position in positions:
            raise ValueError(f"duplicate source position in record baseline: {position}")
        if source_id in records:
            raise ValueError(f"duplicate source identity in record baseline: {source_id}")
        records[source_id] = row
        positions.add(position)
    if positions != set(range(1, len(records) + 1)):
        raise ValueError("record baseline source positions are incomplete")
    return records, hashlib.sha256(content).hexdigest(), data["source_sha256"]


def flag(item, name, default=None):
    value = item.findtext(name)
    if value is None and default is not None:
        return default
    if value not in {"True", "False"}:
        raise ValueError(f"PLU {item.get('UPC')!r}: {name} must be True or False")
    return value == "True"


def read_picklist(path):
    content = Path(path).read_bytes()
    try:
        root = ET.fromstring(content)
    except ET.ParseError as exc:
        raise ValueError(f"cannot parse owner-supplied PickList.xml: {exc}") from exc
    item_containers = root.findall("Items")
    if root.tag != "PickList" or len(item_containers) != 1:
        raise ValueError("expected PickList/Items in source XML")
    rows, occurrences = [], Counter()
    items = item_containers[0].findall("Item")
    if any(child.tag != "Item" for child in list(item_containers[0])):
        raise ValueError("PickList/Items may contain only Item records")
    for position, item in enumerate(items, 1):
        plu = item.get("UPC")
        department_nodes = item.findall("Department")
        visible_nodes = item.findall("IsVisible")
        deleted_nodes = item.findall("DeleteThisItem")
        language_nodes = item.findall("./Languages/Language[@LCID='0409']/Description")
        if (len(department_nodes) != 1 or len(visible_nodes) != 1 or
                len(deleted_nodes) != 1 or len(language_nodes) != 1):
            raise ValueError(f"XML item {position}: missing or ambiguous required source fields")
        name = language_nodes[0].text
        department = department_nodes[0].text or ""
        visible = flag(item, "IsVisible")
        deleted = flag(item, "DeleteThisItem")
        if not isinstance(plu, str) or not plu.strip():
            raise ValueError(f"XML item {position}: missing UPC")
        if not isinstance(name, str) or not name.strip():
            raise ValueError(f"XML item {position}: missing English (0409) description")
        identity = (plu, name, department)
        occurrences[identity] += 1
        logical_identity = [*identity, occurrences[identity]]
        rows.append(dict(
            source_id=source_identity_id(logical_identity),
            logical_identity=logical_identity,
            record_sha256=record_fingerprint(item),
            source_position=position, occurrence=occurrences[identity],
            plu=plu, name=name, department=department, visible=visible, deleted=deleted,
            included=visible and not deleted,
            picture_references=[picture.text or "" for picture in item.findall("./Pictures/Picture")],
        ))
    if not rows:
        raise ValueError("PickList contains no items")
    return rows, hashlib.sha256(content).hexdigest()


def read_assignments(path):
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"cannot read image assignment registry: {exc}") from exc
    if (not isinstance(data, dict) or data.get("schema_version") != 1 or
            not isinstance(data.get("assignments"), list)):
        raise ValueError("expected schema_version 1 and an assignments list")
    result = {}
    for index, row in enumerate(data["assignments"], 1):
        if not isinstance(row, dict):
            raise ValueError(f"assignment {index} must be an object")
        key = row.get("source_id")
        if not isinstance(key, str) or re.fullmatch(r"[0-9a-f]{64}", key) is None:
            raise ValueError(f"assignment {index} has a missing or invalid source_id")
        if key in result:
            raise ValueError(f"duplicate assignment source_id: {key}")
        if not isinstance(row.get("book"), str) or row["book"] not in {
                "Fruit", "Vegetable", "Non-produce"}:
            raise ValueError(f"assignment {key}: unsupported or missing book")
        if not isinstance(row.get("image_mode"), str) or row["image_mode"] not in {
                "PHOTO", "CODE_ONLY"}:
            raise ValueError(f"assignment {key}: unsupported or missing image_mode")
        for field in ("image", "reviewed_plu", "reviewed_product"):
            if not isinstance(row.get(field), str):
                raise ValueError(f"assignment {key}: {field} must be a string")
        if row.get("provenance") is not None and not isinstance(row["provenance"], dict):
            raise ValueError(f"assignment {key}: provenance must be an object or null")
        result[key] = row
    return result
