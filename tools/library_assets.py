"""Resolve deduplicated JPEG storage without merging identity-specific metadata."""
import hashlib
import json
import re
from pathlib import Path

ALIASES = "IMAGE_ALIASES.json"
IMAGE_NAME = re.compile(r"[0-9a-f]{64}\.jpg")


def image_aliases(library):
    index = Path(library) / ALIASES
    if index.is_symlink():
        raise ValueError("image alias index must not be a symlink")
    if not index.exists():
        return {}
    aliases = json.loads(index.read_text(encoding="utf-8"))
    if not isinstance(aliases, dict):
        raise ValueError("image alias index must contain an object")
    for name, row in aliases.items():
        if (not IMAGE_NAME.fullmatch(name) or not isinstance(row, dict)
                or not IMAGE_NAME.fullmatch(row.get("image", ""))
                or name == row["image"] or row["image"] in aliases
                or not re.fullmatch(r"[0-9a-f]{64}", row.get("sha256", ""))):
            raise ValueError(f"invalid image alias: {name}")
    return aliases


def resolve_image(path, aliases=None):
    path = Path(path)
    if path.is_symlink():
        raise ValueError(f"library image must not be a symlink: {path}")
    aliases = image_aliases(path.parent) if aliases is None else aliases
    if path.exists() or path.name not in aliases:
        return path
    row = aliases[path.name]
    target = path.parent / row["image"]
    if (target.is_symlink() or not target.is_file()
            or hashlib.sha256(target.read_bytes()).hexdigest() != row["sha256"]):
        raise ValueError(f"image alias asset changed or missing: {path}")
    return target


def check_aliases(library):
    library = Path(library)
    aliases = image_aliases(library)
    for name, row in aliases.items():
        logical = library / name
        sidecar = library / (name + ".json")
        if sidecar.is_symlink() or not sidecar.is_file():
            raise ValueError(f"image alias metadata missing: {name}")
        metadata = json.loads(sidecar.read_text(encoding="utf-8"))
        if metadata.get("sha256") != row["sha256"]:
            raise ValueError(f"image alias metadata checksum mismatch: {name}")
        resolve_image(logical, aliases)
    return aliases


def physical_names(names, aliases):
    return {aliases[name]["image"] if name in aliases else name for name in names}
