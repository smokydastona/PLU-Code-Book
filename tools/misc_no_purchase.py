"""Crop and verify the No Purchase Entry panel from the owner-supplied scan."""
import hashlib
import io
from pathlib import Path

from PIL import Image
import zxingcpp

SOURCE = Path("sources/binders/Misc_Codes_No_Purchase_Upload.jpg")
SOURCE_SHA256 = "769aab1c558d98513b3f5469f2489095f74ac118d584167353676ec4db858a90"
CROP = (194, 1161, 1104, 1450)
CODE = "0770974140186"


def verify_barcode(image):
    results = zxingcpp.read_barcodes(image)
    if len(results) != 1 or results[0].text != CODE:
        raise ValueError(f"No Purchase Entry barcode must decode as {CODE}")


def crop_panel(content):
    if hashlib.sha256(content).hexdigest() != SOURCE_SHA256:
        raise ValueError("No Purchase Entry scan changed; review it before rebuilding")
    with Image.open(io.BytesIO(content)) as scan:
        panel = scan.convert("RGB").crop(CROP)
    verify_barcode(panel)
    output = io.BytesIO()
    panel.save(output, format="PNG")
    return output.getvalue()
