"""Generate a clean No Purchase Entry panel for the verified scanned code."""
import hashlib
import io
from pathlib import Path

import barcode
from PIL import Image, ImageDraw, ImageFont
import zxingcpp

from binder_theme import DARK_GREEN

SOURCE = Path("sources/binders/Misc_Codes_No_Purchase_Upload.jpg")
SOURCE_SHA256 = "769aab1c558d98513b3f5469f2489095f74ac118d584167353676ec4db858a90"
CROP = (194, 1161, 1104, 1450)
LEGACY_PANEL_SHA256 = "43f557f86fae18793fabff7fc677431e6c9088a1a632fc286e676b6c6e2c4a51"
CODE = "0770974140186"
DESCRIPTION = (
    "For customers who wish to participate without purchasing",
    "participating products, scan the barcode below.",
)


def verify_barcode(image):
    results = zxingcpp.read_barcodes(image)
    if len(results) != 1 or results[0].text != CODE:
        raise ValueError(f"No Purchase Entry barcode must decode as {CODE}")


def generate_panel(content):
    if hashlib.sha256(content).hexdigest() != SOURCE_SHA256:
        raise ValueError("No Purchase Entry scan changed; review it before rebuilding")
    with Image.open(io.BytesIO(content)) as scan:
        verify_barcode(scan.crop(CROP).convert("RGB"))

    panel = Image.new("RGB", (910, 289), "white")
    draw = ImageDraw.Draw(panel)
    draw.rectangle((0, 0, 909, 47), fill="#" + DARK_GREEN)
    draw.text((18, 24), "No Purchase Entry", font=ImageFont.truetype("DejaVuSans-Bold.ttf", 28),
              fill="white", anchor="lm")
    body = ImageFont.truetype("DejaVuSans.ttf", 26)
    for y, line in zip((58, 89), DESCRIPTION):
        draw.text((18, y), line, font=body, fill="black", anchor="lt")

    symbol = barcode.get("ean13", CODE[:-1])
    modules, = symbol.build()
    # Draw whole pixels per module and retain an unobstructed white quiet zone.
    left = 64
    for i, module in enumerate(modules):
        if module == "1":
            x = left + 4 * i
            draw.rectangle((x, 129, x + 3, 244), fill="black")
    draw.text((left + 2 * len(modules), 257), CODE, font=body, fill="black", anchor="mt")
    verify_barcode(panel)
    output = io.BytesIO()
    panel.save(output, format="PNG")
    return output.getvalue()
