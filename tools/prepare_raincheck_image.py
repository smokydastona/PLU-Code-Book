"""Apply audited field-format and product-identity corrections to the owner upload."""
import argparse
import hashlib
import io
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
UPLOAD = Path("sources/binders/Rain_Check_Instructions_Upload.png")
OUTPUT = Path("sources/binders/Rain_Check_Instructions_Updated.png")
UPLOAD_SHA256 = "8377569b18af0d783745f4a8fc0336a66b6baac2ca0fb20a8598f25e725baa44"
OUTPUT_SHA256 = "c211ff2d648e93fd4663ae6f5183023b173c46e9c892416c42125614e6745f3a"
SIZE = (1024, 1536)
PAPER = (250, 250, 250)
CORNER_INSTRUCTION = "Follow these steps to complete the raincheck/voucher correctly."


def digest(content):
    return hashlib.sha256(content).hexdigest()


def font(name, size):
    path = Path("/usr/share/fonts/truetype/dejavu") / name
    return ImageFont.truetype(str(path), size)


def paint_gradient_background(image, box, left_x, right_x):
    pixels = image.load()
    left, top, right, bottom = box
    for y in range(top, bottom):
        start, end = pixels[left_x, y], pixels[right_x, y]
        for x in range(left, right):
            ratio = (x - left) / max(right - left - 1, 1)
            pixels[x, y] = tuple(round(a + (b - a) * ratio) for a, b in zip(start, end))


def replace_field_text(draw, image, box, text, font_obj, color, position):
    draw.rectangle(box, fill=PAPER)
    bounds = draw.textbbox((0, 0), text, font=font_obj)
    x = position[0] - bounds[0]
    y = position[1] - bounds[1]
    draw.text((x, y), text, font=font_obj, fill=color)


def prepare_image(content):
    if digest(content) != UPLOAD_SHA256:
        raise ValueError("Rain Check owner upload changed; compare and review it before editing")
    with Image.open(io.BytesIO(content)) as opened:
        if opened.size != SIZE:
            raise ValueError(f"Rain Check owner image must be {SIZE[0]} x {SIZE[1]} pixels")
        image = opened.convert("RGB")

    body = font("DejaVuSans.ttf", 15)
    body_bullet = font("DejaVuSans.ttf", 16)
    ink = (12, 20, 45)
    draw = ImageDraw.Draw(image)

    paint_gradient_background(image, (624, 207, 997, 250), 623, 998)
    for y, text in (
        (208, "Write your name and department."),
        (228, "Fill in store number, banner, and store name."),
    ):
        draw.text((626, y), "•", font=body_bullet, fill=ink)
        draw.text((646, y), text, font=body, fill=ink)

    paint_gradient_background(image, (624, 300, 997, 330), 623, 998)
    draw.text((626, 303), "•", font=body_bullet, fill=ink)
    draw.text((646, 302), "Write today's date in mm/dd/yy format.", font=body, fill=ink)

    corner_box = (731, 11, 994, 145)
    paint_gradient_background(image, corner_box, 728, 996)
    corner_font = font("DejaVuSans-Bold.ttf", 22)
    corner_text = "Follow these steps\nto complete the\nraincheck/voucher\ncorrectly."
    bounds = draw.multiline_textbbox((0, 0), corner_text, font=corner_font, align="center", spacing=4)
    left, top, right, bottom = bounds
    x = (corner_box[0] + corner_box[2] - right + left) / 2 - left
    y = (corner_box[1] + corner_box[3] - bottom + top) / 2 - top
    draw.multiline_text((x, y), corner_text, font=corner_font, fill=ink, align="center", spacing=4)

    blue = (0, 59, 190)
    date_font = font("DejaVuSans-Oblique.ttf", 26)
    plu_font = font("DejaVuSans-Oblique.ttf", 21)
    replace_field_text(draw, image, (214, 1099, 334, 1128),
                       "11/17/26", date_font, blue, (218, 1106))
    replace_field_text(draw, image, (158, 1235, 280, 1255),
                       "3181", plu_font, blue, (184, 1238))
    replace_field_text(draw, image, (168, 1189, 440, 1212),
                       "Strawberries BC 1 Pound", plu_font, blue, (175, 1192))

    result = io.BytesIO()
    image.save(result, format="PNG", optimize=True)
    return result.getvalue()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args()
    source = args.root / UPLOAD
    output = args.root / OUTPUT
    content = source.read_bytes()
    updated = prepare_image(content)
    if OUTPUT_SHA256 and digest(updated) != OUTPUT_SHA256:
        parser.exit(1, "ERROR: prepared Rain Check image differs from the reviewed checksum\n")
    output.write_bytes(updated)
    print(f"{OUTPUT}: {digest(updated)}")


if __name__ == "__main__":
    main()
