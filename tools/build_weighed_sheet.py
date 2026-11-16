"""Reviewed owner-uploaded weighed-items guide, fitted to one binder page."""
import hashlib
import io
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.worksheet.page import PageMargins
from PIL import Image, ImageDraw, ImageFont

SHEET = "Weighed more $100"
SOURCE = Path("sources/binders/Weighed_Items_Over_100_Updated.jpg")
SOURCE_SHA256 = "bd2689fd49f13d5050753e30d850e3c89001e9db98bcf3ff9930bddd374d3604"
ORIGINAL_CELLS_SHA256 = "e8e096a7650e27f8e67382de3f894e4083fa269230cca9f117fe9a2621be0b21"
ORIGINAL_LABEL_SHA256 = "506394d7b873668c64e1143a56b9455cd7e774baa37387dff27cae5a13e644b6"
PRICE_CORRECTION_BOX = (939, 537, 966, 588)
ACCURACY_REVIEW = (
    "Uploaded guide compared with every original instruction and the original embedded label. "
    "Preserved UPC 052110 / entry 52110, weight 3.830 kg / entry 3830, ending-in-0-or-5 rule, "
    "and examples 3.826 -> 3830, 4.117 -> 4120, 2.941 -> 2945 exactly as supplied. "
    "Label product, date 2024JL18, time 15:01, unit price $35.07/kg and total $134.31 match. "
    "Owner requested correction of the displayed total to $134.32: $35.07/kg x 3.830 kg = "
    "$134.3181, rounded to cents. Only the final price digit is replaced in the label. "
    "Green colors are muted to the binder palette; other label pixels, red warnings, text geometry "
    "and image dimensions are unchanged. Original upload and original workbook are retained."
)


def digest(content):
    return hashlib.sha256(content).hexdigest()


def themed_image(content):
    if digest(content) != SOURCE_SHA256:
        raise ValueError("Weighed guide upload changed; review its accuracy before rebuilding")
    with Image.open(io.BytesIO(content)) as original:
        image = original.convert("RGB")
    pixels = image.load()
    for y in range(image.height):
        for x in range(image.width):
            if 578 <= x < 990 and 205 <= y < 636:
                continue
            red, green, blue = pixels[x, y]
            if green > red + 8 and green > blue + 4:
                # Match dark green while leaving white, black and red content untouched.
                factor = min(green / 80, 1)
                light = max((green - 80) / 175, 0)
                pixels[x, y] = tuple(round(dark * factor + (255 - dark) * light)
                                     for dark in (48, 80, 64))
    # Replace just the final digit, retaining the original price typography and box.
    digit = Image.new("L", (100, 100))
    ImageDraw.Draw(digit).text((0, 0), "2", font=ImageFont.truetype("DejaVuSans-Bold.ttf", 64),
                              fill=255)
    digit = digit.crop(digit.getbbox()).resize((18, 41), Image.Resampling.LANCZOS)
    image.paste(image.getpixel((950, 535)), PRICE_CORRECTION_BOX)
    image.paste((15, 15, 15), (943, 541), digit)
    result = io.BytesIO()
    image.save(result, format="PNG", optimize=True)
    return result.getvalue()


def workbook_bytes(source, image):
    cells = [(cell.coordinate, cell.value) for row in source for cell in row
             if cell.value is not None]
    if source._charts or any(cell.comment or cell.hyperlink for row in source for cell in row):
        raise ValueError("Weighed guide contains additional objects; review before replacement")
    hashes = [digest(picture._data()) for picture in source._images]
    if source.sheet_properties.codeName == "UpdatedWeighedImage":
        if cells or hashes != [digest(image)]:
            raise ValueError("Updated Weighed guide changed; review before replacement")
    elif (digest(json.dumps(cells, ensure_ascii=True).encode()) != ORIGINAL_CELLS_SHA256 or
          hashes != [ORIGINAL_LABEL_SHA256]):
        raise ValueError("Original Weighed instructions or label changed; review their accuracy")
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = SHEET
    sheet.sheet_properties.codeName = "UpdatedWeighedImage"
    sheet.sheet_view.showGridLines = False
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "portrait"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_LETTER
    sheet.page_setup.fitToWidth = sheet.page_setup.fitToHeight = 1
    sheet.page_margins = PageMargins(left=.25, right=.25, top=.25, bottom=.25, header=.1, footer=.1)
    sheet.print_options.horizontalCentered = True
    sheet.column_dimensions["A"].width = 100
    sheet.sheet_format.defaultRowHeight = 15
    sheet.print_area = "A1:A53"
    picture = ExcelImage(io.BytesIO(image))
    picture.width, picture.height = 700, 1050
    picture.anchor = OneCellAnchor(
        _from=AnchorMarker(col=0, row=0, colOff=pixels_to_EMU(100)),
        ext=XDRPositiveSize2D(cx=pixels_to_EMU(700), cy=pixels_to_EMU(1050)))
    sheet.add_image(picture)
    result = io.BytesIO()
    workbook.save(result)
    workbook.close()
    return result.getvalue()


def validate_sheet(sheet, image):
    if (sheet.sheet_properties.codeName != "UpdatedWeighedImage" or len(sheet._images) != 1 or
            digest(sheet._images[0]._data()) != digest(image) or sheet._charts or
            any(cell.value is not None for row in sheet for cell in row) or
            sheet.page_setup.orientation != "portrait" or str(sheet.page_setup.paperSize) != "1" or
            sheet.page_setup.fitToWidth != 1 or sheet.page_setup.fitToHeight != 1 or
            not sheet.sheet_properties.pageSetUpPr.fitToPage or str(sheet.print_area) !=
            "'Weighed more $100'!$A$1:$A$53"):
        raise ValueError("Weighed guide image or one-page print settings changed")
