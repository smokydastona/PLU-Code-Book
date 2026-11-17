"""Owner-uploaded weighed-items guide and rain-check page for the binder."""
import hashlib
import io
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.page import PageMargins
from PIL import Image, ImageDraw, ImageFont
import pymupdf

SHEET = "Weighed more $100"
SOURCE = Path("sources/binders/Weighed_Items_Over_100_Updated.jpg")
SOURCE_SHA256 = "bd2689fd49f13d5050753e30d850e3c89001e9db98bcf3ff9930bddd374d3604"
RAIN_CHECK_SOURCE = Path("sources/binders/Rain_Check_Instructions_Updated.png")
RAIN_CHECK_SHA256 = "c211ff2d648e93fd4663ae6f5183023b173c46e9c892416c42125614e6745f3a"
RAIN_CHECK_UPLOAD = Path("sources/binders/Rain_Check_Instructions_Upload.png")
RAIN_CHECK_UPLOAD_SHA256 = "8377569b18af0d783745f4a8fc0336a66b6baac2ca0fb20a8598f25e725baa44"
RAIN_CHECK_ORIGINAL_PDF = Path("sources/binders/Rain_Check.pdf")
RAIN_CHECK_ORIGINAL_PDF_SHA256 = "9aeb587f43b8022735666e04dd771c89d7d08f15150964f7f095e6dc762af777"
RAIN_CHECK_ACCURACY_REVIEW = (
    "Compared the owner upload with the original Rain Check PDF. Updated the sample date to the form's "
    "mm/dd/yy format, matched the sample product description and PLU 3181 to Strawberries BC 1 Pound, "
    "and added the store number, banner, and store name fields to step 1. The original form, 30-day "
    "validity, quantity-in-words, size, price-per-unit, and SCENE+ instructions remain present. "
    "Removed the confusing accuracy claim from the corner and enlarged the remaining instruction to "
    "'Follow these steps to complete the raincheck/voucher correctly.' The 1 lb sample is 454 g; "
    "$4.99/lb converts to $11.00/kg rounded to cents."
)
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


def rain_check_image(content):
    if digest(content) != RAIN_CHECK_SHA256:
        raise ValueError("Corrected Rain Check image changed; review it before rebuilding")
    with Image.open(io.BytesIO(content)) as image:
        if image.format != "PNG" or image.size != (1024, 1536):
            raise ValueError("Corrected Rain Check image must be a 1024 x 1536 PNG")
        image.verify()
    return content


def workbook_bytes(source, image, rain_check):
    cells = [(cell.coordinate, cell.value) for row in source for cell in row
             if cell.value is not None]
    if source._charts or any(cell.comment or cell.hyperlink for row in source for cell in row):
        raise ValueError("Weighed guide contains additional objects; review before replacement")
    hashes = [digest(picture._data()) for picture in source._images]
    if source.sheet_properties.codeName == "UpdatedWeighedImage":
        if cells or hashes not in ([digest(image)], [digest(image), digest(rain_check)]):
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
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.page_margins = PageMargins(left=.25, right=.25, top=.25, bottom=.25, header=.1, footer=.1)
    sheet.print_options.horizontalCentered = True
    sheet.column_dimensions["A"].width = 100
    sheet.sheet_format.defaultRowHeight = 14
    sheet.print_area = "A1:A106"
    sheet.row_breaks.append(Break(id=53))
    for content, row, width, height in ((image, 0, 650, 975), (rain_check, 55, 560, 840)):
        picture = ExcelImage(io.BytesIO(content))
        picture.anchor = OneCellAnchor(
            _from=AnchorMarker(col=0, row=row, colOff=pixels_to_EMU(100)),
            ext=XDRPositiveSize2D(cx=pixels_to_EMU(width), cy=pixels_to_EMU(height)))
        sheet.add_image(picture)
    result = io.BytesIO()
    workbook.save(result)
    workbook.close()
    return result.getvalue()


def validate_sheet(sheet, image, rain_check):
    if (sheet.sheet_properties.codeName != "UpdatedWeighedImage" or len(sheet._images) != 2 or
            [digest(picture._data()) for picture in sheet._images] !=
            [digest(image), digest(rain_check)] or
            [picture.anchor._from.row for picture in sheet._images] != [0, 55] or
            (sheet._images[1].anchor.ext.cx, sheet._images[1].anchor.ext.cy) !=
            (pixels_to_EMU(560), pixels_to_EMU(840)) or
            [page_break.id for page_break in sheet.row_breaks.brk] != [53] or sheet._charts or
            any(cell.value is not None for row in sheet for cell in row) or
            sheet.page_setup.orientation != "portrait" or str(sheet.page_setup.paperSize) != "1" or
            sheet.page_setup.fitToWidth != 1 or sheet.page_setup.fitToHeight != 0 or
            not sheet.sheet_properties.pageSetUpPr.fitToPage or str(sheet.print_area) !=
            "'Weighed more $100'!$A$1:$A$106"):
        raise ValueError("Weighed guide or Rain Check image or two-page print settings changed")
