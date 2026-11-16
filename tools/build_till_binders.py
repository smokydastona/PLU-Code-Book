"""Generate full till-binder review copies with updated Produce, Bakery and Deli sheets."""
import argparse
import copy
import csv
import hashlib
import io
import json
import posixpath
import re
import shlex
import shutil
import subprocess
import sys
import tempfile
import zipfile
from collections import Counter, defaultdict
from decimal import Decimal
from pathlib import Path

from docx import Document
from docx.image.exceptions import UnrecognizedImageError
from docx.image.image import Image as DocxImage
from docx.image.exceptions import UnrecognizedImageError
from docx.image.image import Image as DocxImage
from docx.oxml.ns import qn
from lxml import etree as ET
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.cell import range_boundaries, get_column_letter
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.page import PageMargins
from PIL import Image, ImageChops, ImageDraw, ImageFont, ImageOps, PngImagePlugin

from check_image_uniqueness import image_fingerprint
from organic_watermarks import render_watermark
from picklist_catalog import read_assignments, read_picklist
from produce_order import display_unique_rows, order_produce
from review_image_references import load_references
from build_bakery_sheet import (BAKERY_SOURCE, TITLE as BAKERY_TITLE,
                                read_source as read_bakery_source,
                                validate_sheet as validate_bakery_sheet)
from sync_picklist import (booklet_order, csv_output, organic, organic_display_section,
                           reconcile)
from working_list import read_working_list
from build_deli_sheet import DELI_SOURCE, read_source as read_deli_source, validate_sheet as validate_deli_sheet
from build_bulk_sheet import TITLE as BULK_TITLE, bulk_sheet, read_bulk, validate_bulk
from binder_theme import DARK_GREEN, PALE_GREEN, row_fill

ROOT = Path(__file__).resolve().parents[1]
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"s": S, "r": R, "p": P}
FILENAMES = {
    "photos": "Complete Till Binder - Updated - Photos.xlsx",
}
MAX_WORKBOOK_IMAGE_SIZE = (192, 120)
MAX_GITHUB_FILE_SIZE = 100 * 1024 * 1024
VISUALS = {
    "Fruit": "FRUIT_VISUAL_PLU_BOOKLET_CORRECTED.docx",
    "Vegetable": "VEGETABLE_VISUAL_PLU_BOOKLET_CORRECTED.docx",
}
MAX_WORKBOOK_IMAGE_SIZE = (192, 120)
MAX_GITHUB_FILE_SIZE = 100 * 1024 * 1024
BLOCKS = (1, 6, 11)
LINES = 20
PAGE_ROWS = LINES + 3
PRODUCE_ROW_HEIGHT = 31
PRODUCE_IMAGE_BOX = (54, 38)
BULK_SHEETS = ("Bulk", "Bulk (2)", "Bulk (3)", "Bulk (4)")
STRETCH_SOURCE = Path("sources/binders/Time_To_Stretch.jpg")
STRETCH_SHEET = "Time To Stretch"
MISC_SHEET = "Misc Codes"
MISC_TITLE = "MISC CODES"
MISC_HEADER = '&C&"DejaVu Sans,Bold"&14&K305040' + MISC_TITLE
SOIL_SHEET = "Soil"
WEIGHED_SHEET = "Weighed more $100"
CONVERSIONS_SHEET = "Conversions"
CONVERSIONS_SOURCE = Path("sources/binders/Conversions_Updated.jpg")
CONVERSIONS_SOURCE_SHA256 = "e3a51d715ee020604a65041c046dbd4092ebe043c390c82c3f427e98775ae504"
CONVERSIONS_ACCURACY_REVIEW = (
    "Owner-uploaded image visually checked against all three original chart images: volume values, "
    "weight values, kilograms-to-pounds values, price tiers and rounding note match. The $4.99 price "
    "converts to $11.00/kg in both the original chart and uploaded image; a prior editable draft's "
    "$11.02 value was incorrect and is not used."
)
CONVERSION_IMAGE_HASHES = (
    "61982efebf133df82dff82c7cdd0851cbba7bc15dd40ea7279c1b6f9027b03ec",
    "6430fafaf191a266ba30d6de7fe66aac4a27ff5d0b0645168d70d9476b9b9fd1",
    "f9d0fa1844873f90904a105d56ec2b0d514fc0fb263c8557f24b9dfdc6bfe56b",
)
DIRECTORY_SHEET = "Store Directory"
DIRECTORY_SOURCE_HASHES = (
    "c5beab3173c87b3ab5876bcad374bb85e5cefc57ba616edb749791852cc3bc69",
    "0274edd4cd7893d0f087f73e380295bf562ac80c0e10a4022ae7f69b686418a3",
)
DIRECTORY_SOURCES = tuple(Path(f"sources/binders/Store_Directory_Original_{i}.png") for i in (1, 2))
DELIVERY_SHEET = "Delivery Charges"
DELIVERY_SOURCE_HASHES = (
    "f17f10fad1443053096d9de47c49c5e3f031fc47d1e91101b1a989c3b575c77a",
    "ab85fab3c2cb942b7ab40e9c9effa9f4083113882b52732f5d45510d7cd8207d",
)
DELIVERY_THEMED_HASHES = (
    "4ec5149eb8cb31d40846c21d5f8eb65cbe63b8dcbdde5b628df0a0c8e94e613e",
    "137a0598e8f1b21c952cfc6a6888b48c83ef818d813adf4cfc88003c737494fc",
)
DELIVERY_HEADER_BANDS = (
    ((320, 438, 1370, 496),),
    ((320, 530, 1370, 594), (320, 1275, 1370, 1340)),
)
DIRECTORY_HIGHLIGHTS = (
    ((75, 58, 460, 181),),
    (
    (202, 143, 449, 191), (1136, 135, 1293, 184), (1483, 460, 1563, 510),
    (4, 773, 151, 858), (520, 1118, 705, 1167), (1099, 1106, 1234, 1154),
    (1278, 1143, 1394, 1187),
    *((x - 17, 888, x + 17, 922) for x in (385, 452, 519, 616, 689, 744, 799, 856, 911)),
    ),
)
DIRECTORY_EMERGENCY_HIGHLIGHTS = (
    (973, 1233, 1633, 1273),
    (101, 319, 139, 364), (1032, 169, 1071, 213),
    (1463, 927, 1504, 972), (777, 1100, 815, 1144), (345, 1100, 384, 1144),
)
PDF_EXPORT_FORMAT = 'pdf:calc_pdf_Export:{"UseLosslessCompression":{"type":"boolean","value":"true"},' \
                    '"ReduceImageResolution":{"type":"boolean","value":"false"}}'


def sha(content):
    return hashlib.sha256(content).hexdigest()


def docx_compatible_image(content):
    try:
        DocxImage.from_blob(content)
    except UnrecognizedImageError:
        with Image.open(io.BytesIO(content)) as image:
            output = io.BytesIO()
            image.save(output, format="PNG")
        return output.getvalue()
    return content


def workbook_photo_copy(content):
    with Image.open(io.BytesIO(content)) as source:
        image = ImageOps.exif_transpose(source)
        image.thumbnail(MAX_WORKBOOK_IMAGE_SIZE, Image.Resampling.LANCZOS)
        if "A" in image.getbands() or "transparency" in image.info:
            has_transparency = image.convert("RGBA").getchannel("A").getextrema()[0] < 255
        else:
            has_transparency = False
        profile = {}
        if image.info.get("icc_profile"):
            profile["icc_profile"] = image.info["icc_profile"]
        output = io.BytesIO()
        if has_transparency:
            image.save(output, format="PNG", optimize=True, **profile)
        else:
            image.convert("RGB").save(
                output, format="JPEG", quality=90, optimize=True, progressive=True, **profile)
        return output.getvalue()


def xml(content):
    return ET.fromstring(content, ET.XMLParser(resolve_entities=False, no_network=True))


def serialize(element):
    return ET.tostring(element, encoding="UTF-8", xml_declaration=True, standalone=True)


def resolve_part(part, target):
    result = posixpath.normpath(target.lstrip("/") if target.startswith("/") else
                               posixpath.join(posixpath.dirname(part), target))
    if result.startswith("../"):
        raise ValueError(f"out-of-package relationship: {target}")
    return result


def rels_part(part):
    folder, filename = posixpath.split(part)
    return posixpath.join(folder, "_rels", filename + ".rels")


def sheet_info(archive, name="Produce"):
    workbook = xml(archive.read("xl/workbook.xml"))
    relationships = {rel.get("Id"): rel.get("Target") for rel in
                     xml(archive.read("xl/_rels/workbook.xml.rels"))}
    for index, sheet in enumerate(workbook.find(f"{{{S}}}sheets")):
        if sheet.get("name") == name:
            return resolve_part("xl/workbook.xml", relationships[sheet.get(f"{{{R}}}id")]), index
    raise ValueError(f"original binder has no {name!r} sheet")


def theme_operational_sheets(binder, destination):
    workbook = load_workbook(binder)
    donor = Workbook()
    edits = {}
    try:
        if SOIL_SHEET in workbook:
            styles_sheet = donor.active
            soil = workbook[SOIL_SHEET]
            heading_coordinate = next((coordinate for coordinate in ("B1", "E1")
                                       if soil[coordinate].value == "SOIL BARCODES"), None)
            if heading_coordinate is None:
                raise ValueError("Soil heading changed; review its layout before applying the theme")
            if any((cell.value is not None and cell.coordinate != heading_coordinate) or
                   cell.comment or cell.hyperlink for cell in (soil.cell(1, col) for col in range(2, 11))):
                raise ValueError("Soil heading area contains additional data; review before applying the theme")
            for merged in soil.merged_cells.ranges:
                if merged.min_row <= 1 <= merged.max_row and merged.min_col <= 10 and merged.max_col >= 2:
                    if str(merged) != "B1:J1":
                        raise ValueError("Soil heading area has an unexpected merge")
            for col in range(2, 11):
                cell = styles_sheet.cell(1, col)
                cell.fill = PatternFill("solid", fgColor="305040")
                if col == 2:
                    cell.font = Font(name="DejaVu Sans", size=13, bold=True, color="FFFFFF")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            edits[SOIL_SHEET] = {styles_sheet.cell(1, col).coordinate: styles_sheet.cell(1, col)
                                 for col in range(2, 11)}
        if WEIGHED_SHEET in workbook:
            styles_sheet = donor.create_sheet()
            edits[WEIGHED_SHEET] = {}
            for row in workbook[WEIGHED_SHEET]:
                for source in row:
                    if not isinstance(source.value, str):
                        continue
                    if source.row != 1 and not source.value.startswith(("STEP ", "IMPORTANT ")):
                        continue
                    target = styles_sheet[source.coordinate]
                    target.font = Font(name="DejaVu Sans", size=18, bold=True, color="305040")
                    target.alignment = copy.copy(source.alignment)
                    target.number_format = source.number_format
                    target.fill = copy.copy(source.fill)
                    target.border = copy.copy(source.border)
                    target.protection = copy.copy(source.protection)
                    edits[WEIGHED_SHEET][source.coordinate] = target
        content = io.BytesIO()
        donor.save(content)
    finally:
        workbook.close()
        donor.close()
    with zipfile.ZipFile(binder) as original, zipfile.ZipFile(io.BytesIO(content.getvalue())) as style_source:
        styles, style_ids = merge_styles(original.read("xl/styles.xml"), style_source.read("xl/styles.xml"))
        changes = {"xl/styles.xml": styles}
        for name, cells in edits.items():
            path, _ = sheet_info(original, name)
            sheet = xml(original.read(path))
            for view in sheet.findall(f"{{{S}}}sheetViews/{{{S}}}sheetView"):
                view.set("showGridLines", "0")
            if name == SOIL_SHEET and heading_coordinate == "E1":
                row = sheet.find(f"{{{S}}}sheetData/{{{S}}}row[@r='1']")
                heading = row.find(f"{{{S}}}c[@r='E1']")
                existing = row.find(f"{{{S}}}c[@r='B1']")
                if existing is not None:
                    row.remove(existing)
                heading.set("r", "B1")
                row.remove(heading)
                row.insert(0, heading)
            for coordinate, styled in cells.items():
                row_number = styled.row
                row = sheet.find(f"{{{S}}}sheetData/{{{S}}}row[@r='{row_number}']")
                cell = row.find(f"{{{S}}}c[@r='{coordinate}']")
                if cell is None:
                    cell = ET.Element(f"{{{S}}}c", r=coordinate)
                    following = next((existing for existing in row.findall(f"{{{S}}}c")
                                      if range_boundaries(existing.get("r"))[0] > styled.column), None)
                    if following is None:
                        row.append(cell)
                    else:
                        following.addprevious(cell)
                cell.set("s", str(style_ids[styled.style_id]))
            if name == SOIL_SHEET:
                merges = sheet.find(f"{{{S}}}mergeCells")
                if merges is None:
                    merges = ET.Element(f"{{{S}}}mergeCells")
                    preceding = {"sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
                                 "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter",
                                 "sortState", "dataConsolidate", "customSheetViews"}
                    index = next((i for i, child in enumerate(sheet)
                                  if ET.QName(child).localname not in preceding), len(sheet))
                    sheet.insert(index, merges)
                if not any(node.get("ref") == "B1:J1" for node in merges):
                    ET.SubElement(merges, f"{{{S}}}mergeCell", ref="B1:J1")
                merges.set("count", str(len(merges)))
                print_options = sheet.find(f"{{{S}}}printOptions")
                if print_options is None:
                    print_options = ET.Element(f"{{{S}}}printOptions")
                    margins = sheet.find(f"{{{S}}}pageMargins")
                    if margins is None:
                        raise ValueError("Soil sheet has no page margins")
                    margins.addprevious(print_options)
                print_options.set("horizontalCentered", "1")
            changes[path] = serialize(sheet)
        with zipfile.ZipFile(destination, "x", compression=zipfile.ZIP_DEFLATED) as output:
            output.comment = original.comment
            for entry in original.infolist():
                output.writestr(entry, changes.get(entry.filename, original.read(entry.filename)))


def move_floral_codes(binder, destination):
    workbook = load_workbook(binder)
    donor = Workbook()
    themed = donor.active
    try:
        misc = workbook[MISC_SHEET]
        headings = [cell for row in misc for cell in row if cell.value == "Floral Codes"]
        soil = workbook[SOIL_SHEET] if SOIL_SHEET in workbook else None
        if not headings:
            if soil is None:
                shutil.copyfile(binder, destination)
                return []
            if (soil["L1"].value != "Floral Codes" or
                    str(soil.print_area) != "'Soil'!$B$1:$M$66" or
                    "L1:M1" not in {str(item) for item in soil.merged_cells.ranges} or
                    soil.page_setup.scale != 70 or soil.page_setup.fitToWidth is not None or
                    soil.page_setup.fitToHeight is not None):
                raise ValueError("Floral Codes section is missing or ambiguous; review the template")
            entries = []
            for number in range(3, 67):
                code, name = soil[f"L{number}"], soil[f"M{number}"]
                if code.value is None and name.value is None:
                    continue
                if (number != len(entries) + 3 or code.value is None or
                        not isinstance(name.value, str) or name.data_type == "f" or code.data_type == "f" or
                        code.comment or name.comment or code.hyperlink or name.hyperlink):
                    raise ValueError("Existing Soil Floral Codes entries are incomplete or ambiguous")
                if code.data_type == "n" and code.number_format != "General" and not re.fullmatch(
                        r"0+", code.number_format):
                    raise ValueError("Existing Soil Floral Codes has an unsupported numeric format")
                entries.append(dict(code=code.value, name=name.value, code_type=code.data_type,
                                    number_format=code.number_format, soil_code=code.coordinate,
                                    soil_name=name.coordinate))
            if not entries:
                raise ValueError("Existing Soil Floral Codes section has no entries")
            names = [entry["name"] for entry in entries]
            if len(set(names)) != len(names) or any(cell.value in names for row in misc for cell in row
                                                   if isinstance(cell.value, str)):
                raise ValueError("Floral Codes entries are duplicated or remain in Misc Codes")
            shutil.copyfile(binder, destination)
            return entries
        if len(headings) != 1 or headings[0].column != 2 or soil is None:
            raise ValueError("Floral Codes section or Soil sheet is ambiguous; review the template")
        heading = headings[0]
        if any(misc.cell(heading.row, col).value is not None for col in (1, 3, 4)):
            raise ValueError("Floral Codes heading shares a row with other data; review the template")
        entries = []
        for base in (1, 3):
            number = heading.row + 1
            while misc.cell(number, base).value is not None or misc.cell(number, base + 1).value is not None:
                code, name = misc.cell(number, base), misc.cell(number, base + 1)
                if (code.value is None or not isinstance(name.value, str) or name.data_type == "f" or
                        code.data_type == "f" or code.comment or name.comment or code.hyperlink or name.hyperlink):
                    raise ValueError("Floral Codes has incomplete or unsupported entries; review the template")
                if code.data_type == "n" and code.number_format != "General" and not re.fullmatch(
                        r"0+", code.number_format):
                    raise ValueError("Floral Codes has an unsupported numeric format; review the template")
                entries.append((code, name))
                number += 1
        if not entries or len(entries) > 60:
            raise ValueError("Floral Codes cannot fit beside Soil; review the template")
        last_row = max(name.row for _, name in entries)
        if any(cell.value is not None or cell.comment or cell.hyperlink
               for row in misc.iter_rows(min_row=last_row + 1) for cell in row):
            raise ValueError("Unexpected content after Floral Codes; review the template")
        if (len({name.value for _, name in entries}) != len(entries) or
                any(cell.value is not None or cell.comment or cell.hyperlink
                    for row in misc.iter_rows(min_row=heading.row, max_row=last_row, min_col=5)
                    for cell in row)):
            raise ValueError("Floral Codes section contains additional or duplicate data")
        if any(cell.comment or cell.hyperlink for row in misc.iter_rows(
                min_row=heading.row, max_row=last_row, max_col=4) for cell in row):
            raise ValueError("Floral Codes section contains comments or hyperlinks")
        if heading.comment or heading.hyperlink or misc.merged_cells.ranges:
            raise ValueError("Floral Codes contains unsupported metadata or merged cells; review the template")
        for sheet in workbook:
            for row in sheet:
                for cell in row:
                    if cell.data_type == "f" and (sheet == misc or "misc codes" in cell.value.lower()):
                        raise ValueError("Misc Codes formulas require review before moving Floral Codes")
        if any("misc codes" in (definition.attr_text or "").lower()
               for definition in workbook.defined_names.values()):
            raise ValueError("Misc Codes defined names require review before moving Floral Codes")
        if any(cell.value is not None or cell.comment or cell.hyperlink or cell.has_style
               for row in soil.iter_rows(min_col=11, max_col=13, min_row=1, max_row=66) for cell in row):
            raise ValueError("Soil floral destination contains data; review the template")
        if any(r.max_col >= 11 and r.min_row <= 66 for r in soil.merged_cells.ranges):
            raise ValueError("Soil floral destination contains merged cells; review the template")
        if any(p.anchor._from.col >= 10 or
               (getattr(p.anchor, "to", None) is not None and p.anchor.to.col >= 10)
               for p in soil._images):
            raise ValueError("Soil images overlap the floral destination; review the template")
        if any(dim.min is not None and dim.min <= 13 and (dim.max or dim.min) >= 11
               for dim in soil.column_dimensions.values()):
            raise ValueError("Soil floral destination has custom columns; review the template")
        if (str(soil.print_area) != "'Soil'!$B$1:$J$66" or soil.page_setup.scale != 70 or
                soil.page_setup.fitToWidth is not None or soil.page_setup.fitToHeight is not None):
            raise ValueError("Soil print layout changed; review it before moving Floral Codes")
        themed["L1"].font = Font(name="DejaVu Sans", size=13, bold=True, color="FFFFFF")
        themed["L1"].alignment = Alignment(horizontal="center", vertical="center")
        themed["L1"].fill = PatternFill("solid", fgColor=DARK_GREEN)
        themed["M1"].fill = PatternFill("solid", fgColor=DARK_GREEN)
        for index, (code, name) in enumerate(entries):
            for col, original in ((12, code), (13, name)):
                cell = themed.cell(index + 3, col)
                cell.font = Font(name="DejaVu Sans", size=12, bold=col == 12)
                cell.alignment = Alignment(vertical="center", horizontal="center" if col == 12 else "left")
                cell.number_format = original.number_format
                cell.protection = copy.copy(original.protection)
                cell.fill = row_fill(index)
        content = io.BytesIO()
        donor.save(content)
        source_coordinates = [(code.coordinate, name.coordinate) for code, name in entries]
        records = [dict(code=code.value, name=name.value, code_type=code.data_type,
                        number_format=code.number_format, source_code=code.coordinate,
                        source_name=name.coordinate, soil_code=f"L{index + 3}", soil_name=f"M{index + 3}")
                   for index, (code, name) in enumerate(entries)]
        heading_coordinate = heading.coordinate
    finally:
        workbook.close()
        donor.close()
    with zipfile.ZipFile(binder) as original, zipfile.ZipFile(io.BytesIO(content.getvalue())) as style_source:
        soil_path, soil_index = sheet_info(original, SOIL_SHEET)
        misc_path, _ = sheet_info(original, MISC_SHEET)
        soil_xml, misc_xml = xml(original.read(soil_path)), xml(original.read(misc_path))
        styles, style_ids = merge_styles(original.read("xl/styles.xml"), style_source.read("xl/styles.xml"))
        misc_data, soil_data = misc_xml.find(f"{{{S}}}sheetData"), soil_xml.find(f"{{{S}}}sheetData")

        def move_cell(source_coordinate, target_coordinate):
            source_row = misc_data.find(f"{{{S}}}row[@r='{range_boundaries(source_coordinate)[1]}']")
            node = source_row.find(f"{{{S}}}c[@r='{source_coordinate}']")
            if node is None:
                raise ValueError(f"Missing Floral Codes XML cell {source_coordinate}")
            source_row.remove(node)
            node.set("r", target_coordinate)
            node.set("s", str(style_ids[themed[target_coordinate].style_id]))
            target_row = soil_data.find(f"{{{S}}}row[@r='{range_boundaries(target_coordinate)[1]}']")
            if target_row is None:
                raise ValueError(f"Soil row for {target_coordinate} is missing; review the template")
            target_row.append(node)

        move_cell(heading_coordinate, "L1")
        for index, (code, name) in enumerate(source_coordinates):
            move_cell(code, f"L{index + 3}")
            move_cell(name, f"M{index + 3}")
        merges = soil_xml.find(f"{{{S}}}mergeCells")
        if merges is None:
            raise ValueError("Soil heading merge is missing")
        ET.SubElement(merges, f"{{{S}}}mergeCell", ref="L1:M1")
        merges.set("count", str(len(merges)))
        column_widths = soil_xml.find(f"{{{S}}}cols")
        if column_widths is None:
            column_widths = ET.Element(f"{{{S}}}cols")
            soil_data.addprevious(column_widths)
        for index, width in ((11, 2), (12, 10), (13, 30)):
            ET.SubElement(column_widths, f"{{{S}}}col", min=str(index), max=str(index),
                          width=str(width), customWidth="1")
        soil_xml.find(f"{{{S}}}dimension").set("ref", "B1:M66")
        book = xml(original.read("xl/workbook.xml"))
        names = book.find(f"{{{S}}}definedNames")
        areas = [node for node in names if node.get("name") == "_xlnm.Print_Area" and
                 node.get("localSheetId") == str(soil_index)]
        if len(areas) != 1:
            raise ValueError("Soil print area is missing or ambiguous")
        areas[0].text = "'Soil'!$B$1:$M$66"
        changes = {soil_path: serialize(soil_xml), misc_path: serialize(misc_xml),
                   "xl/styles.xml": styles, "xl/workbook.xml": serialize(book)}
        with zipfile.ZipFile(destination, "x", compression=zipfile.ZIP_DEFLATED) as output:
            output.comment = original.comment
            for entry in original.infolist():
                output.writestr(entry, changes.get(entry.filename, original.read(entry.filename)))
    return records


def fit_misc_sheet(binder, destination):
    workbook = load_workbook(binder)
    donor = Workbook()
    themed = donor.active
    styled = {}
    try:
        source = workbook[MISC_SHEET]
        indices = {1: 0, 3: 0}
        for row in source.iter_rows():
            number = row[0].row
            for base in (1, 3):
                left, right = source.cell(number, base), source.cell(number, base + 1)
                heading = ((number == 1 and
                            (isinstance(left.value, str) or isinstance(right.value, str))) or
                           (isinstance(right.value, str) and right.data_type != "f" and left.value is None))
                active = left.value is not None or right.value is not None
                if (base == 3 and not active and isinstance(source.cell(number, 2).value, str) and
                        source.cell(number, 1).value is None and
                        any(source.cell(number + 1, col).value is not None for col in (3, 4))):
                    heading = True
                if heading:
                    indices[base] = 0
                fill = (PatternFill("solid", fgColor=DARK_GREEN) if heading else
                        row_fill(indices[base]) if active else PatternFill())
                if active and not heading:
                    indices[base] += 1
                for cell in (left, right):
                    target = themed[cell.coordinate]
                    target.value = cell.value
                    target.alignment = copy.copy(cell.alignment)
                    target.number_format = cell.number_format
                    target.protection = copy.copy(cell.protection)
                    font = copy.copy(cell.font)
                    font.name = "DejaVu Sans"
                    font.sz = 10
                    font.color = "FFFFFF" if heading else "000000"
                    font.bold = heading or cell.column == base or bool(font.bold)
                    target.font = font
                    target.fill = fill
                    target.border = (Border(bottom=Side(style="hair", color="D0D0D0"))
                                     if active and not heading else Border())
                    styled[cell.coordinate] = target
        content = io.BytesIO()
        donor.save(content)
    finally:
        workbook.close()
        donor.close()
    with zipfile.ZipFile(binder) as original:
        path, _ = sheet_info(original, MISC_SHEET)
        sheet = xml(original.read(path))
        with zipfile.ZipFile(io.BytesIO(content.getvalue())) as styles_source:
            styles, style_ids = merge_styles(original.read("xl/styles.xml"), styles_source.read("xl/styles.xml"))
        for row in sheet.findall(f"{{{S}}}sheetData/{{{S}}}row"):
            number = int(row.get("r"))
            for column in range(1, 5):
                coordinate = f"{get_column_letter(column)}{number}"
                if coordinate not in styled:
                    continue
                cell = row.find(f"{{{S}}}c[@r='{coordinate}']")
                if cell is None:
                    cell = ET.Element(f"{{{S}}}c", r=coordinate)
                    following = next((node for node in row.findall(f"{{{S}}}c")
                                      if range_boundaries(node.get("r"))[0] > column), None)
                    if following is None:
                        row.append(cell)
                    else:
                        following.addprevious(cell)
                cell.set("s", str(style_ids[styled[coordinate].style_id]))
        for view in sheet.findall(f"{{{S}}}sheetViews/{{{S}}}sheetView"):
            view.set("showGridLines", "0")
        properties = sheet.find(f"{{{S}}}sheetPr")
        if properties is None:
            properties = ET.Element(f"{{{S}}}sheetPr")
            sheet.insert(0, properties)
        page_properties = properties.find(f"{{{S}}}pageSetUpPr")
        if page_properties is None:
            page_properties = ET.SubElement(properties, f"{{{S}}}pageSetUpPr")
        page_properties.set("fitToPage", "1")
        setup = sheet.find(f"{{{S}}}pageSetup")
        if setup is None:
            margins = sheet.find(f"{{{S}}}pageMargins")
            if margins is None:
                raise ValueError("Misc Codes sheet has no page margins")
            setup = ET.Element(f"{{{S}}}pageSetup")
            margins.addnext(setup)
        setup.set("orientation", "portrait")
        setup.set("paperSize", "1")
        setup.set("fitToWidth", "1")
        setup.set("fitToHeight", "1")
        setup.attrib.pop("scale", None)
        header_footer = sheet.find(f"{{{S}}}headerFooter")
        if header_footer is None:
            header_footer = ET.Element(f"{{{S}}}headerFooter")
            setup.addnext(header_footer)
        if any(header_footer.get(name) in ("1", "true") for name in ("differentFirst", "differentOddEven")):
            raise ValueError("Misc Codes has alternate print headers; review them before theming")
        header = header_footer.find(f"{{{S}}}oddHeader")
        if header is None:
            header = ET.SubElement(header_footer, f"{{{S}}}oddHeader")
        if header.text not in (None, "", MISC_HEADER, MISC_HEADER.replace("&14&K", "&14 &K")):
            raise ValueError("Misc Codes has an existing print header; review it before theming")
        header.text = MISC_HEADER
        for tag in ("rowBreaks", "colBreaks"):
            breaks = sheet.find(f"{{{S}}}{tag}")
            if breaks is not None:
                sheet.remove(breaks)
        with zipfile.ZipFile(destination, "x", compression=zipfile.ZIP_DEFLATED) as output:
            output.comment = original.comment
            for entry in original.infolist():
                output.writestr(entry, serialize(sheet) if entry.filename == path else
                                styles if entry.filename == "xl/styles.xml" else original.read(entry.filename))


def theme_conversions_sheet(binder, destination, conversion_image):
    if sha(conversion_image) != CONVERSIONS_SOURCE_SHA256:
        raise ValueError("Uploaded Conversions source changed; review its accuracy before rebuilding")
    with Image.open(io.BytesIO(conversion_image)) as image:
        if image.format != "JPEG" or image.size != (1024, 1536):
            raise ValueError("Uploaded Conversions image dimensions or format changed; review before rebuilding")
    workbook = load_workbook(binder)
    try:
        source = workbook[CONVERSIONS_SHEET]
        if (len(source._images) == 1 and
                sha(source._images[0]._data()) == CONVERSIONS_SOURCE_SHA256):
            if any(cell.value is not None for row in source for cell in row) or source._charts:
                raise ValueError("Updated Conversions image sheet has unexpected data; review before rebuilding")
            shutil.copyfile(binder, destination)
            return
        if not source._images and source.sheet_properties.codeName == "UpdatedConversionsImage":
            raise ValueError("Updated Conversions image is missing; review the template")
        if not source._images and source.sheet_properties.codeName != "EditableConversions":
            raise ValueError("Conversions sheet is not a recognized source or generated image sheet")
        if source._images and (len(source._images) != 3 or
                               tuple(sha(image._data()) for image in source._images) != CONVERSION_IMAGE_HASHES):
            raise ValueError("Conversions source charts changed; review their accuracy before replacing")
        if source._charts or source.merged_cells.ranges:
            raise ValueError("Conversions source contains unsupported content; review before conversion")
        if source._images:
            if any(cell.comment or cell.hyperlink or
                   (cell.value is not None and not (cell.column in (8, 9) and 22 <= cell.row <= 33))
                   for row in source for cell in row):
                raise ValueError("Conversions source contains additional cell content; review before conversion")
            if (source["I22"].value != "Weight              " or
                    source["H23"].value != "Kilograms" or source["I23"].value != "Pounds" or
                    source["H22"].value is not None or
                    any(source.cell(row, col).value is None for row in range(24, 34) for col in (8, 9))):
                raise ValueError("Conversions source weight table changed; review before conversion")
    finally:
        workbook.close()
    donor = Workbook()
    sheet = donor.active
    sheet.title = CONVERSIONS_SHEET
    sheet.sheet_properties.codeName = "UpdatedConversionsImage"
    sheet.sheet_view.showGridLines = False
    sheet.column_dimensions["A"].width = 100
    sheet.sheet_format.defaultRowHeight = 15
    sheet.page_setup.orientation = "portrait"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_LETTER
    sheet.page_setup.fitToWidth = sheet.page_setup.fitToHeight = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_margins = PageMargins(left=.25, right=.25, top=.25, bottom=.25, header=.1, footer=.1)
    sheet.print_options.horizontalCentered = True
    sheet.print_area = "A1:A52"
    picture = ExcelImage(io.BytesIO(conversion_image))
    picture.width, picture.height = 700, 1050
    picture.anchor = OneCellAnchor(
        _from=AnchorMarker(col=0, row=0, colOff=pixels_to_EMU(100)),
        ext=XDRPositiveSize2D(cx=pixels_to_EMU(picture.width), cy=pixels_to_EMU(picture.height)))
    sheet.add_image(picture)
    content = io.BytesIO()
    donor.save(content)
    donor.close()
    replace_sheet(binder, content.getvalue(), destination, CONVERSIONS_SHEET)


def misc_print_text(path):
    workbook = load_workbook(path, data_only=True)
    try:
        sheet = workbook[MISC_SHEET]
        if (sheet.page_setup.orientation != "portrait" or str(sheet.page_setup.paperSize) != "1" or
                sheet.page_setup.fitToWidth != 1 or sheet.page_setup.fitToHeight != 1 or
                not sheet.sheet_properties.pageSetUpPr.fitToPage or
                sheet.row_breaks.brk or sheet.col_breaks.brk):
            raise ValueError("Misc Codes must be configured to fit one portrait US Letter page")
        text = [" ".join(cell.value.split()) for row in sheet for cell in row
                if isinstance(cell.value, str) and cell.value.strip()]
        if not text:
            raise ValueError("Misc Codes has no text to verify in the printed page")
        numbers = Counter(Decimal(str(cell.value)) for row in sheet for cell in row
                          if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool))
        return text, numbers
    finally:
        workbook.close()


def misc_page_matches(page, expected):
    text, numbers = expected
    printed = " ".join((page.extract_text() or "").split())
    found = Counter(Decimal(value.replace(",", "")) for value in
                    re.findall(r"(?<![\d.])-?\d[\d,]*(?:\.\d+)?(?![\d.])", printed))
    return all(value in printed for value in text) and not numbers - found


def validate_conversions_image(sheet, image):
    if (len(sheet._images) != 1 or sha(sheet._images[0]._data()) != sha(image) or
            any(cell.value is not None for row in sheet for cell in row) or sheet._charts or
            sheet.sheet_properties.codeName != "UpdatedConversionsImage" or
            sheet.page_setup.orientation != "portrait" or str(sheet.page_setup.paperSize) != "1" or
            sheet.page_setup.fitToWidth != 1 or sheet.page_setup.fitToHeight != 1 or
            not sheet.sheet_properties.pageSetUpPr.fitToPage):
        raise ValueError("Conversions image or one-page print settings differ from the approved source")


def validate_directory_template(archive):
    part, _ = sheet_info(archive, DIRECTORY_SHEET)
    sheet = xml(archive.read(part))
    allowed = {"sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
               "pageMargins", "pageSetup", "rowBreaks", "colBreaks", "drawing", "printOptions"}
    if any(ET.QName(node).localname not in allowed for node in sheet):
        raise ValueError("Store Directory has unsupported sheet features; review the template")
    drawings = sheet.findall(f"{{{S}}}drawing")
    if len(drawings) != 1:
        raise ValueError("Store Directory must have one image drawing")
    rels = xml(archive.read(rels_part(part)))
    if any(item.get("TargetMode") == "External" or
           not item.get("Type", "").endswith(("/drawing", "/printerSettings")) for item in rels):
        raise ValueError("Store Directory has unsupported relationships; review the template")
    relation = [item for item in rels if item.get("Id") == drawings[0].get(f"{{{R}}}id")]
    if len(relation) != 1 or not relation[0].get("Type", "").endswith("/drawing"):
        raise ValueError("Store Directory image drawing relationship is invalid")
    drawing = xml(archive.read(resolve_part(part, relation[0].get("Target"))))
    xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    drawingml = "http://schemas.openxmlformats.org/drawingml/2006/main"
    if (len(drawing) != 2 or any(
            ET.QName(anchor).localname not in {"oneCellAnchor", "twoCellAnchor"} or
            len(anchor.findall(f"{{{xdr}}}pic")) != 1 or
            any(ET.QName(node).localname not in {"from", "to", "ext", "pic", "clientData"} for node in anchor)
            for anchor in drawing)):
        raise ValueError("Store Directory has unsupported drawing objects; review the template")
    for crop in drawing.iter(f"{{{drawingml}}}srcRect"):
        if any(int(value) != 0 for value in crop.attrib.values()):
            raise ValueError("Store Directory has cropped images; review the template")
    for transform in drawing.iter(f"{{{drawingml}}}xfrm"):
        if any(transform.get(name, "0") not in {"0", "false"} for name in ("rot", "flipH", "flipV")):
            raise ValueError("Store Directory has rotated or mirrored images; review the template")
    if any(ET.QName(node).localname in {"hlinkClick", "hlinkHover"} for node in drawing.iter()):
        raise ValueError("Store Directory image hyperlinks require review")


def themed_directory_image(content, index, emergency_highlights=True, parking_labels=True, row_stripes=True):
    if sha(content) != DIRECTORY_SOURCE_HASHES[index]:
        raise ValueError(f"Store Directory page {index + 1} changed; review its map highlights before rebuilding")
    with Image.open(io.BytesIO(content)) as image:
        original = image.convert("RGB").transpose(Image.Transpose.ROTATE_270)
    highlights = Image.new("RGB", original.size, "white")
    shape = ImageDraw.Draw(highlights)
    regions = Image.new("L", original.size)
    region_draw = ImageDraw.Draw(regions)
    if index == 0 and row_stripes:
        for left, right, top, count in ((80, 560, 190, 35), (575, 1060, 72, 39),
                                         (1080, 1555, 71, 39), (1570, 2050, 72, 36)):
            for row in range(1, count, 2):
                y = round(top + row * 30.1)
                box = (left, y, right, y + 28)
                shape.rectangle(box, fill="#" + PALE_GREEN)
                region_draw.rectangle(box, fill=255)
    for box in DIRECTORY_HIGHLIGHTS[index]:
        shape.rounded_rectangle(box, radius=8, fill="#DCE6DA", outline="#305040", width=2)
        region_draw.rounded_rectangle(box, radius=8, fill=255)
    if index == 1 and emergency_highlights:
        for box in DIRECTORY_EMERGENCY_HIGHLIGHTS:
            shape.rounded_rectangle(box, radius=4, fill="#FCE4D6", outline="#9C0006", width=2)
            region_draw.rounded_rectangle(box, radius=4, fill=255)
    if index == 1 and parking_labels:
        font = ImageFont.truetype("DejaVuSans-Bold.ttf", 24)
        for box, text, fill, outline in (
            ((580, 1226, 900, 1270), "Front Parking Lot", "#DCE6DA", "#305040"),
            ((1504, 620, 1644, 725), "Back\nParking\nLot", "#FCE4D6", "#9C0006"),
        ):
            shape.rounded_rectangle(box, radius=8, fill=fill, outline=outline, width=2)
            region_draw.rounded_rectangle(box, radius=8, fill=255)
            bounds = shape.multiline_textbbox((0, 0), text, font=font, spacing=4, align="center")
            x = (box[0] + box[2] - bounds[2] - bounds[0]) / 2
            y = (box[1] + box[3] - bounds[3] - bounds[1]) / 2
            shape.multiline_text((x, y), text, font=font, spacing=4, align="center", fill="black")
    white = [channel.point(lambda value: 255 if value == 255 else 0)
             for channel in original.split()]
    white_mask = ImageChops.multiply(ImageChops.multiply(white[0], white[1]), white[2])
    original.paste(highlights, mask=ImageChops.multiply(white_mask, regions))
    metadata = PngImagePlugin.PngInfo()
    metadata.add_text("store_directory_source_sha256", DIRECTORY_SOURCE_HASHES[index])
    output = io.BytesIO()
    original.save(output, format="PNG", pnginfo=metadata, optimize=True)
    return output.getvalue()


def themed_delivery_image(content, index):
    if not 0 <= index < len(DELIVERY_SOURCE_HASHES):
        raise ValueError(f"Delivery Charges page index is invalid: {index}")
    expected_hash = DELIVERY_SOURCE_HASHES[index]
    if sha(content) != expected_hash:
        raise ValueError(f"Delivery Charges page {index + 1} changed; review its barcode content")
    with Image.open(io.BytesIO(content)) as image:
        if image.format != "PNG" or image.size != (1700, 2200):
            raise ValueError(f"Delivery Charges page {index + 1} has unexpected dimensions or format")
        original = image.convert("RGB")
    pixels = original.load()
    green = (48, 80, 64)
    for left, top, right, bottom in DELIVERY_HEADER_BANDS[index]:
        for y in range(top, bottom):
            for x in range(left, right):
                luminance = sum(pixels[x, y]) / 3
                blend = luminance / 255
                pixels[x, y] = tuple(round(255 * (1 - blend) + color * blend) for color in green)
    metadata = PngImagePlugin.PngInfo()
    metadata.add_text("delivery_charges_source_sha256", expected_hash)
    output = io.BytesIO()
    original.save(output, format="PNG", pnginfo=metadata, optimize=True)
    return output.getvalue()


def delivery_sheet(sheet):
    if (len(sheet._images) != 2 or sheet._charts or sheet.merged_cells.ranges or
            any(cell.value is not None for row in sheet for cell in row) or
            any(cell.comment or cell.hyperlink for row in sheet for cell in row)):
        raise ValueError("Delivery Charges must contain exactly two page images and no cell data")
    pictures = sorted(sheet._images, key=lambda image: (image.anchor._from.row, image.anchor._from.col))
    if (pictures[0].anchor._from.row == pictures[1].anchor._from.row or
            pictures[0].anchor._from.row >= pictures[1].anchor._from.row):
        raise ValueError("Delivery Charges page order is ambiguous")
    supplied = [image._data() for image in pictures]
    images = []
    for index, content in enumerate(supplied):
        if sha(content) == DELIVERY_SOURCE_HASHES[index]:
            themed = themed_delivery_image(content, index)
            if sha(themed) != DELIVERY_THEMED_HASHES[index]:
                raise ValueError(f"Delivery Charges theme for page {index + 1} changed; review it before rebuilding")
            images.append(themed)
        elif sha(content) == DELIVERY_THEMED_HASHES[index]:
            images.append(content)
        else:
            raise ValueError(f"Delivery Charges page {index + 1} changed; review its barcode content")
    workbook = Workbook()
    result = workbook.active
    result.title = DELIVERY_SHEET
    result.sheet_view.showGridLines = False
    result.column_dimensions["A"].width = 105
    result.sheet_format.defaultRowHeight = 15
    result.page_setup.orientation = "portrait"
    result.page_setup.paperSize = result.PAPERSIZE_LETTER
    result.page_setup.scale = 100
    result.sheet_properties.pageSetUpPr.fitToPage = False
    result.page_margins = PageMargins(left=.25, right=.25, top=.25, bottom=.25, header=.1, footer=.1)
    result.row_breaks.append(Break(id=49))
    result.print_area = "A1:A100"
    for index, content in enumerate(images):
        picture = ExcelImage(io.BytesIO(content))
        picture.width, picture.height = 728, 942
        picture.anchor = OneCellAnchor(
            _from=AnchorMarker(col=0, row=50 * index, colOff=pixels_to_EMU(6)),
            ext=XDRPositiveSize2D(cx=pixels_to_EMU(picture.width), cy=pixels_to_EMU(picture.height)))
        result.add_image(picture)
    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue(), images


def validate_delivery(sheet, images):
    if (len(images) != 2 or len(sheet._images) != 2 or
            [sha(image._data()) for image in sheet._images] != [sha(content) for content in images]):
        raise ValueError("Delivery Charges page images differ from the reviewed source derivatives")
    if (sheet.sheet_view.showGridLines or sheet.page_setup.orientation != "portrait" or
            str(sheet.page_setup.paperSize) != "1" or sheet.page_setup.scale != 100 or
            sheet.sheet_properties.pageSetUpPr.fitToPage or
            [item.id for item in sheet.row_breaks.brk if item.man] != [49] or
            str(sheet.print_area) != "'Delivery Charges'!$A$1:$A$100" or
            [picture.anchor._from.row for picture in sheet._images] != [0, 50] or
            any((picture.anchor.ext.cx, picture.anchor.ext.cy) !=
                (pixels_to_EMU(728), pixels_to_EMU(942)) for picture in sheet._images)):
        raise ValueError("Delivery Charges must print as two ordered portrait US Letter pages")


def directory_sheet(sheet, originals=None):
    if (len(sheet._images) != 2 or sheet._charts or sheet.merged_cells.ranges or
            any(cell.value is not None and (not isinstance(cell.value, str) or cell.value.strip())
                for row in sheet for cell in row) or
            any(cell.comment or cell.hyperlink for row in sheet for cell in row)):
        raise ValueError("Store Directory must contain exactly two images and no other content; review the template")
    pictures = sorted(sheet._images, key=lambda image: (image.anchor._from.row, image.anchor._from.col))
    if pictures[0].anchor._from.row == pictures[1].anchor._from.row:
        raise ValueError("Store Directory image order is ambiguous")
    supplied = [image._data() for image in pictures]
    if originals is None:
        originals = [(ROOT / path).read_bytes() for path in DIRECTORY_SOURCES]
    if len(originals) != 2:
        raise ValueError("Store Directory needs two original page images")
    images = [themed_directory_image(content, index) for index, content in enumerate(originals)]
    for index, content in enumerate(supplied):
        if content not in (originals[index], images[index],
                           themed_directory_image(originals[index], index, parking_labels=False),
                           themed_directory_image(originals[index], index, emergency_highlights=False,
                                                  parking_labels=False),
                           themed_directory_image(originals[index], index, row_stripes=False),
                           themed_directory_image(originals[index], index, parking_labels=False, row_stripes=False),
                           themed_directory_image(originals[index], index, emergency_highlights=False,
                                                  parking_labels=False, row_stripes=False)):
            raise ValueError(f"Store Directory page {index + 1} changed; review its map highlights before rebuilding")
    workbook = Workbook()
    result = workbook.active
    result.title = DIRECTORY_SHEET
    result.sheet_view.showGridLines = False
    result.column_dimensions["A"].width = 135
    result.sheet_format.defaultRowHeight = 15
    result.page_setup.orientation = "landscape"
    result.page_setup.paperSize = result.PAPERSIZE_LETTER
    result.page_setup.scale = 100
    result.sheet_properties.pageSetUpPr.fitToPage = False
    result.page_margins = PageMargins(left=.25, right=.25, top=.4, bottom=.4, header=.1, footer=.1)
    result.row_breaks.append(Break(id=36))
    result.print_area = "A1:A72"
    for index, content in enumerate(images):
        with Image.open(io.BytesIO(content)) as image:
            if image.width <= 0 or image.height <= 0:
                raise ValueError("Store Directory image has invalid dimensions")
            scale = min(930 / image.width, 675 / image.height)
            width, height = image.width * scale, image.height * scale
        picture = ExcelImage(io.BytesIO(content))
        picture.anchor = OneCellAnchor(
            _from=AnchorMarker(col=0, row=37 * index, colOff=pixels_to_EMU((930 - width) / 2)),
            ext=XDRPositiveSize2D(cx=pixels_to_EMU(width), cy=pixels_to_EMU(height)))
        result.add_image(picture)
    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue(), images


def directory_image_pixels(content):
    with Image.open(io.BytesIO(content)) as image:
        return image.convert("RGB").tobytes()


def validate_directory(sheet, images):
    if len(sheet._images) != 2 or [sha(image._data()) for image in sheet._images] != [sha(b) for b in images]:
        raise ValueError("Store Directory images differ from the themed derivatives")
    if (sheet.page_setup.orientation != "landscape" or str(sheet.page_setup.paperSize) != "1" or
            sheet.page_setup.scale != 100 or
            [b.id for b in sheet.row_breaks.brk if b.man] != [36]):
        raise ValueError("Store Directory must have two landscape US Letter pages")


def combine_bulk(binder, destination):
    """Combine the four original print regions without resaving unrelated workbook parts."""
    with zipfile.ZipFile(binder) as original:
        workbook = xml(original.read("xl/workbook.xml"))
        sheets = workbook.find(f"{{{S}}}sheets")
        sheet_names = [node.get("name") for node in sheets]
        present = [name for name in BULK_SHEETS if name in sheet_names]
        if not any(name in present for name in BULK_SHEETS[1:]):
            shutil.copyfile(binder, destination)
            return False
        if present != list(BULK_SHEETS) or [sheet_names.index(n) for n in BULK_SHEETS] != list(
                range(sheet_names.index("Bulk"), sheet_names.index("Bulk") + 4)):
            raise ValueError("the four Bulk sheets must be present and consecutive")
        paths = [sheet_info(original, name)[0] for name in BULK_SHEETS]
        trees = [xml(original.read(path)) for path in paths]
        permitted = {"sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
                     "mergeCells", "pageMargins", "pageSetup"}
        for path, tree in zip(paths, trees):
            if any(ET.QName(node).localname not in permitted for node in tree):
                raise ValueError(f"{path}: unsupported Bulk sheet feature; review before merging")
            relations = rels_part(path)
            if relations in original.namelist():
                for rel in xml(original.read(relations)):
                    if not rel.get("Type", "").endswith("/printerSettings"):
                        raise ValueError(f"{path}: unsupported Bulk relationship; review before merging")
        for tree in trees[1:]:
            for tag in ("cols", "sheetFormatPr", "pageMargins"):
                if ET.tostring(trees[0].find(f"{{{S}}}{tag}"), method="c14n") != ET.tostring(
                        tree.find(f"{{{S}}}{tag}"), method="c14n"):
                    raise ValueError(f"Bulk sheets have incompatible {tag}")
        names = workbook.find(f"{{{S}}}definedNames")
        indexes = [sheet_names.index(name) for name in BULK_SHEETS]
        areas = []
        for index in indexes:
            found = [n for n in names if n.get("localSheetId") == str(index) and
                     n.get("name") == "_xlnm.Print_Area"] if names is not None else []
            if len(found) != 1 or not found[0].text or "!" not in found[0].text:
                raise ValueError("each Bulk sheet must have one rectangular print area")
            bounds = range_boundaries(found[0].text.split("!", 1)[1].replace("$", ""))
            if bounds[0] != 2 or bounds[1] != 4 or bounds[2] != 9:
                raise ValueError("unexpected Bulk print area; review before merging")
            areas.append(bounds)
        # Do not silently drop owner edits outside the four printed regions.
        headers = []
        for tree, (_, start, _, end) in zip(trees, areas):
            header = []
            for row in tree.find(f"{{{S}}}sheetData"):
                for cell in row:
                    col, r, _, _ = range_boundaries(cell.get("r"))
                    if cell.find(f"{{{S}}}f") is not None:
                        raise ValueError("Bulk formulas require manual review before merging")
                    if cell.find(f"{{{S}}}v") is not None or cell.find(f"{{{S}}}is") is not None:
                        if r == 1 and col == 2:
                            header.append(ET.tostring(cell, method="c14n"))
                        elif not (start <= r <= end and 2 <= col <= 9):
                            raise ValueError("Bulk contains data outside its print area")
            headers.append(header)
        if any(header != headers[0] for header in headers[1:]):
            raise ValueError("Bulk sheet headers differ; review before merging")
        for path in original.namelist():
            if path.startswith("xl/worksheets/") and path.endswith(".xml") and path not in paths:
                tree = xml(original.read(path))
                for formula in tree.iter(f"{{{S}}}f"):
                    if formula.text and any(name in formula.text for name in BULK_SHEETS[1:]):
                        raise ValueError(f"{path}: formula references a Bulk sheet being merged")
        target = trees[0]
        data = target.find(f"{{{S}}}sheetData")
        for row in list(data):
            if int(row.get("r")) > areas[0][3]:
                data.remove(row)
        merges = target.find(f"{{{S}}}mergeCells")
        last = areas[0][3]
        breaks = []
        for tree, (_, start, _, end) in zip(trees[1:], areas[1:]):
            breaks.append(last)
            offset = last + 1 - start
            for row in tree.find(f"{{{S}}}sheetData"):
                row_number = int(row.get("r"))
                if start <= row_number <= end:
                    node = copy.deepcopy(row)
                    node.set("r", str(row_number + offset))
                    for cell in node:
                        col, r, _, _ = range_boundaries(cell.get("r"))
                        cell.set("r", f"{get_column_letter(col)}{r + offset}")
                    data.append(node)
            incoming = tree.find(f"{{{S}}}mergeCells")
            if incoming is not None:
                for merge in incoming:
                    min_col, min_row, max_col, max_row = range_boundaries(merge.get("ref"))
                    if min_row == 1 and max_row == 2:
                        continue
                    if min_row < start or max_row > end:
                        raise ValueError("Bulk merged cells extend beyond the print area")
                    if merges is None:
                        raise ValueError("Bulk merge layout requires manual review")
                    node = copy.deepcopy(merge)
                    node.set("ref", f"{get_column_letter(min_col)}{min_row + offset}:"
                                    f"{get_column_letter(max_col)}{max_row + offset}")
                    merges.append(node)
            last += end - start + 1
        if merges is not None:
            merges.set("count", str(len(merges)))
        target.find(f"{{{S}}}dimension").set("ref", f"B1:I{last}")
        setup = target.find(f"{{{S}}}pageSetup")
        setup.attrib.clear()
        setup.set("paperSize", "1")
        setup.set("orientation", "landscape")
        setup.set("scale", "65")
        target.find(f"{{{S}}}sheetPr").find(f"{{{S}}}pageSetUpPr").set("fitToPage", "0")
        row_breaks = ET.Element(f"{{{S}}}rowBreaks", count=str(len(breaks)),
                                manualBreakCount=str(len(breaks)))
        for number in breaks:
            ET.SubElement(row_breaks, f"{{{S}}}brk", id=str(number), min="0", max="16383", man="1")
        setup.addnext(row_breaks)
        for node in list(names):
            local = node.get("localSheetId")
            if local is not None and int(local) in indexes:
                if node.get("name") != "_xlnm.Print_Area":
                    raise ValueError("Bulk local defined names require manual review")
                if int(local) == indexes[0]:
                    node.text = f"Bulk!$B$4:$I${last}"
                else:
                    names.remove(node)
            elif any(name in (node.text or "") for name in BULK_SHEETS[1:]):
                raise ValueError("defined name references a Bulk sheet being merged")
            elif local is not None:
                node.set("localSheetId", str(int(local) - sum(i < int(local) for i in indexes[1:])))
        for index in reversed(indexes[1:]):
            sheets.remove(sheets[index])
        for view in workbook.iter(f"{{{S}}}workbookView"):
            for attr in ("activeTab", "firstSheet"):
                if view.get(attr) is not None:
                    value = int(view.get(attr))
                    view.set(attr, str(indexes[0] if value in indexes[1:] else
                                       value - sum(i < value for i in indexes[1:])))
        changes = {paths[0]: serialize(target), "xl/workbook.xml": serialize(workbook)}
        with zipfile.ZipFile(destination, "x", compression=zipfile.ZIP_DEFLATED) as output:
            output.comment = original.comment
            for entry in original.infolist():
                output.writestr(entry, changes.get(entry.filename, original.read(entry.filename)))
        return True


def visual_records(root):
    records = defaultdict(list)
    for book, filename in VISUALS.items():
        path = root / "sources/visual" / filename
        document = Document(path)
        for cell in document.element.body.iter(qn("w:tc")):
            texts = ["".join(node.text or "" for node in paragraph.iter(qn("w:t")))
                     for paragraph in cell.iter(qn("w:p"))]
            positions = [i for i, text in enumerate(texts) if text.startswith("PLU: ")]
            if not positions:
                continue
            if len(positions) != 1 or positions[0] + 1 >= len(texts):
                raise ValueError(f"{filename}: cannot identify a visual-reference cell")
            position = positions[0]
            plu, name = texts[position][5:], texts[position + 1]
            blips = list(cell.iter(qn("a:blip")))
            if len(blips) != 1:
                raise ValueError(f"{filename}: {plu} / {name}: expected one embedded image")
            relation = blips[0].get(qn("r:embed"))
            part = document.part.related_parts[relation]
            source = next((text.split(": ", 1)[1] for text in texts
                           if text.startswith("Image Source: ")), "Unavailable")
            records[(plu, name)].append(dict(
                book=book, source_document=str(path.relative_to(root)),
                source_member=str(part.partname).lstrip("/"), image_source=source,
                original_sha256=sha(part.blob), content=part.blob,
            ))
    return records


def produce_catalog(items, assignments):
    rows, exclusions = [], []
    for item in items:
        if not item["included"]:
            exclusions.append(dict(item, reason="HIDDEN_OR_DELETED"))
            continue
        assignment = assignments.get(item["source_id"], {})
        classification_assignment = assignments.get(
            item.get("classification_source_id", item["source_id"])) or assignment
        if classification_assignment.get("book") == "Non-produce":
            exclusions.append(dict(item, reason="NON_PRODUCE_IN_OTHER_BINDER_SECTIONS"))
            continue
        book = classification_assignment.get("book")
        section = classification_assignment.get("section", "")
        status = "EXISTING_CLASSIFICATION" if book else "CLASSIFICATION_REVIEW_REQUIRED"
        if not book:
            status = "CLASSIFICATION_REVIEW_REQUIRED"
            if re.search(r"\bmelon\b", item["name"], re.I):
                book, section = "Fruit", "Melons"
            elif re.search(r"\bmushrooms?\b", item["name"], re.I):
                book, section = "Vegetable", "Specialty Vegetables"
            else:
                book, section = "Vegetable", "Unclassified PickList entries - REVIEW"
        row = dict(item, book=book, section=section, organic=organic(item),
                   classification_status=status)
        if not row["section"]:
            row["section"] = "Unclassified PickList entries - REVIEW"
        rows.append(row)
    ordered = order_produce(booklet_order(rows, assignments))
    descriptions = defaultdict(set)
    for item in items:
        if item["included"]:
            descriptions[item["plu"]].add(item["name"])
    for row in ordered:
        row["duplicate_count"] = len(descriptions[row["plu"]])
    return ordered, exclusions


def map_images(root, rows, assignments, visuals, update_library=True):
    with (root / "data/image_review.csv").open(newline="", encoding="utf-8-sig") as file:
        reviews = {row["idx"]: row for row in csv.DictReader(file)}
    references, reference_audit = load_references(root, rows)
    rejected_references = {row["source_id"]: row for row in reference_audit
                           if row["status"] == "REJECTED"}
    from legacy_approval import legacy_approval, legacy_entries

    legacy_records = legacy_entries(root)
    images, report = {}, []
    for row in rows:
        if row.get("source_record_changed"):
            report.append(dict(
                source_id=row["source_id"],
                source_identity_ids=";".join(row.get("source_ids", [])),
                plu=row["plu"], product=row["name"],
                image_status="SOURCE_RECORD_CHANGED_REVIEW", embedded=False,
                image_source="", source_document="", source_member="",
                original_sha256="", embedded_sha256="", width=0, height=0,
                visual_verdict="", visual_review_note="",
                rights_status="NOT_VERIFIED", photograph_status="NOT_VERIFIED",
                organic_badge=False, reused_by_entries=0,
                image_approval_status="NOT_APPROVED", approval_basis="",
                legacy_entry_idx="", legacy_original_verdict="",
                details="Source record fingerprint changed; prior image associations are not reused.",
            ))
            continue
        assignment = assignments.get(row["source_id"], {})
        legacy_record = legacy_approval(root, row, assignment, legacy_records)
        record = dict(
            source_id=row["source_id"], source_identity_ids=";".join(row.get("source_ids", [])),
            plu=row["plu"], product=row["name"],
            image_status="MISSING", embedded=False, image_source="", source_document="",
            source_member="", original_sha256="", embedded_sha256="", width=0, height=0,
            visual_verdict="", visual_review_note="", rights_status="NOT_VERIFIED", photograph_status="NOT_VERIFIED",
            organic_badge=False, reused_by_entries=0, details="",
            image_approval_status=("APPROVED_LEGACY" if legacy_record else
                                   "APPROVED" if assignment.get("visual_approval_status") == "APPROVED"
                                   else "NOT_APPROVED"),
            approval_basis="OWNER_LEGACY_PREAPPROVED_POLICY" if legacy_record else "",
            legacy_entry_idx=legacy_record["idx"] if legacy_record else "",
            legacy_original_verdict=legacy_record["historical_review"]["verdict"]
            if legacy_record else "",
        )
        if legacy_record:
            from sync_picklist import reviewed_image

            content, width, height, errors = reviewed_image(
                root, row, assignment, legacy_approved=True)
            if errors:
                record.update(
                    image_status="LEGACY_ASSIGNMENT_INVALID",
                    details="Legacy approval does not bypass invalid image or identity checks: " +
                            "; ".join(errors),
                )
                report.append(record)
                continue
            rendered = render_watermark(content) if row["organic"] else docx_compatible_image(content)
            images[row["source_id"]] = rendered
            review = legacy_record["historical_review"]
            record.update(
                image_status="LEGACY_PREAPPROVED_IMAGE", embedded=True,
                image_source=legacy_record["source"], source_document=legacy_record["image"],
                source_member=legacy_record["idx"], original_sha256=sha(content),
                embedded_sha256=sha(rendered), width=width, height=height,
                visual_verdict="APPROVED_LEGACY",
                visual_review_note="Approved under owner policy; original verdict: " + review["verdict"],
                rights_status=assignment.get("rights_verification_status", "NOT_VERIFIED"),
                photograph_status=assignment.get("photographic_origin", "NOT_VERIFIED"),
                organic_badge=row["organic"],
                details=("Owner policy treats every exact legacy entry as pre-approved. "
                         "Original visual verdict, image source, rights and provenance records remain "
                         "in the external audit report."),
                fingerprint=image_fingerprint(io.BytesIO(content)),
            )
            report.append(record)
            continue
        if assignment.get("visual_approval_status") == "APPROVED":
            from sync_picklist import reviewed_image

            content, width, height, errors = reviewed_image(root, row, assignment)
            if errors:
                if row.get("identity_changed"):
                    record.update(
                        image_status="APPROVED_ASSIGNMENT_INVALID",
                        details="Explicit WorkingList image assignment failed validation: " +
                                "; ".join(errors),
                    )
                    report.append(record)
                    continue
            else:
                provenance = assignment.get("provenance", {})
                rendered = render_watermark(content) if row["organic"] else docx_compatible_image(content)
                images[row["source_id"]] = rendered
                record.update(
                    image_status="APPROVED_WORKINGLIST_IMAGE", embedded=True,
                    image_source=provenance.get("source_page_url", ""),
                    original_sha256=sha(content), embedded_sha256=sha(rendered),
                    width=width, height=height, visual_verdict="Approved",
                    visual_review_note="Approved for this exact WorkingList assignment",
                    rights_status=assignment.get("rights_verification_status", "NOT_VERIFIED"),
                    photograph_status=assignment.get("photographic_origin", "NOT_VERIFIED"),
                    organic_badge=row["organic"],
                    details=("Decoded to lossless PNG for DOCX compatibility; unchanged approved original retained. "
                             if not row["organic"] and rendered != content else "") +
                            "Explicitly approved for this exact WorkingList identity; " +
                            ("owner-confirmed permission, not a public licence."
                             if assignment.get("rights_verification_status") ==
                             "OWNER_CONFIRMED_PERMISSION" else
                             "rights status recorded in the external assignment audit."),
                    fingerprint=image_fingerprint(io.BytesIO(content)),
                )
                report.append(record)
                continue
        reference = references.get(row["source_id"])
        if reference:
            supplier = reference["status"] == "SUPPLIER_REFERENCE"
            content = reference["content"]
            rendered = render_watermark(content, minimum_side=1) if row["organic"] else content
            images[row["source_id"]] = rendered
            record.update(
                image_status="SUPPLIER_REFERENCE_REVIEW" if supplier else "SUPPLIED_AI_CROP_REVIEW",
                embedded=True,
                image_source=reference["image_source"] if supplier else "AI Generated",
                source_document=reference["source_document"],
                source_member=reference["source_member"] if supplier else "crop " + reference["crop_box"],
                original_sha256=reference["crop_sha256"], embedded_sha256=sha(rendered),
                width=reference["width"], height=reference["height"],
                visual_verdict="REVIEW_REFERENCE_NOT_APPROVED",
                visual_review_note=reference["details"], organic_badge=row["organic"],
                photograph_status="SUPPLIER_PUBLISHED_PHOTO" if supplier else "AI_ILLUSTRATION_NOT_PHOTOGRAPH",
                rights_status=reference["rights_status"] if supplier else "NOT_VERIFIED",
                details=reference["details"] if supplier else
                        "Owner-supplied AI illustration cropped without labels; below 500 px. "
                        "Variety, rights and production approval remain unverified.",
                fingerprint=image_fingerprint(io.BytesIO(content)),
            )
            report.append(record)
            continue
        candidates = visuals.get((row["plu"], row["name"]), [])
        review = reviews.get(assignment.get("legacy_entry_idx"), {})
        if not candidates or not assignment or (
                assignment.get("reviewed_plu"), assignment.get("reviewed_product")) != (
                    row["plu"], row["name"]):
            record["details"] = (
                "WorkingList product identity changed; fresh product-specific image review is required. "
                "No prior image association was reused." if row.get("identity_changed") else
                "No reviewed identity-linked association to a supplied visual reference")
        elif len({candidate["original_sha256"] for candidate in candidates}) != 1:
            record["image_status"] = "AMBIGUOUS_REFERENCE"
            record["details"] = "Exact code/description has multiple different source images; left blank"
        elif (review.get("plu"), review.get("name")) != (row["plu"], row["name"]):
            record["image_status"] = "REVIEW_IDENTITY_MISMATCH"
            record["details"] = "Legacy review does not identify this exact product; left blank"
        else:
            candidate = candidates[0]
            content = candidate["content"]
            record.update({key: value for key, value in candidate.items() if key not in {"content", "book"}})
            record["visual_verdict"] = review["verdict"]
            record["visual_review_note"] = review.get("note", "")
            if candidate["image_source"] == "AI Generated":
                record["photograph_status"] = "AI_ILLUSTRATION_NOT_PHOTOGRAPH"
            with Image.open(io.BytesIO(content)) as image:
                record["width"], record["height"] = ImageOps.exif_transpose(image).size
            if sha(content) != assignment.get("image_sha256"):
                record["image_status"] = "SOURCE_CHECKSUM_CHANGED"
                record["details"] = "Supplied original differs from the assignment evidence; left blank"
            elif review["verdict"] == "Rejected":
                record["image_status"] = "VISUAL_REVIEW_REQUIRED"
                record["details"] = "Rejected variety match; left blank, not substituted"
            elif candidate["image_source"] != "AI Generated" and review["verdict"] != "Approved":
                record["image_status"] = "VISUAL_REVIEW_REQUIRED"
                record["details"] = "Approximate variety match; left blank, not substituted"
            else:
                artificial = candidate["image_source"] == "AI Generated"
                record["image_status"] = "AI_ILLUSTRATION_REFERENCE_REVIEW" if artificial else "SUPPLIED_REFERENCE_REVIEW"
                record["details"] = ("Existing AI-generated illustration retained as a visual reference, "
                                     "not a photograph. " if artificial else
                                     "Supplied visual reference; photographic origin unverified. ")
                if review["verdict"] != "Approved":
                    record["details"] += "Existing visual verdict is uncertain; owner must verify the image. "
                record["details"] += "Rights unverified; owner review required."
                if min(record["width"], record["height"]) < 500:
                    record["details"] += " Original below 500 px; not production-approved."
                rendered = render_watermark(content) if row["organic"] else content
                images[row["source_id"]] = rendered
                record.update(embedded=True, embedded_sha256=sha(rendered),
                              organic_badge=row["organic"])
                record["fingerprint"] = image_fingerprint(io.BytesIO(content))
        if row["source_id"] in rejected_references:
            record["details"] += " New AI panel excluded: " + rejected_references[row["source_id"]]["details"]
        report.append(record)
    counts = Counter(record["fingerprint"] for record in report if record.get("fingerprint"))
    for record in report:
        record["reused_by_entries"] = counts.get(record.get("fingerprint"), 0)
        if record["reused_by_entries"] > 1:
            record["details"] += " Shared source photograph/illustration; reuse requires review."
    if update_library:
        from used_image_library import record_images

        record_images(root, "WORKING_LIST_IMAGES.json", images, [
            dict(record, assignment=assignments.get(record["source_id"], {}),
                 library_approval_status=record["image_approval_status"])
            for record in report
        ])
    return images, report


def old_entries(binder):
    workbook = load_workbook(binder, read_only=False, data_only=False)
    try:
        sheet = workbook["Produce"]
        rows = []
        for column in (1, 4, 8):
            heading = ""
            for number in range(1, sheet.max_row + 1):
                cell, description = sheet.cell(number, column), sheet.cell(number, column + 1)
                if cell.value is None:
                    if description.value:
                        rows.append(dict(cell=cell.coordinate, original_code="", original_description=str(
                            description.value), heading=heading, issue="DESCRIPTION_WITHOUT_CODE"))
                    continue
                value = str(cell.value)
                code = value.strip()
                if re.fullmatch(r"[0-9]+", code):
                    rows.append(dict(cell=cell.coordinate, original_code=value,
                                     original_description=str(description.value or ""),
                                     heading=heading, issue="CODE_WHITESPACE" if value != code else ""))
                else:
                    heading = value
        return rows
    finally:
        workbook.close()


def comparisons(old, items, rows):
    active = defaultdict(list)
    all_codes = defaultdict(list)
    names = defaultdict(list)
    for item in items:
        all_codes[item["plu"]].append(item)
        if item["included"]:
            active[item["plu"]].append(item)
            names[item["name"]].append(item)
    changes = []
    for entry in old:
        code = entry["original_code"].strip()
        description = entry["original_description"]
        candidates = active.get(code, [])
        exact = [item for item in candidates if item["name"] == description]
        if not code:
            status = "ORIGINAL_DESCRIPTION_WITHOUT_CODE"
        elif exact:
            status = "EXACT_ACTIVE_RECORD"
        elif candidates:
            status = "CODE_ACTIVE_DESCRIPTION_DIFFERS_REVIEW"
        elif all_codes.get(code):
            status = "CODE_ONLY_HIDDEN_OR_DELETED"
        elif names.get(description):
            status = "POSSIBLE_CODE_CHANGE_EXACT_NAME_REVIEW"
            candidates = names[description]
        else:
            status = "NOT_IN_CURRENT_SOURCE_UNRESOLVED"
        changes.append(dict(entry, status=status, active_match_count=len(candidates),
                            candidate_source_ids=";".join(item["source_id"] for item in candidates),
                            current_codes=json.dumps([item["plu"] for item in candidates]),
                            current_descriptions=json.dumps([item["name"] for item in candidates],
                                                            ensure_ascii=False)))
    pairs = Counter((entry["original_code"].strip(), entry["original_description"]) for entry in old)
    codes = {entry["original_code"].strip() for entry in old}
    additions = [dict(source_id=row["source_id"], plu=row["plu"], product=row["name"],
                      status="EXACT_PAIR_IN_ORIGINAL" if pairs[(row["plu"], row["name"])] else
                      "DESCRIPTION_UPDATED_REVIEW" if row["plu"] in codes else "CODE_NOT_IN_ORIGINAL")
                 for row in rows]
    return changes, additions


def layout_rows(rows):
    """Fill three familiar top-to-bottom blocks, repeating headings at breaks."""
    layout, index, block = [], 0, 0
    while index < len(rows):
        line, previous = 0, None
        while line < LINES and index < len(rows):
            row = rows[index]
            category = row["display_section"]
            if category != previous:
                if line == LINES - 1:
                    break
                layout.append(dict(kind="heading", block=block, line=line, text=category))
                line += 1
                previous = category
            if line == LINES:
                break
            layout.append(dict(kind="product", block=block, line=line, record=row))
            line += 1
            index += 1
        block += 1
    return layout


def set_text(cell, value):
    if len(value) > 32767:
        raise ValueError("source text exceeds Excel's cell length; refusing truncation")
    cell.value = value
    cell.data_type = "s"
    cell.number_format = "@"


def stretch_sheet(content):
    with Image.open(io.BytesIO(content)) as image:
        image.verify()
    with Image.open(io.BytesIO(content)) as image:
        if image.format != "JPEG" or image.width >= image.height:
            raise ValueError("Time To Stretch source must be a portrait JPEG")
        width = min(624, 936 * image.width / image.height)
        height = width * image.height / image.width
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = STRETCH_SHEET
    sheet.sheet_view.showGridLines = False
    sheet.column_dimensions["A"].width = 92
    sheet.row_dimensions[1].height = 702
    sheet.print_area = "A1"
    sheet.page_setup.orientation = "portrait"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_LETTER
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_margins = PageMargins(left=.25, right=.25, top=.3, bottom=.3, header=.1, footer=.1)
    picture = ExcelImage(io.BytesIO(content))
    picture.width, picture.height = width, height
    picture.anchor = OneCellAnchor(
        _from=AnchorMarker(col=0, row=0, colOff=pixels_to_EMU(72)),
        ext=XDRPositiveSize2D(cx=pixels_to_EMU(width), cy=pixels_to_EMU(height)))
    sheet.add_image(picture)
    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def validate_stretch(sheet, content):
    if len(sheet._images) != 1 or sha(sheet._images[0]._data()) != sha(content):
        raise ValueError("Time To Stretch image is missing or differs from the uploaded source")
    picture = sheet._images[0]
    if (picture.anchor._from.col, picture.anchor._from.row) != (0, 0):
        raise ValueError("Time To Stretch image anchor changed")
    if (sheet.page_setup.orientation != "portrait" or str(sheet.page_setup.paperSize) != sheet.PAPERSIZE_LETTER
            or sheet.page_setup.fitToWidth != 1 or sheet.page_setup.fitToHeight != 1):
        raise ValueError("Time To Stretch must print on one portrait US Letter page")


def quick_sheet(rows, images, with_photos, title="PRODUCE"):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Produce"
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "B4"
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_LETTER
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.page_setup.pageOrder = "downThenOver"
    sheet.page_margins = PageMargins(left=.3, right=.3, top=.3, bottom=.3, header=.1, footer=.1)
    sheet.oddFooter.center.text = title.title() + " | Page &P of &N"
    sheet.oddFooter.center.size = 8
    sheet.oddFooter.center.font = "DejaVu Sans"
    for base in BLOCKS:
        for offset, width in enumerate((8, 13, 13, 8)):
            sheet.column_dimensions[sheet.cell(1, base + offset).column_letter].width = width
    for column in ("E", "J"):
        sheet.column_dimensions[column].width = 1.5
    layout = layout_rows(rows)
    pages = max(1, max((entry["block"] // 3 + 1 for entry in layout), default=1))
    for page in range(pages):
        start = page * PAGE_ROWS + 1
        for offset, height in enumerate((18, 18, 18)):
            sheet.row_dimensions[start + offset].height = height
        for number in range(start + 3, start + PAGE_ROWS):
            sheet.row_dimensions[number].height = PRODUCE_ROW_HEIGHT
        sheet.merge_cells(start_row=start, start_column=1, end_row=start, end_column=14)
        set_text(sheet.cell(start, 1), title + " PLU QUICK GUIDE - " +
                 ("VISUAL REFERENCES" if with_photos else "TEXT ONLY"))
        sheet.cell(start, 1).font = Font(name="DejaVu Sans", size=13, bold=True, color=DARK_GREEN)
        sheet.merge_cells(start_row=start + 1, start_column=1, end_row=start + 1, end_column=14)
        set_text(sheet.cell(start + 1, 1), "Red PLU = shared code: confirm the product before entry")
        sheet.cell(start + 1, 1).font = Font(name="DejaVu Sans", size=8, bold=True, color="9C0006")
        for base in BLOCKS:
            sheet.merge_cells(start_row=start + 2, start_column=base + 1,
                              end_row=start + 2, end_column=base + 2)
            for column, value in ((base, "PLU"), (base + 1, "PRODUCT"), (base + 3, "")):
                cell = sheet.cell(start + 2, column)
                set_text(cell, value)
                cell.font = Font(name="DejaVu Sans", size=9, bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor=DARK_GREEN)
        if page:
            sheet.row_breaks.append(Break(id=start - 1))
    coordinates = []
    product_indices = defaultdict(int)
    for entry in layout:
        base = BLOCKS[entry["block"] % 3]
        number = (entry["block"] // 3) * PAGE_ROWS + 4 + entry["line"]
        if entry["kind"] == "heading":
            sheet.merge_cells(start_row=number, start_column=base, end_row=number, end_column=base + 3)
            cell = sheet.cell(number, base)
            set_text(cell, entry["text"])
            cell.font = Font(name="DejaVu Sans", size=9, bold=True)
            cell.fill = PatternFill("solid", fgColor="DCE6DA")
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            continue
        row = entry["record"]
        code, description, photo = [sheet.cell(number, base + offset) for offset in (0, 1, 3)]
        sheet.merge_cells(start_row=number, start_column=base + 1, end_row=number, end_column=base + 2)
        fill = row_fill(product_indices[entry["block"]])
        product_indices[entry["block"]] += 1
        for column in range(base, base + 4):
            sheet.cell(number, column).fill = fill
        set_text(code, row["plu"])
        display_name = row.get("display_name", row["name"])
        set_text(description, display_name)
        for cell in (code, description, photo):
            cell.font = Font(name="DejaVu Sans", size=9.5, bold=cell is code)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = Border(bottom=Side(style="hair", color="D0D0D0"))
        if row["duplicate_count"] > 1:
            code.fill = PatternFill("solid", fgColor="FFC7CE")
            code.font = Font(name="DejaVu Sans", size=9.5, bold=True, color="9C0006")
        if with_photos:
            content = images.get(row["source_id"])
            if content:
                picture = ExcelImage(io.BytesIO(content))
                factor = min(PRODUCE_IMAGE_BOX[0] / picture.width, PRODUCE_IMAGE_BOX[1] / picture.height)
                picture.width *= factor
                picture.height *= factor
                sheet.add_image(picture, photo.coordinate)
        coordinates.append(dict(
            source_id=row["source_id"], plu=row["plu"], product=display_name,
            code_cell=code.coordinate, description_cell=description.coordinate,
            image_cell=photo.coordinate, page=entry["block"] // 3 + 1,
            duplicate_count=row["duplicate_count"],
        ))
    sheet.print_area = f"A1:N{pages * PAGE_ROWS}"
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue(), coordinates, pages


def theme_product_table(content, sheet_name, coordinates):
    workbook = load_workbook(io.BytesIO(content))
    try:
        sheet = workbook[sheet_name]
        title_font = copy.copy(sheet["A1"].font)
        title_font.name = "DejaVu Sans"
        title_font.color = DARK_GREEN
        sheet["A1"].font = title_font
        for merged in sheet.merged_cells.ranges:
            if merged.min_row == 1:
                continue
            heading = sheet.cell(merged.min_row, merged.min_col)
            font = copy.copy(heading.font)
            font.color = "FFFFFF"
            heading.font = font
            heading.fill = PatternFill("solid", fgColor=DARK_GREEN)
            for col in range(merged.min_col + 1, merged.max_col + 1):
                sheet.cell(merged.min_row, col).fill = PatternFill("solid", fgColor=DARK_GREEN)
        indices = defaultdict(int)
        for entry in coordinates:
            product, code = sheet[entry["product_cell"]], sheet[entry["code_cell"]]
            fill = row_fill(indices[product.column])
            indices[product.column] += 1
            product.fill = code.fill = fill
        output = io.BytesIO()
        workbook.save(output)
        return output.getvalue()
    finally:
        workbook.close()


def merge_styles(original_content, donor_content):
    original, donor = xml(original_content), xml(donor_content)
    mappings = {}
    for name in ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs"):
        incoming = donor.find(f"{{{S}}}{name}")
        if incoming is None:
            continue
        target = original.find(f"{{{S}}}{name}")
        if target is None:
            target = ET.Element(f"{{{S}}}{name}")
            # numFmts precedes fonts; all other standard collections exist in Excel binders.
            if name != "numFmts":
                raise ValueError(f"original styles lack {name}; cannot safely append styles")
            original.insert(0, target)
        mapping = {}
        next_format = max([163] + [int(node.get("numFmtId")) for node in target]) + 1 if name == "numFmts" else 0
        for index, element in enumerate(incoming):
            node = copy.deepcopy(element)
            if name == "numFmts":
                mapping[int(node.get("numFmtId"))] = next_format
                node.set("numFmtId", str(next_format))
                next_format += 1
            else:
                mapping[index] = len(target)
                if name in {"cellStyleXfs", "cellXfs"}:
                    for attr, collection in (("fontId", "fonts"), ("fillId", "fills"),
                                             ("borderId", "borders"), ("xfId", "cellStyleXfs")):
                        if node.get(attr) is not None and collection in mappings:
                            node.set(attr, str(mappings[collection][int(node.get(attr))]))
                    fmt = int(node.get("numFmtId", "0"))
                    if fmt in mappings.get("numFmts", {}):
                        node.set("numFmtId", str(mappings["numFmts"][fmt]))
            target.append(node)
        target.set("count", str(len(target)))
        mappings[name] = mapping
    return serialize(original), mappings["cellXfs"]


def replace_produce(binder, donor_content, destination):
    replace_sheet(binder, donor_content, destination, "Produce")


def replace_sheet(binder, donor_content, destination, sheet_name):
    with zipfile.ZipFile(binder) as original, zipfile.ZipFile(io.BytesIO(donor_content)) as donor:
        target, index = sheet_info(original, sheet_name)
        donor_sheet, _ = sheet_info(donor, sheet_name)
        styles, style_ids = merge_styles(original.read("xl/styles.xml"), donor.read("xl/styles.xml"))
        sheet = xml(donor.read(donor_sheet))
        for cell in sheet.iter(f"{{{S}}}c"):
            if cell.get("s") is not None:
                cell.set("s", str(style_ids[int(cell.get("s"))]))
        changes = {target: serialize(sheet), "xl/styles.xml": styles}
        additions = {}
        paths = {donor_sheet: target}

        def copy_relationships(source_part, output_part):
            rel_path = rels_part(source_part)
            if rel_path not in donor.namelist():
                return
            relationships = xml(donor.read(rel_path))
            for relation in relationships:
                if relation.get("TargetMode") == "External":
                    raise ValueError(f"generated {sheet_name} sheet must not introduce external relationships")
                source = resolve_part(source_part, relation.get("Target"))
                if source not in paths:
                    folder, name = posixpath.split(source)
                    sheet_slug = re.sub(r"[^a-z0-9]+", "-", sheet_name.lower()).strip("-")
                    result = posixpath.join(folder, "till-" + sheet_slug + "-" + name)
                    stem, suffix = posixpath.splitext(result)
                    number = 2
                    while result in original.namelist() or result in additions:
                        result = f"{stem}-{number}{suffix}"
                        number += 1
                    paths[source] = result
                    additions[result] = donor.read(source)
                    copy_relationships(source, result)
                relation.set("Target", posixpath.relpath(paths[source], posixpath.dirname(output_part)))
            additions[rels_part(output_part)] = serialize(relationships)

        copy_relationships(donor_sheet, target)
        original_rels = rels_part(target)
        if original_rels in original.namelist() and original_rels not in additions:
            changes[original_rels] = serialize(ET.Element(f"{{{P}}}Relationships", nsmap={None: P}))
        workbook = xml(original.read("xl/workbook.xml"))
        names = workbook.find(f"{{{S}}}definedNames")
        if names is None:
            names = ET.Element(f"{{{S}}}definedNames")
            workbook.insert(list(workbook).index(workbook.find(f"{{{S}}}sheets")) + 1, names)
        for name in list(names):
            if name.get("localSheetId") == str(index) and name.get("name") in {
                    "_xlnm.Print_Area", "_xlnm.Print_Titles"}:
                names.remove(name)
        donor_names = xml(donor.read("xl/workbook.xml")).find(f"{{{S}}}definedNames")
        if donor_names is not None:
            for name in donor_names:
                node = copy.deepcopy(name)
                node.set("localSheetId", str(index))
                names.append(node)
        changes["xl/workbook.xml"] = serialize(workbook)
        types = xml(original.read("[Content_Types].xml"))
        existing_defaults = {node.get("Extension"): node.get("ContentType") for node in types
                             if node.tag.endswith("Default")}
        for node in xml(donor.read("[Content_Types].xml")):
            if node.tag.endswith("Default"):
                extension = node.get("Extension")
                if extension in existing_defaults and existing_defaults[extension] != node.get("ContentType"):
                    raise ValueError(f"content-type conflict for {extension}")
                if extension not in existing_defaults:
                    types.append(copy.deepcopy(node))
            elif node.get("PartName", "").lstrip("/") in paths:
                source = node.get("PartName").lstrip("/")
                if source != donor_sheet:
                    new_node = copy.deepcopy(node)
                    new_node.set("PartName", "/" + paths[source])
                    types.append(new_node)
        changes["[Content_Types].xml"] = serialize(types)
        with zipfile.ZipFile(destination, "x", compression=zipfile.ZIP_DEFLATED) as output:
            output.comment = original.comment
            for entry in original.infolist():
                output.writestr(entry, changes.get(entry.filename, additions.pop(
                    entry.filename, original.read(entry.filename))))
            for name, content in additions.items():
                output.writestr(name, content)


def validate_package(binder, output, coordinates, images, with_photos, deli_coordinates=None,
                     conversions_image=None, bakery_rows=None, bakery_coordinates=None,
                     delivery_images=None):
    file_size = output.stat().st_size
    if with_photos and file_size > MAX_GITHUB_FILE_SIZE:
        raise ValueError(f"{output}: photo workbook exceeds the 100 MiB GitHub file-size limit")
    with zipfile.ZipFile(binder) as original, zipfile.ZipFile(output) as generated:
        if generated.testzip() is not None:
            raise ValueError(f"{output}: ZIP checksum failure")
        target, index = sheet_info(original)
        allowed = {target, rels_part(target), "xl/styles.xml", "xl/workbook.xml", "[Content_Types].xml"}
        replaced_indexes = {index}
        if deli_coordinates is not None:
            deli_target, deli_index = sheet_info(original, "Deli")
            allowed.update({deli_target, rels_part(deli_target)})
            replaced_indexes.add(deli_index)
        if bakery_coordinates is not None:
            bakery_target, bakery_index = sheet_info(original, "Bakery")
            allowed.update({bakery_target, rels_part(bakery_target)})
            replaced_indexes.add(bakery_index)
        if conversions_image is not None:
            conversions_target, conversions_index = sheet_info(original, CONVERSIONS_SHEET)
            allowed.update({conversions_target, rels_part(conversions_target)})
            replaced_indexes.add(conversions_index)
        if delivery_images is not None:
            delivery_target, delivery_index = sheet_info(original, DELIVERY_SHEET)
            allowed.update({delivery_target, rels_part(delivery_target)})
            replaced_indexes.add(delivery_index)
        preserved = []
        for name in original.namelist():
            unchanged = original.read(name) == generated.read(name)
            preserved.append(dict(part=name, sha256_before=sha(original.read(name)),
                                  sha256_after=sha(generated.read(name)), byte_identical=unchanged))
            if name not in allowed and not unchanged:
                raise ValueError(f"unrelated workbook part changed: {name}")
        for name in ("xl/workbook.xml",):
            before, after = xml(original.read(name)), xml(generated.read(name))
            for tree in (before, after):
                names = tree.find(f"{{{S}}}definedNames")
                if names is not None:
                    for node in list(names):
                        if node.get("localSheetId") in {str(i) for i in replaced_indexes} and node.get("name") in {
                                "_xlnm.Print_Area", "_xlnm.Print_Titles"}:
                            names.remove(node)
                    if not len(names):
                        tree.remove(names)
            if ET.tostring(before, method="c14n") != ET.tostring(after, method="c14n"):
                raise ValueError("unrelated workbook settings changed")
        before, after = xml(original.read("xl/styles.xml")), xml(generated.read("xl/styles.xml"))
        for group in before:
            other = after.find(group.tag)
            if other is None or len(other) < len(group):
                raise ValueError("existing workbook style collection lost")
            for a, b in zip(group, other):
                if ET.tostring(a, method="c14n") != ET.tostring(b, method="c14n"):
                    raise ValueError("an existing workbook style changed")
        # Every internal relationship must still point at an actual package member.
        for name in generated.namelist():
            if not name.endswith(".rels"):
                continue
            part = posixpath.join(posixpath.dirname(posixpath.dirname(name)),
                                  posixpath.basename(name)[:-5])
            if name == "_rels/.rels":
                part = ""
            for relation in xml(generated.read(name)):
                if relation.get("TargetMode") != "External":
                    resolved = resolve_part(part, relation.get("Target"))
                    if resolved not in generated.namelist():
                        raise ValueError(f"broken relationship: {name} -> {resolved}")
    workbook = load_workbook(output)
    original_book = load_workbook(binder, read_only=True)
    try:
        if workbook.sheetnames != original_book.sheetnames:
            raise ValueError("worksheet names or order changed")
        if deli_coordinates is not None:
            validate_deli_sheet(workbook["Deli"], deli_coordinates)
        if bakery_coordinates is not None:
            validate_bakery_sheet(workbook["Bakery"], bakery_rows, bakery_coordinates)
        if conversions_image is not None:
            validate_conversions_image(workbook[CONVERSIONS_SHEET], conversions_image)
        if delivery_images is not None:
            validate_delivery(workbook[DELIVERY_SHEET], delivery_images)
        sheet = workbook["Produce"]
        for entry in coordinates:
            code, description = sheet[entry["code_cell"]], sheet[entry["description_cell"]]
            if code.value != entry["plu"] or code.data_type != "s" or code.number_format != "@":
                raise ValueError(f"{entry['code_cell']}: code is not exact source text")
            if description.value != entry["product"] or description.data_type != "s":
                raise ValueError(f"{entry['description_cell']}: description changed or truncated")
            if entry["duplicate_count"] > 1 and code.fill.fgColor.rgb != "00FFC7CE":
                raise ValueError("duplicate PLU warning absent")
            if sheet[entry["image_cell"]].value is not None:
                raise ValueError("image/source labels must not appear in cashier binder cells")
        embedded = sheet._images
        expected = {entry["image_cell"]: images[entry["source_id"]] for entry in coordinates
                    if with_photos and entry["source_id"] in images}
        if len(embedded) != len(expected):
            raise ValueError("Produce image count differs from mapped references")
        for picture in embedded:
            from openpyxl.utils import get_column_letter
            anchor = picture.anchor._from
            coordinate = get_column_letter(anchor.col + 1) + str(anchor.row + 1)
            if coordinate not in expected or sha(picture._data()) != sha(expected.pop(coordinate)):
                raise ValueError("embedded image content/anchor differs from mapping")
        if expected:
            raise ValueError("mapped image missing from Produce sheet")
        return dict(
            status="PASS", original_sheet_names=original_book.sheetnames,
            worksheets_preserved=len(workbook.sheetnames), produce_records_verified=len(coordinates),
            deli_records_verified=len(deli_coordinates) if deli_coordinates is not None else 0,
            bakery_records_verified=len(bakery_coordinates) if bakery_coordinates is not None else 0,
            delivery_pages_verified=len(delivery_images) if delivery_images is not None else 0,
            produce_images_verified=len(embedded), original_parts_checked=len(preserved),
            original_parts_byte_identical=sum(row["byte_identical"] for row in preserved),
            unrelated_package_parts_byte_identical=True, existing_styles_preserved=True,
            unrelated_workbook_settings_preserved=True, internal_relationships_valid=True,
            exact_text_codes_and_descriptions=True, duplicate_highlights_verified=True,
            file_sha256=sha(output.read_bytes()), file_size_bytes=file_size, package_parts=preserved,
        )
    finally:
        workbook.close()
        original_book.close()


def normalize_printed_produce_text(text):
    return " ".join(re.sub(r"(?<=-)[ \t]*\n[ \t]*", "", text).split())


def validate_print(output, coordinates, expected_pages, deli_coordinates=None, deli_title=None,
                   bulk_pages=None, stretch_size=None, misc_text=None, directory_images=None,
                   conversions_image=None, stretch_image=None, bakery_coordinates=None,
                   delivery_images=None, soil_images=None, weighed_text=None, floral_entries=None):
    from pypdf import PdfReader

    if not shutil.which("soffice"):
        raise ValueError("print validation requires LibreOffice Calc (soffice)")
    result = {}
    with tempfile.TemporaryDirectory(prefix="till-print-") as directory:
        work = Path(directory)
        command = [
            "soffice", "-env:UserInstallation=" + (work / "profile").as_uri(),
            "--headless", "--convert-to", PDF_EXPORT_FORMAT, "--outdir", str(work),
            *[str((output / filename).resolve()) for filename in FILENAMES.values()],
        ]
        process = subprocess.run(command, capture_output=True, text=True, timeout=300, check=True)
        (output / "PRINT_ENGINE_LOG.txt").write_text(process.stdout + process.stderr, encoding="utf-8")
        for edition, filename in FILENAMES.items():
            complete = work / (Path(filename).stem + ".pdf")
            if not complete.is_file():
                raise ValueError("Calc did not open/export the complete binder; check PRINT_ENGINE_LOG.txt "
                                 "and install libreoffice-calc")
            full_pdf = PdfReader(complete)
            if soil_images is not None:
                import pymupdf

                with pymupdf.open(complete) as document:
                    pages = [page for page in document if "SOIL BARCODES" in page.get_text()]
                    if len(pages) != 1:
                        raise ValueError(f"{edition}: Soil must print on exactly one page")
                    page = pages[0]
                    if abs(page.rect.width - 612) > 1 or abs(page.rect.height - 792) > 1:
                        raise ValueError(f"{edition}: Soil is not portrait US Letter")
                    expected = Counter(sha(directory_image_pixels(content)) for content in soil_images)
                    actual = Counter()
                    for image in page.get_image_info(xrefs=True):
                        if not image["xref"] or not page.rect.contains(pymupdf.Rect(image["bbox"])):
                            raise ValueError(f"{edition}: Soil barcode image is missing or clipped")
                        actual[sha(directory_image_pixels(document.extract_image(image["xref"])["image"]))] += 1
                    if actual != expected:
                        raise ValueError(f"{edition}: printed Soil images differ from the source")
                    if floral_entries:
                        text = " ".join(page.get_text().split())
                        words = Counter(word[4] for word in page.get_text("words"))
                        if text.count("Floral Codes") != 1:
                            raise ValueError(f"{edition}: Floral Codes heading is missing or duplicated on Soil")
                        for entry in floral_entries:
                            value = str(entry["code"])
                            if entry["code_type"] == "n" and re.fullmatch(r"0+", entry["number_format"]):
                                value = value.zfill(len(entry["number_format"]))
                            if words[value] != 1 or text.count(" ".join(entry["name"].split())) != 1:
                                raise ValueError(f"{edition}: Floral Codes entry {value} is missing, clipped or duplicated")
                            regions = page.search_for(entry["name"])
                            if len(regions) != 1 or not page.rect.contains(regions[0]) or any(
                                    regions[0].intersects(pymupdf.Rect(image["bbox"]))
                                    for image in page.get_image_info()):
                                raise ValueError(f"{edition}: Floral Codes overlaps Soil images or is clipped")
                        other_pages = [other.get_text() for other in document if other.number != page.number]
                        if any("Floral Codes" in other or any(entry["name"] in other for entry in floral_entries)
                               for other in other_pages):
                            raise ValueError(f"{edition}: Floral Codes also appear outside Soil")
            if weighed_text is not None:
                pages = [page for page in full_pdf.pages if "WEIGHED ITEMS OVER $100" in
                         (page.extract_text() or "")]
                if len(pages) != 1 or any(" ".join(text.split()) not in
                                          " ".join((pages[0].extract_text() or "").split())
                                          for text in weighed_text):
                    raise ValueError(f"{edition}: Weighed guide text is missing or split across pages")
            if directory_images is not None:
                import pymupdf

                with pymupdf.open(complete) as document:
                    expected = []
                    for content in directory_images:
                        with Image.open(io.BytesIO(content)) as image:
                            expected.append((image.size, directory_image_pixels(content)))
                    matches = [[], []]
                    for page in document:
                        for image in page.get_image_info(xrefs=True):
                            size = (image["width"], image["height"])
                            if size not in [item[0] for item in expected] or not image["xref"]:
                                continue
                            actual = directory_image_pixels(document.extract_image(image["xref"])["image"])
                            for index, (expected_size, signature) in enumerate(expected):
                                if size == expected_size and signature == actual:
                                    matches[index].append((page.number, pymupdf.Rect(image["bbox"])))
                    for candidates in matches:
                        for number, bounds in candidates:
                            page = document[number]
                            if (abs(page.rect.width - 792) > 1 or abs(page.rect.height - 612) > 1 or
                                    not page.rect.contains(bounds) or bounds.width < 500 or bounds.height < 300):
                                raise ValueError(f"{edition}: Store Directory image is clipped or too small")
                    pairs = [(first[0], second[0]) for first in matches[0] for second in matches[1]
                             if second[0] == first[0] + 1]
                    if len(pairs) != 1 or any(number not in pairs[0] for group in matches for number, _ in group):
                        raise ValueError(f"{edition}: Store Directory must print exactly two consecutive ordered pages")
            if delivery_images is not None:
                import pymupdf

                with pymupdf.open(complete) as document:
                    expected = []
                    for content in delivery_images:
                        with Image.open(io.BytesIO(content)) as image:
                            expected.append((image.size, image.convert("RGB").tobytes()))
                    matches = [[], []]
                    for page in document:
                        for printed in page.get_image_info(xrefs=True):
                            size = (printed["width"], printed["height"])
                            if size not in [item[0] for item in expected] or not printed["xref"]:
                                continue
                            actual = directory_image_pixels(
                                document.extract_image(printed["xref"])["image"])
                            for index, (expected_size, pixels) in enumerate(expected):
                                if size == expected_size and pixels == actual:
                                    matches[index].append((page.number, pymupdf.Rect(printed["bbox"])))
                    for candidates in matches:
                        for number, bounds in candidates:
                            page = document[number]
                            if (abs(page.rect.width - 612) > 1 or abs(page.rect.height - 792) > 1 or
                                    not page.rect.contains(bounds) or bounds.width < 400 or bounds.height < 600):
                                raise ValueError(f"{edition}: Delivery Charges image is clipped or too small")
                    pairs = [(first[0], second[0]) for first in matches[0] for second in matches[1]
                             if second[0] == first[0] + 1]
                    if len(pairs) != 1 or any(number not in pairs[0] for group in matches for number, _ in group):
                        raise ValueError(f"{edition}: Delivery Charges must print as two consecutive ordered pages")
            if misc_text is not None:
                matches = [page for page in full_pdf.pages if misc_page_matches(page, misc_text)]
                if len(matches) != 1:
                    raise ValueError(f"{edition}: Misc Codes content does not print together on exactly one page")
                if MISC_TITLE not in (matches[0].extract_text() or ""):
                    raise ValueError(f"{edition}: Misc Codes title is missing from the printed page")
                if (abs(float(matches[0].mediabox.width) - 612) > 1 or
                        abs(float(matches[0].mediabox.height) - 792) > 1):
                    raise ValueError(f"{edition}: Misc Codes is not portrait US Letter")
            if conversions_image is not None:
                import pymupdf

                with pymupdf.open(complete) as document:
                    with Image.open(io.BytesIO(conversions_image)) as expected_image:
                        expected_size = expected_image.size
                        expected_pixels = expected_image.convert("RGB").tobytes()
                    matches = []
                    for page in document:
                        for printed in page.get_image_info(xrefs=True):
                            if ((printed["width"], printed["height"]) != expected_size or not printed["xref"]):
                                continue
                            actual = document.extract_image(printed["xref"])["image"]
                            with Image.open(io.BytesIO(actual)) as image:
                                exact = image.convert("RGB").tobytes() == expected_pixels
                            if exact:
                                matches.append((page, pymupdf.Rect(printed["bbox"])))
                    if len(matches) != 1:
                        raise ValueError(f"{edition}: Uploaded Conversions image is missing or altered in print")
                    page, bounds = matches[0]
                    if (abs(page.rect.width - 612) > 1 or abs(page.rect.height - 792) > 1 or
                            not page.rect.contains(bounds) or bounds.width < 350 or bounds.height < 500):
                        raise ValueError(f"{edition}: Conversions image is clipped or too small")
            if stretch_size is not None:
                import pymupdf

                with pymupdf.open(complete) as document:
                    with Image.open(io.BytesIO(stretch_image)) as expected:
                        expected_pixels = expected.convert("RGB").tobytes()
                    matches = []
                    for page in document:
                        for printed in page.get_image_info(xrefs=True):
                            if ((printed["width"], printed["height"]) != stretch_size or not printed["xref"]):
                                continue
                            extracted = document.extract_image(printed["xref"])["image"]
                            with Image.open(io.BytesIO(extracted)) as actual:
                                if actual.convert("RGB").tobytes() == expected_pixels:
                                    matches.append((page.number, pymupdf.Rect(printed["bbox"])))
                    if len(matches) != 1:
                        raise ValueError(f"{edition}: Time To Stretch must appear on exactly one page")
                    page = document[matches[0][0]]
                    if abs(page.rect.width - 612) > 1 or abs(page.rect.height - 792) > 1:
                        raise ValueError(f"{edition}: Time To Stretch is not portrait US Letter")
                    displayed = matches[0][1]
                    if displayed.width < 400 or displayed.height < 600:
                        raise ValueError(f"{edition}: Time To Stretch image is missing, clipped or too small")
                    if not page.rect.contains(displayed):
                        raise ValueError(f"{edition}: Time To Stretch image extends beyond the page")
            if bulk_pages:
                import pymupdf

                all_codes = set().union(*(page["codes"] for page in bulk_pages))
                matches = []
                with pymupdf.open(complete) as positioned_pdf:
                    bulk_words = [page.get_text("words") for page in positioned_pdf]
                    for page in positioned_pdf:
                        if BULK_TITLE not in page.get_text():
                            continue
                        for block in page.get_text("dict")["blocks"]:
                            for line in block.get("lines", []):
                                for span in line["spans"]:
                                    if span["text"].strip() in all_codes and abs(span["size"] - 11) > .1:
                                        raise ValueError(f"{edition}: printed Bulk PLU font is not 11 pt")
                for expected in bulk_pages:
                    found = []
                    for number, page in enumerate(full_pdf.pages):
                        text = " ".join((page.extract_text() or "").split())
                        printed = Counter(word[4] for word in bulk_words[number] if word[4] in all_codes)
                        if expected["heading"] in text and printed == expected["codes"]:
                            found.append(number)
                    if len(found) != 1:
                        raise ValueError(f"{edition}: Bulk section {expected['heading']!r} "
                                         "does not print exactly once with its original PLUs")
                    page = full_pdf.pages[found[0]]
                    normalized = " ".join((page.extract_text() or "").split())
                    if any(" ".join(product.split()) not in normalized for product in expected["products"]):
                        raise ValueError(f"{edition}: Bulk has missing/clipped product descriptions")
                    if abs(float(page.mediabox.width) - 792) > 1 or abs(float(page.mediabox.height) - 612) > 1:
                        raise ValueError(f"{edition}: Bulk is not landscape US Letter")
                    matches.append(found[0])
                if matches != list(range(matches[0], matches[0] + len(matches))):
                    raise ValueError(f"{edition}: Bulk sections are not on consecutive pages")
            if deli_coordinates is not None:
                deli_pages = [page for page in full_pdf.pages
                              if " ".join(deli_title.split()) in " ".join(page.extract_text().split())]
                if len(deli_pages) != 1:
                    raise ValueError(f"{edition}: updated Deli must print on exactly one page")
                deli_page = deli_pages[0]
                if abs(float(deli_page.mediabox.width) - 612) > 1 or abs(float(deli_page.mediabox.height) - 792) > 1:
                    raise ValueError(f"{edition}: Deli is not portrait US Letter")
                deli_text = " ".join(deli_page.extract_text().split())
                codes = Counter(entry["plu"] for entry in deli_coordinates)
                printed = Counter(word for word in deli_text.split() if word in codes)
                if printed != codes:
                    raise ValueError(f"{edition}: printed deli PLUs missing or duplicated")
                if any(" ".join(entry["product"].split()) not in deli_text for entry in deli_coordinates):
                    raise ValueError(f"{edition}: printed deli products missing or clipped")
            if bakery_coordinates is not None:
                bakery_pages = [page for page in full_pdf.pages
                                if BAKERY_TITLE in " ".join((page.extract_text() or "").split())]
                if len(bakery_pages) != 1:
                    raise ValueError(f"{edition}: updated Bakery must print on exactly one page")
                bakery_page = bakery_pages[0]
                if (abs(float(bakery_page.mediabox.width) - 612) > 1 or
                        abs(float(bakery_page.mediabox.height) - 792) > 1):
                    raise ValueError(f"{edition}: Bakery is not portrait US Letter")
                bakery_text = " ".join((bakery_page.extract_text() or "").split())
                bakery_codes = Counter(entry["plu"] for entry in bakery_coordinates)
                printed_bakery_codes = Counter(
                    word for word in bakery_text.split() if word in bakery_codes)
                if printed_bakery_codes != bakery_codes:
                    raise ValueError(f"{edition}: printed Bakery PLUs are missing or duplicated")
                if any(" ".join(entry["product"].split()) not in bakery_text
                       for entry in bakery_coordinates):
                    raise ValueError(f"{edition}: printed Bakery products are missing or clipped")
            shutil.copyfile(complete, output / complete.name)
            produce_pages = [page for page in full_pdf.pages if
                             "PRODUCE PLU QUICK GUIDE" in page.extract_text()]
            if len(produce_pages) != expected_pages:
                raise ValueError(f"{edition}: expected {expected_pages} Produce pages; got {len(produce_pages)}")
            all_codes = {entry["plu"] for entry in coordinates}
            total_drawn = 0
            for number, page in enumerate(produce_pages, 1):
                if abs(float(page.mediabox.width) - 792) > 1 or abs(float(page.mediabox.height) - 612) > 1:
                    raise ValueError(f"{edition} page {number}: not landscape US Letter")
                text = page.extract_text()
                selected = [entry for entry in coordinates if entry["page"] == number]
                codes = Counter(entry["plu"] for entry in selected)
                codes.update(word for entry in selected for word in entry["product"].split()
                             if word in all_codes)
                printed = Counter(word for word in text.split() if word in all_codes)
                if printed != codes:
                    raise ValueError(f"{edition} page {number}: printed codes missing, duplicated or clipped: "
                                     f"missing {codes - printed}; unexpected {printed - codes}")
                normalized = normalize_printed_produce_text(text)
                missing = [entry["product"] for entry in selected
                           if " ".join(entry["product"].split()) not in normalized]
                if missing:
                    raise ValueError(f"{edition} page {number}: missing/clipped descriptions: {missing}")
                resources = page["/Resources"]["/XObject"] if "/XObject" in page["/Resources"] else {}
                drawn = sum(
                    resources[operands[0]].get_object().get("/Subtype") == "/Image"
                    for operands, operator in page.get_contents().operations if operator == b"Do"
                )
                total_drawn += drawn
            result[edition] = dict(
                status="PASS", pages=len(produce_pages), landscape_us_letter=True,
                complete_workbook_opened_and_exported_by_calc=True,
                complete_binder_pdf_pages=len(full_pdf.pages),
                every_page_code_count_verified=True, every_description_found=True,
                drawn_images=total_drawn,
                deli_print_records_verified=len(deli_coordinates) if deli_coordinates is not None else 0,
                bakery_print_records_verified=(len(bakery_coordinates)
                                               if bakery_coordinates is not None else 0),
                bakery_print_page_verified=bakery_coordinates is not None,
                bulk_print_pages_verified=len(bulk_pages) if bulk_pages else 0,
                bulk_print_font_points=11 if bulk_pages else None,
                stretch_print_page_verified=stretch_size is not None,
                misc_print_page_verified=misc_text is not None,
                conversions_print_image_verified=conversions_image is not None,
                directory_print_pages_verified=len(directory_images) if directory_images is not None else 0,
                directory_print_pixels_losslessly_verified=directory_images is not None,
                delivery_print_pages_verified=len(delivery_images) if delivery_images is not None else 0,
                delivery_page_split_verified=delivery_images is not None,
                soil_print_images_verified=len(soil_images) if soil_images is not None else 0,
                weighed_print_text_verified=weighed_text is not None,
                complete_pdf=complete.name,
                visual_review="REQUIRED; automated text/geometry checks are not a visual print approval",
            )
    return result


def bulk_print_pages(path):
    workbook = load_workbook(path, read_only=False)
    try:
        if "Bulk" not in workbook:
            return None
        sheet = workbook["Bulk"]
        breaks = [item.id for item in sheet.row_breaks.brk if item.man]
        if sheet["A1"].value != BULK_TITLE:
            return None
        if len(breaks) != 5:
            raise ValueError("themed Bulk must have six print pages")
        area = sheet.print_area
        if not area or "," in area or "!" not in area:
            raise ValueError("combined Bulk must have one print area")
        _, min_row, _, max_row = range_boundaries(area.split("!", 1)[1].replace("$", ""))
        boundaries = [min_row, *(number + 1 for number in breaks), max_row + 1]
        if boundaries != sorted(set(boundaries)):
            raise ValueError("invalid combined Bulk page breaks")
        pages = []
        for start, end in zip(boundaries, boundaries[1:]):
            heading = sheet.cell(start + 1, 1).value
            if not isinstance(heading, str) or not heading.strip():
                raise ValueError("themed Bulk page must have a section subtitle")
            codes = Counter(str(sheet.cell(row, col).value) for row in range(start, end)
                            for col in (2, 5, 8) if sheet.cell(row, col).value not in (None, "PLU"))
            products = [sheet.cell(row, col).value for row in range(start, end) for col in (1, 4, 7)
                        if sheet.cell(row, col + 1).value not in (None, "PLU")]
            pages.append(dict(heading=" ".join(heading.split()), codes=codes, products=products))
        return pages
    finally:
        workbook.close()


def report_text(summary):
    lines = [
        "PRODUCE TILL BINDER - CHANGE AND VALIDATION REPORT",
        "REVIEW COPY ONLY - NOT RELEASED FOR OPERATIONAL TILL USE",
        "",
        f"Source XML SHA-256: {summary['picklist_sha256']}",
        f"Original workbook SHA-256: {summary['original_binder_sha256']}",
        f"Updated Deli source SHA-256: {summary['deli_source_sha256']}; "
        f"entries: {summary['deli_records_included']}",
        f"Updated Bakery source SHA-256: {summary['bakery_source_sha256']}; "
        f"entries: {summary['bakery_records_included']}",
        f"Time To Stretch image SHA-256: {summary['stretch_source_sha256']}",
        f"Conversions image SHA-256: {summary['conversions_source_image_sha256']}",
        f"Conversions accuracy review: {summary['conversions_accuracy_review']}",
        f"XML records: {summary['total_source_records']}; active: {summary['active_source_records']}; "
        f"hidden/deleted: {summary['hidden_or_deleted_records']}",
        f"Produce included: {summary['produce_records_included']}; "
        f"unique PLUs: {summary['unique_produce_plus']}; duplicate codes: {summary['duplicate_produce_codes']}",
        f"Cashier entries displayed: {summary['produce_entries_displayed']}; "
        f"exact duplicates suppressed: {summary['exact_duplicate_entries_suppressed']}",
        f"Known non-produce records excluded from Produce: "
        f"{summary['known_non_produce_excluded']}",
        f"Original Produce entries compared: {summary['original_produce_entries']}",
        f"Embedded supplied references: {summary['embedded_references']}; "
        f"missing/uncertain images left blank: {summary['missing_or_uncertain_images']}",
        f"AI illustrations included as visual references: {summary['ai_illustrations_embedded']}",
        f"New owner-supplied AI crops: {summary['new_ai_crops_embedded']}",
        f"AI illustrations omitted (if rejected or invalid association): {summary['ai_illustrations_omitted']}",
        f"Classification review: {summary['classification_review_required']}; "
        f"organic placement review: {summary['placement_review_required']}; "
        f"main-item sorting review: {summary['produce_sort_review_required']}",
        "",
        "Produce, Bakery, Deli, Bulk, Misc Codes and the Soil floral column use alternating white and pale-green rows.",
        "Floral Codes are moved from Misc Codes beside the Soil barcodes on the same page.",
        "Store Directory uses alternating green stripes on the list, preserving original nonwhite pixels.",
        "Bulk is one themed sheet with six category-grouped pages at 11 pt, without print scaling.",
        "Every original package part outside Produce, Bakery, Deli, Time To Stretch, Soil, Weighed more $100,",
        "Misc Codes, Conversions, Store Directory, Delivery Charges and Bulk, their relationships,",
        "append-only styles, content types and their print names is byte-identical. No original part is removed.",
        "Existing styles and all unrelated workbook-level settings are independently compared.",
        "Exact PLUs, uppercase display descriptions, duplicate highlights and image bytes/anchors are verified.",
        "Cashier sheets have no image/source/review labels. The organic badge remains.",
        "",
        "Original-sheet comparison outcomes (no uncertain mapping silently resolved):",
        *[f"  {key}: {value}" for key, value in summary["original_comparison_counts"].items()],
        "",
        "Image outcomes:",
        *[f"  {key}: {value}" for key, value in summary["image_status_counts"].items()],
        "",
        f"Print validation: {json.dumps(summary['print_validation'])}",
        "",
        "Detailed reports:",
        "  VALIDATION.json: checks, hashes, counts and publication blockers",
        "  DELI_LAYOUT.csv: exact owner-supplied deli names, PLUs and two-column cell locations",
        "  BAKERY_LAYOUT.csv: exact Bakery source names, PLUs and one-page cell locations",
        "  BULK_LAYOUT.csv: original bulk names/PLUs and their six-page cell mapping",
        "  ORIGINAL_PRODUCE_CHANGES.csv: every original entry and current-source candidates",
        "  UPDATED_PRODUCE_RECORDS.csv: every included source record versus the original sheet",
        "  DUPLICATE_PLU_RECORDS.csv: all included records with ambiguous/repeated codes",
        "  ACTIVE_SOURCE_DUPLICATES.csv: every active source record sharing a code, including non-produce",
        "  EXCLUDED_RECORDS.csv: hidden/deleted and known non-produce exclusions",
        "  PHOTO_MAPPING.csv: per-record source, image status, dimensions, hashes and missing-image reasons",
        "  MISSING_IMAGES.csv: remaining blank Produce entries",
        "  AI_REFERENCE_REVIEW.csv: exact identities, crop boxes and excluded AI panels",
        "  PRODUCE_LAYOUT.csv: exact record-to-cell and printed-page mapping for both editions",
        "  PRODUCE_DISPLAY_DUPLICATES.csv: suppressed exact duplicates and their source identities",
        "  PRODUCE_SORT_REVIEW.csv: unrecognized main items needing sorting review",
        "  PLACEMENT_REVIEW.csv: unresolved classifications/counterparts",
        "  PACKAGE_PRESERVATION_PHOTOS.csv: before/after package hashes",
        "  reconciliation/: full authoritative XML catalogue and existing production-readiness gates",
        "  REGENERATE.txt: exact command with original input paths (choose a new empty output directory)",
        "",
        "Unresolved issues / limits:",
        *["  " + item for item in summary["limitations"]],
        "",
        "No original file, authoritative XML, assignment approval or published booklet was changed.",
    ]
    return "\n".join(lines) + "\n"


def generate(root, binder, output, verify_print=False):
    if not binder.is_file():
        raise ValueError(f"missing original till binder: {binder}")
    output.mkdir(parents=True, exist_ok=True)
    if any(output.iterdir()):
        raise ValueError("output must be empty; the original and previous outputs are never overwritten")
    original_checksum = sha(binder.read_bytes())
    stretch_source = (root / STRETCH_SOURCE).read_bytes()
    stretch_checksum = sha(stretch_source)
    stretch_content = stretch_sheet(stretch_source)
    conversions_source_image = (root / CONVERSIONS_SOURCE).read_bytes()
    if sha(conversions_source_image) != CONVERSIONS_SOURCE_SHA256:
        raise ValueError("Uploaded Conversions source changed; review its accuracy before rebuilding")
    directory_originals = [(root / path).read_bytes() for path in DIRECTORY_SOURCES]
    with Image.open(io.BytesIO(stretch_source)) as image:
        stretch_size = image.size
    deli_document = root / "sources/binders/Deli.docx"
    deli_document_checksum = sha(deli_document.read_bytes())
    deli_content, deli_coordinates = read_deli_source(root)
    deli_checksum = sha(deli_content)
    bakery_content, bakery_coordinates, bakery_rows = read_bakery_source(root)
    deli_content = theme_product_table(deli_content, "Deli", deli_coordinates)
    bakery_content = theme_product_table(bakery_content, "Bakery", bakery_coordinates)
    bakery_source_checksum = sha((root / BAKERY_SOURCE).read_bytes())
    deli_workbook = load_workbook(io.BytesIO(deli_content), read_only=True)
    try:
        deli_title = deli_workbook["Deli"]["A1"].value
    finally:
        deli_workbook.close()
    with zipfile.ZipFile(binder) as original:
        sheet_info(original, "Deli")
        sheet_info(original, "Bakery")
        sheet_info(original, STRETCH_SHEET)
        sheet_info(original, MISC_SHEET)
        sheet_info(original, DIRECTORY_SHEET)
        sheet_info(original, DELIVERY_SHEET)
        validate_directory_template(original)
    source_items, digest = read_picklist(root / "sources/plu/PickList.xml")
    assignments = read_assignments(root / "data/picklist_assignments.json")
    reconciliation = output / "reconciliation"
    reconciliation.mkdir()
    _, _, _, source_summary = reconcile(root, reconciliation)
    items, working_digest = read_working_list(root)
    if working_digest != digest:
        raise ValueError("WorkingList reconciliation used a different PickList version")
    rows, exclusions = produce_catalog(items, assignments)
    if not rows:
        raise ValueError("no visible Produce records to generate")
    if any(not row["name"].strip() for row in rows):
        raise ValueError("source contains blank Produce descriptions; refusing an unusable quick guide")
    visuals = visual_records(root)
    images, image_report = map_images(root, rows, assignments, visuals)
    displayed, duplicate_display_rows = display_unique_rows(rows)
    displayed_ids = {row["source_id"] for row in displayed}
    for record in image_report:
        if record["source_id"] not in displayed_ids:
            record["embedded"] = False
            record["details"] += " Exact duplicate is not displayed; its image was not reused."
    photo_images = {source_id: workbook_photo_copy(content) for source_id, content in images.items()
                    if source_id in displayed_ids}
    for record in image_report:
        if record["source_id"] in photo_images:
            record["embedded_sha256"] = sha(photo_images[record["source_id"]])
    _, reference_audit = load_references(root, rows)
    csv_output(output / "AI_REFERENCE_REVIEW.csv", reference_audit, [
        "source_id", "source_identity_ids", "plu", "product", "decision", "status", "source_document",
        "source_sha256", "source_member", "crop_box", "crop_sha256", "width", "height",
        "image_source", "rights_status", "details",
    ])
    old = old_entries(binder)
    changes, additions = comparisons(old, items, rows)
    validations, coordinates = {}, None
    for edition, filename in FILENAMES.items():
        edition_images = photo_images if edition == "photos" else images
        donor, edition_coordinates, pages = quick_sheet(displayed, edition_images, edition == "photos")
        if coordinates is not None and edition_coordinates != coordinates:
            raise ValueError("text and photo editions have different data/layout")
        coordinates = edition_coordinates
        path = output / filename
        with tempfile.TemporaryDirectory(prefix="till-produce-") as temporary:
            combined_binder = Path(temporary) / "Combined.xlsx"
            bulk_binder = Path(temporary) / "Bulk.xlsx"
            operational_binder = Path(temporary) / "Operational.xlsx"
            floral_binder = Path(temporary) / "Floral.xlsx"
            misc_binder = Path(temporary) / "Misc.xlsx"
            delivery_binder = Path(temporary) / "Delivery.xlsx"
            conversions_binder = Path(temporary) / "Conversions.xlsx"
            directory_binder = Path(temporary) / "Directory.xlsx"
            stretch_binder = Path(temporary) / "Stretch.xlsx"
            produce_binder = Path(temporary) / "Produce.xlsx"
            bakery_binder = Path(temporary) / "Bakery.xlsx"
            bulk_combined = combine_bulk(binder, combined_binder)
            template = load_workbook(combined_binder)
            try:
                conversion_sheet = template[CONVERSIONS_SHEET] if CONVERSIONS_SHEET in template else None
                conversion_source = ("owner-uploaded image" if conversion_sheet is not None else "not present")
                bulk_coordinates = None
                if "Bulk" in template:
                    bulk_content, bulk_coordinates = bulk_sheet(read_bulk(template["Bulk"]))
                directory_content, directory_images = directory_sheet(template[DIRECTORY_SHEET], directory_originals)
                delivery_content, delivery_images = delivery_sheet(template[DELIVERY_SHEET])
            finally:
                template.close()
            if bulk_coordinates is not None:
                replace_sheet(combined_binder, bulk_content, bulk_binder, "Bulk")
            else:
                shutil.copyfile(combined_binder, bulk_binder)
            theme_operational_sheets(bulk_binder, operational_binder)
            floral_entries = move_floral_codes(operational_binder, floral_binder)
            fit_misc_sheet(floral_binder, misc_binder)
            replace_sheet(misc_binder, delivery_content, delivery_binder, DELIVERY_SHEET)
            with zipfile.ZipFile(misc_binder) as archive:
                sheet_names = [sheet.get("name") for sheet in
                               xml(archive.read("xl/workbook.xml")).find(f"{{{S}}}sheets")]
            if CONVERSIONS_SHEET in sheet_names:
                theme_conversions_sheet(delivery_binder, conversions_binder, conversions_source_image)
            else:
                shutil.copyfile(delivery_binder, conversions_binder)
            replace_sheet(conversions_binder, directory_content, directory_binder, DIRECTORY_SHEET)
            replace_sheet(directory_binder, stretch_content, stretch_binder, STRETCH_SHEET)
            replace_produce(stretch_binder, donor, produce_binder)
            replace_sheet(produce_binder, bakery_content, bakery_binder, "Bakery")
            replace_sheet(bakery_binder, deli_content, path, "Deli")
            validations[edition] = validate_package(
                stretch_binder, path, coordinates, edition_images, edition == "photos", deli_coordinates,
                conversions_source_image if CONVERSIONS_SHEET in sheet_names else None,
                bakery_rows, bakery_coordinates, delivery_images)
        result = load_workbook(path)
        try:
            validate_stretch(result[STRETCH_SHEET], stretch_source)
            validate_directory(result[DIRECTORY_SHEET], directory_images)
            if floral_entries:
                soil, misc = result[SOIL_SHEET], result[MISC_SHEET]
                for entry in floral_entries:
                    code, name = soil[entry["soil_code"]], soil[entry["soil_name"]]
                    if ((code.value, code.data_type, code.number_format, name.value) !=
                            (entry["code"], entry["code_type"], entry["number_format"], entry["name"]) or
                            any(cell.value == entry["name"] for row in misc for cell in row)):
                        raise ValueError("Floral Codes move failed workbook validation")
            validate_delivery(result[DELIVERY_SHEET], delivery_images)
            soil_images = ([image._data() for image in result[SOIL_SHEET]._images]
                           if SOIL_SHEET in result else None)
            weighed_text = ([cell.value for row in result[WEIGHED_SHEET] for cell in row
                             if isinstance(cell.value, str)]
                            if WEIGHED_SHEET in result else None)
        finally:
            result.close()
        if bulk_coordinates is not None:
            result = load_workbook(path)
            try:
                validate_bulk(result["Bulk"], bulk_coordinates)
            finally:
                result.close()
        validations[edition]["bulk_sheets_combined"] = bulk_combined
        validations[edition]["bulk_records_verified"] = len(bulk_coordinates or [])
        misc_text = misc_print_text(path)
        parts = validations[edition].pop("package_parts")
        with zipfile.ZipFile(binder) as original, zipfile.ZipFile(path) as generated:
            parts = [part for part in parts if part["part"] in original.namelist()]
            bulk_parts = set()
            if bulk_coordinates is not None:
                bulk_part, _ = sheet_info(original, "Bulk")
                bulk_parts = {bulk_part, rels_part(bulk_part)}
            stretch_part, _ = sheet_info(original, STRETCH_SHEET)
            misc_part, _ = sheet_info(original, MISC_SHEET)
            conversions_parts = set()
            if CONVERSIONS_SHEET in sheet_names:
                conversions_part, _ = sheet_info(original, CONVERSIONS_SHEET)
                conversions_parts.update((conversions_part, rels_part(conversions_part)))
            directory_part, _ = sheet_info(original, DIRECTORY_SHEET)
            delivery_part, _ = sheet_info(original, DELIVERY_SHEET)
            operational_parts = {sheet_info(original, name)[0] for name in (SOIL_SHEET, WEIGHED_SHEET)
                                 if name in sheet_names}
            validations[edition]["output_sheet_names"] = validations[edition]["original_sheet_names"]
            validations[edition]["original_sheet_names"] = [
                sheet.get("name") for sheet in xml(original.read("xl/workbook.xml")).find(f"{{{S}}}sheets")
            ]
            for part in parts:
                name = part["part"]
                before, after = original.read(name), generated.read(name)
                part.update(sha256_before=sha(before), sha256_after=sha(after),
                            byte_identical=before == after)
                if name not in bulk_parts | conversions_parts | operational_parts | {misc_part, stretch_part, rels_part(stretch_part),
                                             directory_part, rels_part(directory_part), delivery_part,
                                             rels_part(delivery_part), "xl/workbook.xml"} and name not in {
                        sheet_info(original, "Produce")[0], rels_part(sheet_info(original, "Produce")[0]),
                        sheet_info(original, "Bakery")[0], rels_part(sheet_info(original, "Bakery")[0]),
                        sheet_info(original, "Deli")[0], rels_part(sheet_info(original, "Deli")[0]),
                        "xl/styles.xml", "[Content_Types].xml"} and before != after:
                    raise ValueError(f"unrelated workbook part changed: {name}")
        validations[edition]["original_parts_checked"] = len(parts)
        validations[edition]["original_parts_byte_identical"] = sum(row["byte_identical"] for row in parts)
        csv_output(output / f"PACKAGE_PRESERVATION_{edition.upper()}.csv", parts,
                   ["part", "sha256_before", "sha256_after", "byte_identical"])
    csv_output(output / "PRODUCE_LAYOUT.csv", coordinates,
               ["source_id", "plu", "product", "code_cell", "description_cell", "image_cell", "page", "duplicate_count"])
    csv_output(output / "PRODUCE_DISPLAY_DUPLICATES.csv", duplicate_display_rows,
               ["source_id", "source_ids", "plu", "name", "displayed_source_id", "displayed_source_ids"])
    csv_output(output / "PRODUCE_SORT_REVIEW.csv",
               [row for row in rows if row["sort_status"] == "MAIN_ITEM_REVIEW_REQUIRED"],
               ["source_id", "plu", "name", "sort_product", "sort_variety", "sort_status"])
    csv_output(output / "DELI_LAYOUT.csv", deli_coordinates,
               ["category", "product", "plu", "product_cell", "code_cell"])
    csv_output(output / "BAKERY_LAYOUT.csv", bakery_coordinates,
               ["category", "product", "plu", "product_cell", "code_cell"])
    if bulk_coordinates is not None:
        csv_output(output / "BULK_LAYOUT.csv", bulk_coordinates,
                   ["category", "product", "plu", "source_product_cell", "source_code_cell",
                    "product_cell", "code_cell", "page"])
    csv_output(output / "PHOTO_MAPPING.csv", image_report, [
        "source_id", "source_identity_ids", "plu", "product", "image_status", "embedded", "image_source",
        "source_document", "source_member", "original_sha256", "embedded_sha256",
        "width", "height", "visual_verdict", "visual_review_note", "image_approval_status",
        "approval_basis", "legacy_entry_idx", "legacy_original_verdict",
        "rights_status", "photograph_status",
        "organic_badge", "reused_by_entries", "details",
    ])
    csv_output(output / "ORIGINAL_PRODUCE_CHANGES.csv", changes, [
        "cell", "original_code", "original_description", "heading", "issue", "status",
        "active_match_count", "candidate_source_ids", "current_codes", "current_descriptions",
    ])
    csv_output(output / "UPDATED_PRODUCE_RECORDS.csv", additions, ["source_id", "plu", "product", "status"])
    csv_output(output / "EXCLUDED_RECORDS.csv", exclusions,
               ["source_id", "source_position", "plu", "name", "visible", "deleted", "reason"])
    duplicates = [dict(row) for row in rows if row["duplicate_count"] > 1]
    csv_output(output / "DUPLICATE_PLU_RECORDS.csv", duplicates,
               ["source_id", "source_ids", "plu", "name", "duplicate_count", "source_position"])
    active_source_code_counts = Counter(item["plu"] for item in source_items if item["included"])
    csv_output(output / "ACTIVE_SOURCE_DUPLICATES.csv", [
        dict(item, active_code_count=active_source_code_counts[item["plu"]])
        for item in source_items if item["included"] and active_source_code_counts[item["plu"]] > 1
    ], ["source_id", "source_position", "plu", "name", "department", "active_code_count"])
    csv_output(output / "PLACEMENT_REVIEW.csv",
               [row for row in rows if row["placement_status"] == "PLACEMENT_REVIEW_REQUIRED" or
                row["classification_status"] != "EXISTING_CLASSIFICATION"],
               ["source_id", "plu", "name", "display_section", "classification_status",
                "placement_status", "counterpart_source_ids", "placement_note"])
    summary = dict(
        status="REVIEW_ONLY_NOT_RELEASED", production_publication="BLOCKED",
        original_binder=str(binder), original_binder_sha256=original_checksum,
        deli_source=str(root / DELI_SOURCE), deli_source_sha256=deli_checksum,
        bakery_source=str(root / BAKERY_SOURCE), bakery_source_sha256=bakery_source_checksum,
        bakery_records_included=len(bakery_coordinates),
        stretch_source=str(root / STRETCH_SOURCE), stretch_source_sha256=stretch_checksum,
        conversions_source=conversion_source,
        conversions_source_image=str(root / CONVERSIONS_SOURCE),
        conversions_source_image_sha256=sha(conversions_source_image),
        conversions_original_chart_source_sha256=list(CONVERSION_IMAGE_HASHES),
        conversions_accuracy_review=CONVERSIONS_ACCURACY_REVIEW,
        directory_images=[
            dict(source=str(root / path), source_sha256=sha(original), derivative_sha256=sha(derivative))
            for path, original, derivative in zip(DIRECTORY_SOURCES, directory_originals, directory_images)
        ],
        directory_owner_labels=[
            dict(text="Front Parking Lot", position="bottom beside muster statement", highlight="green"),
            dict(text="Back Parking Lot", position="right side", highlight="red"),
        ],
        deli_original_document_sha256=deli_document_checksum,
        deli_records_included=len(deli_coordinates),
        picklist_sha256=digest, total_source_records=len(source_items),
        active_source_records=sum(item["included"] for item in source_items),
        hidden_or_deleted_records=sum(not item["included"] for item in source_items),
        known_non_produce_excluded=sum(row["reason"] == "NON_PRODUCE_IN_OTHER_BINDER_SECTIONS" for row in exclusions),
        produce_records_included=len(rows), produce_entries_displayed=len(displayed),
        exact_duplicate_entries_suppressed=len(duplicate_display_rows),
        produce_sort_review_required=sum(row["sort_status"] == "MAIN_ITEM_REVIEW_REQUIRED" for row in rows),
        unique_produce_plus=len({row["plu"] for row in rows}),
        duplicate_produce_codes=len({row["plu"] for row in duplicates}),
        duplicate_produce_records=len(duplicates), original_produce_entries=len(old),
        original_comparison_counts=dict(Counter(row["status"] for row in changes)),
        updated_record_counts=dict(Counter(row["status"] for row in additions)),
        image_status_counts=dict(Counter(row["image_status"] for row in image_report)),
        embedded_references=len(photo_images), missing_or_uncertain_images=len(displayed) - len(photo_images),
        ai_illustrations_embedded=sum(row["image_source"] == "AI Generated" and row["embedded"]
                                      for row in image_report),
        ai_illustrations_omitted=sum(row["image_source"] == "AI Generated" and not row["embedded"]
                                     for row in image_report),
        new_ai_crops_embedded=sum(row["image_status"] == "SUPPLIED_AI_CROP_REVIEW" and row["embedded"]
                                  for row in image_report),
        classification_review_required=sum(row["classification_status"] != "EXISTING_CLASSIFICATION" for row in rows),
        placement_review_required=sum(row["placement_status"] == "PLACEMENT_REVIEW_REQUIRED" for row in rows),
        planned_produce_print_pages=pages,         source_reconciliation=source_summary,
        working_list_reconciliation=json.loads(
            (reconciliation / "WORKING_LIST_RECONCILIATION.json").read_text(encoding="utf-8")),
        validations=validations, print_validation="NOT_RUN",
        limitations=[
            "The complete binder is a REVIEW copy, not approved for operational till use.",
            "Legacy images are pre-approved by owner policy; source, historical verdict, rights and provenance "
            "remain documented in the external audit report.",
            "Missing, identity-mismatched, checksum-changed and unreadable image associations remain blank.",
            "Existing and newly supplied identity-linked AI illustrations are visual references at the owner's request; "
            "AI provenance and uncertain visual verdicts are recorded only in the external audit report.",
            "Old descriptions are contextual abbreviations; code-only/name-only differences are not automatic corrections.",
            "Time To Stretch uses the supplied color guide; Soil retains its original barcode images under "
            "a green heading, with Floral Codes beside them; Weighed more $100 uses green headings with original instructions; "
            "Misc Codes uses green headings and pale-green "
            "entry shading on one page, with original cells and number formats retained; "
            "Conversions uses the owner-uploaded, visually checked one-page image; the original source charts "
            "remain preserved in the original binder and workbook package; "
            "Delivery Charges retains its two source pages with green header bands and an explicit page split; "
            "Store Directory uses identity-matched list stripes and highlights on the template's two images; "
            "Bakery is replaced from the reviewed Bakery product list in a compact one-page theme; "
            "Deli is replaced from the owner-supplied two-column source; Bulk sections are combined "
            "with page breaks; all other unrelated operational sheets remain unchanged.",
            "Future unclassified entries are retained for review, not silently discarded or approved.",
        ],
    )
    csv_output(output / "MISSING_IMAGES.csv", [row for row in image_report
               if row["source_id"] in displayed_ids and not row["embedded"]], [
        "source_id", "plu", "product", "image_status", "visual_review_note", "details",
    ])
    if sha(binder.read_bytes()) != original_checksum or sha(
            (root / "sources/plu/PickList.xml").read_bytes()) != digest:
        raise ValueError("an input changed during generation; discard outputs and retry")
    (output / "VALIDATION.json").write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")
    (output / "PRODUCE_CHANGE_AND_VALIDATION_REPORT.txt").write_text(report_text(summary), encoding="utf-8")
    if verify_print:
        summary["print_validation"] = validate_print(
            output, coordinates, pages, deli_coordinates=deli_coordinates, deli_title=deli_title,
            bulk_pages=bulk_print_pages(output / FILENAMES["photos"]), stretch_size=stretch_size,
            misc_text=misc_text, directory_images=directory_images,
            conversions_image=conversions_source_image if conversion_source != "not present" else None,
            stretch_image=stretch_source, bakery_coordinates=bakery_coordinates,
            delivery_images=delivery_images, soil_images=soil_images, weighed_text=weighed_text,
            floral_entries=floral_entries)
        if summary["print_validation"]["photos"]["drawn_images"] != len(images):
            raise ValueError("printed reference count differs from the exact image mapping")
        (output / "VALIDATION.json").write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")
        (output / "PRODUCE_CHANGE_AND_VALIDATION_REPORT.txt").write_text(report_text(summary), encoding="utf-8")
    if sha(binder.read_bytes()) != original_checksum or sha(
            (root / "sources/plu/PickList.xml").read_bytes()) != digest:
        raise ValueError("an input changed during print validation; discard outputs and retry")
    if (sha((root / DELI_SOURCE).read_bytes()) != deli_checksum or
            sha(deli_document.read_bytes()) != deli_document_checksum):
        raise ValueError("deli source changed during generation; discard outputs and retry")
    if sha((root / BAKERY_SOURCE).read_bytes()) != bakery_source_checksum:
        raise ValueError("Bakery source changed during generation; discard outputs and retry")
    if sha((root / STRETCH_SOURCE).read_bytes()) != stretch_checksum:
        raise ValueError("Time To Stretch source changed during generation; discard outputs and retry")
    if sha((root / CONVERSIONS_SOURCE).read_bytes()) != CONVERSIONS_SOURCE_SHA256:
        raise ValueError("Conversions source changed during generation; discard outputs and retry")
    if any((root / path).read_bytes() != content for path, content in zip(DIRECTORY_SOURCES, directory_originals)):
        raise ValueError("Store Directory originals changed during generation; discard outputs and retry")
    command = [sys.executable, str((ROOT / "tools/build_till_binders.py").resolve()),
               "--root", str(root.resolve()), "--binder", str(binder.resolve()),
               "--output", str(output.resolve())]
    if verify_print:
        command.append("--verify-print")
    (output / "REGENERATE.txt").write_text(
        "Exact generation command (output must be empty; use a new directory when rerunning):\n\n" +
        shlex.join(command) + "\n", encoding="utf-8")
    return summary


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--binder", type=Path, required=True, help="Complete binder template (never modified)")
    parser.add_argument("--output", type=Path, required=True, help="Empty directory for the review binder and reports")
    parser.add_argument("--verify-print", action="store_true",
                        help="Export the binder PDF and verify printed Produce, Bakery and Deli records")
    args = parser.parse_args()
    try:
        summary = generate(args.root, args.binder, args.output, args.verify_print)
    except (OSError, ValueError, KeyError, zipfile.BadZipFile, ET.XMLSyntaxError,
            subprocess.SubprocessError) as exc:
        parser.exit(1, f"ERROR: {exc}\n")
    print(json.dumps({key: summary[key] for key in (
        "status", "produce_records_included", "embedded_references", "missing_or_uncertain_images",
        "duplicate_produce_codes", "planned_produce_print_pages", "bakery_records_included")}, indent=2))


if __name__ == "__main__":
    main()
