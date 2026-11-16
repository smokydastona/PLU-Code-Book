"""Build the compact, one-page Bakery sheet from the reviewed product list."""
import csv
import io
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins

from text_fit import line_count

BAKERY_SOURCE = Path("sources/binders/Bakery_Products.csv")
TITLE = "BAKERY PRODUCT & PLU REFERENCE GUIDE"
PRODUCT_WIDTH, CODE_WIDTH, GAP_WIDTH = 40, 8, 2
ROW_HEIGHT = 24
BODY_POINTS = 12
FONT = "DejaVu Sans"
VALID_CATEGORIES = {"Bulk Buns", "Bagels", "Muffins"}


def _read_products(path):
    with Path(path).open(encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        if reader.fieldnames != ["category", "plu", "product"]:
            raise ValueError("Bakery source must have category,plu,product columns")
        rows = list(reader)
    if not rows:
        raise ValueError("Bakery source contains no products")
    previous_category = None
    closed_categories = set()
    result = []
    seen_products = set()
    for line, row in enumerate(rows, 2):
        category, plu, product = (row.get(key) for key in ("category", "plu", "product"))
        if category not in VALID_CATEGORIES or not plu or not plu.isascii() or not plu.isdecimal():
            raise ValueError(f"Bakery source line {line}: invalid category or PLU")
        if not product or not product.strip():
            raise ValueError(f"Bakery source line {line}: product description is blank")
        if category != previous_category:
            if category in closed_categories:
                raise ValueError(f"Bakery source line {line}: category is split into multiple groups")
            if previous_category is not None:
                closed_categories.add(previous_category)
            previous_category = category
        if (plu, product) in seen_products:
            raise ValueError(f"Bakery source line {line}: duplicate product/PLU pair")
        seen_products.add((plu, product))
        result.append(dict(category=category, plu=plu, product=product))
    if set(row["category"] for row in result) != VALID_CATEGORIES:
        raise ValueError("Bakery source must contain all three product groups")
    return result


def _group_panel(rows, column_start, sheet, coordinates):
    if not rows:
        return
    product_column, code_column = column_start, column_start + 1
    row_number = 3
    groups = []
    for product in rows:
        if not groups or groups[-1][0] != product["category"]:
            groups.append((product["category"], []))
        groups[-1][1].append(product)
    border = Border(bottom=Side(style="hair", color="CCCCCC"))
    for category, products in groups:
        sheet.merge_cells(start_row=row_number, start_column=product_column,
                          end_row=row_number, end_column=code_column)
        heading = sheet.cell(row_number, product_column, category)
        heading.data_type = "s"
        heading.font = Font(name=FONT, size=BODY_POINTS + 1, bold=True, color="FFFFFF")
        heading.fill = PatternFill("solid", fgColor="375623")
        sheet.row_dimensions[row_number].height = max(
            sheet.row_dimensions[row_number].height or 0, 24)
        row_number += 1
        for column, label in ((product_column, "Product"), (code_column, "PLU")):
            cell = sheet.cell(row_number, column, label)
            cell.font = Font(name=FONT, size=BODY_POINTS, bold=True, color="375623")
            cell.border = Border(bottom=Side(style="thin", color="375623"))
        sheet.row_dimensions[row_number].height = 20
        row_number += 1
        for entry in products:
            product = sheet.cell(row_number, product_column, entry["product"])
            code = sheet.cell(row_number, code_column, entry["plu"])
            for cell in (product, code):
                cell.data_type = "s"
                cell.font = Font(name=FONT, size=BODY_POINTS)
                cell.alignment = Alignment(wrap_text=True, vertical="center",
                                           horizontal="left" if cell.column == product_column else "center")
                cell.border = border
            code.number_format = "@"
            lines = line_count(entry["product"], BODY_POINTS, PRODUCT_WIDTH)
            sheet.row_dimensions[row_number].height = max(
                sheet.row_dimensions[row_number].height or 0, ROW_HEIGHT + 15 * (lines - 1))
            coordinates.append(dict(
                category=category, product=entry["product"], plu=entry["plu"],
                product_cell=product.coordinate, code_cell=code.coordinate))
            row_number += 1
        row_number += 1


def workbook_bytes(rows):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Bakery"
    sheet.sheet_view.showGridLines = False
    sheet.merge_cells("A1:E1")
    sheet["A1"] = TITLE
    sheet["A1"].data_type = "s"
    sheet["A1"].font = Font(name=FONT, size=16, bold=True, color="375623")
    sheet["A1"].alignment = Alignment(horizontal="center")
    sheet.row_dimensions[1].height = 30
    for column, width in (("A", PRODUCT_WIDTH), ("B", CODE_WIDTH), ("C", GAP_WIDTH),
                          ("D", PRODUCT_WIDTH), ("E", CODE_WIDTH)):
        sheet.column_dimensions[column].width = width

    boundaries = [index for index in range(1, len(rows))
                  if rows[index - 1]["category"] != rows[index]["category"]]
    if not boundaries:
        raise ValueError("Bakery needs at least two product groups")
    midpoint = min(boundaries, key=lambda index: abs(len(rows) - 2 * index))
    coordinates = []
    _group_panel(rows[:midpoint], 1, sheet, coordinates)
    _group_panel(rows[midpoint:], 4, sheet, coordinates)
    sheet.print_area = f"A1:E{sheet.max_row}"
    sheet.print_options.horizontalCentered = True
    sheet.page_setup.orientation = "portrait"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_LETTER
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_margins = PageMargins(left=.25, right=.25, top=.3, bottom=.3, header=.1, footer=.1)
    sheet.freeze_panes = "A3"
    content = io.BytesIO()
    workbook.save(content)
    workbook.close()
    return content.getvalue(), coordinates


def validate_sheet(sheet, expected, coordinates):
    if sheet["A1"].value != TITLE or sheet["A1"].data_type != "s":
        raise ValueError("Bakery title changed")
    if (sheet.page_setup.orientation != "portrait" or
            str(sheet.page_setup.paperSize) != str(sheet.PAPERSIZE_LETTER) or
            sheet.page_setup.fitToWidth != 1 or sheet.page_setup.fitToHeight != 1):
        raise ValueError("Bakery must be configured to print on one portrait US Letter page")
    if len(coordinates) != len(expected):
        raise ValueError("Bakery entry count differs from the source list")
    for entry, source in zip(coordinates, expected):
        if (entry["category"], entry["product"], entry["plu"]) != (
                source["category"], source["product"], source["plu"]):
            raise ValueError("Bakery layout order differs from the source list")
        product, code = sheet[entry["product_cell"]], sheet[entry["code_cell"]]
        if product.value != entry["product"] or product.data_type != "s":
            raise ValueError(f"Bakery product changed at {entry['product_cell']}")
        if code.value != entry["plu"] or code.data_type != "s" or code.number_format != "@":
            raise ValueError(f"Bakery PLU changed or lost its text format at {entry['code_cell']}")


def read_source(root):
    source_path = Path(root) / BAKERY_SOURCE
    rows = _read_products(source_path)
    content, coordinates = workbook_bytes(rows)
    generated = load_workbook(io.BytesIO(content))
    try:
        if generated.sheetnames != ["Bakery"]:
            raise ValueError("generated Bakery source must contain only the Bakery worksheet")
        validate_sheet(generated["Bakery"], rows, coordinates)
    finally:
        generated.close()
    return content, coordinates, rows
