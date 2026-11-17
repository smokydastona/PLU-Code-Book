import csv
import io
import json
import shutil
import subprocess
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from unittest.mock import patch

from docx import Document
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font
from openpyxl.worksheet.pagebreak import Break
from PIL import Image, ImageChops, ImageOps

from build_till_binders import (BULK_SHEETS, CONVERSIONS_ACCURACY_REVIEW, CONVERSIONS_SOURCE, CONVERSIONS_SOURCE_SHA256, DELIVERY_HEADER_BANDS, DELIVERY_SHEET, DIRECTORY_EMERGENCY_HIGHLIGHTS, DIRECTORY_SOURCES, FILENAMES, MAX_GITHUB_FILE_SIZE, PDF_EXPORT_FORMAT, STRETCH_SOURCE, VISUALS, bulk_print_pages,
                               MISC_TITLE, SOIL_BARCODE_SOURCES, SOIL_BARCODE_WIDTH, SOIL_BARCODE_HEIGHT, SOIL_PRODUCT_IMAGE_SIZE, soil_product_image, add_continuation_labels, combine_bulk, comparisons, decoded_soil_barcode, directory_sheet, fit_misc_sheet, generate, layout_rows, misc_page_matches, misc_print_text, normalize_soil_barcodes, order_misc_before_soil, restore_floral_codes, style_soil_guide,
                               docx_compatible_image,
                               delivery_sheet, map_images, produce_catalog, quick_sheet, replace_produce, replace_sheet,
                               sha, sheet_info, stretch_sheet, theme_conversions_sheet, theme_operational_sheets,
                               theme_product_table,
                               themed_directory_image, themed_delivery_image, validate_delivery, validate_package, visual_records,
                               workbook_photo_copy)
from picklist_catalog import read_picklist, read_assignments, write_record_baseline
from review_image_references import load_references
from review_booklets import build_review_booklets, build_working_list_review
from organic_watermarks import render_watermark
from build_deli_sheet import convert
from build_bulk_sheet import (bulk_sheet, read_bulk, remove_discontinued_bulk_entries,
                              validate_bulk)
from build_bakery_sheet import (BAKERY_SOURCE, read_source as read_bakery_source,
                                validate_sheet as validate_bakery_sheet, workbook_bytes as bakery_workbook_bytes)
from binder_theme import PALE_GREEN, display_product_name
from build_weighed_sheet import SOURCE as WEIGHED_SOURCE


class TillBinderChecks(unittest.TestCase):
    def test_original_bulk_sections_are_combined_without_losing_cells_or_other_parts(self):
        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "combined.xlsx"
            self.assertTrue(combine_bulk(source, output))
            original = load_workbook(source)
            combined = load_workbook(output)
            try:
                self.assertEqual(combined.sheetnames, [name for name in original.sheetnames
                                                      if name not in BULK_SHEETS[1:]])
                sheet = combined["Bulk"]
                self.assertEqual([b.id for b in sheet.row_breaks.brk], [41, 73, 118])
                self.assertEqual(str(sheet.print_area), "'Bulk'!$B$4:$I$155")
                offset = 0
                for name, end in zip(BULK_SHEETS, (41, 35, 48, 40)):
                    for row in original[name].iter_rows(min_row=4, max_row=end, min_col=2, max_col=9):
                        for cell in row:
                            result = sheet.cell(cell.row + offset, cell.column)
                            self.assertEqual(result.value, cell.value)
                            self.assertEqual(result.style_id, cell.style_id)
                    offset += end - 3
                self.assertIsNone(bulk_print_pages(output))
                with zipfile.ZipFile(source) as before, zipfile.ZipFile(output) as after:
                    for name in before.namelist():
                        if name not in {"xl/workbook.xml", "xl/worksheets/sheet9.xml"}:
                            self.assertEqual(before.read(name), after.read(name), name)
                second = Path(directory) / "second.xlsx"
                self.assertFalse(combine_bulk(output, second))
                self.assertEqual(output.read_bytes(), second.read_bytes())
            finally:
                original.close()
                combined.close()

    def test_six_page_bulk_theme_preserves_all_records_and_rebuilds_from_updated_template(self):
        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        with tempfile.TemporaryDirectory() as directory:
            combined = Path(directory) / "combined.xlsx"
            combine_bulk(source, combined)
            workbook = load_workbook(combined)
            try:
                entries = read_bulk(workbook["Bulk"])
            finally:
                workbook.close()
            content, coordinates = bulk_sheet(entries)
            bulk_path = Path(directory) / "bulk.xlsx"
            bulk_path.write_bytes(content)
            bulk_pages = bulk_print_pages(bulk_path)
            self.assertEqual(sum(sum(page["codes"].values()) for page in bulk_pages), 400)
            self.assertEqual(sum(len(page["products"]) for page in bulk_pages), 400)
            themed = load_workbook(io.BytesIO(content))
            try:
                validate_bulk(themed["Bulk"], coordinates)
                self.assertEqual(len(coordinates), 400)
                self.assertEqual(len(themed["Bulk"].row_breaks.brk), 5)
                self.assertEqual([sum(entry["page"] == page for entry in coordinates)
                                  for page in range(1, 7)], [66, 68, 77, 86, 43, 60])
                self.assertFalse(themed["Bulk"].sheet_view.showGridLines)
                rebuilt, new_coordinates = bulk_sheet(read_bulk(themed["Bulk"]))
                self.assertEqual([(e["category"], e["product"], e["plu"], e["page"]) for e in coordinates],
                                 [(e["category"], e["product"], e["plu"], e["page"]) for e in new_coordinates])
                regenerated = load_workbook(io.BytesIO(rebuilt))
                try:
                    validate_bulk(regenerated["Bulk"], new_coordinates)
                finally:
                    regenerated.close()
            finally:
                themed.close()
            with self.assertRaisesRegex(ValueError, "categories have changed"):
                altered = load_workbook(io.BytesIO(content))
                try:
                    altered["Bulk"]["A4"] = "Unknown category"
                    read_bulk(altered["Bulk"])
                finally:
                    altered.close()

    def test_bulk_merge_rejects_unprinted_owner_data(self):
        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        with tempfile.TemporaryDirectory() as directory:
            changed = Path(directory) / "changed.xlsx"
            original = load_workbook(source)
            original["Bulk (2)"]["B40"] = "Do not discard"
            original.save(changed)
            original.close()
            with self.assertRaisesRegex(ValueError, "outside its print area"):
                combine_bulk(changed, Path(directory) / "result.xlsx")

    @unittest.skipUnless(shutil.which("soffice"), "LibreOffice Calc is required")
    def test_bulk_continuations_print_adjacent_at_half_size_without_clipping(self):
        import pymupdf

        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        combined = self.root / "combined-bulk.xlsx"
        combine_bulk(source, combined)
        workbook = load_workbook(combined)
        try:
            entries = read_bulk(workbook["Bulk"])
        finally:
            workbook.close()
        labels = []
        content, coordinates = bulk_sheet(entries, labels)
        workbook = load_workbook(io.BytesIO(content))
        try:
            validate_bulk(workbook["Bulk"], coordinates)
            for label in labels:
                cell = workbook["Bulk"].cell(label["row"] + 1, label["column"] + 1)
                self.assertEqual(cell.value, label["title"])
                self.assertEqual(cell.alignment.horizontal, "center")
                self.assertEqual(label["size"] * 2, cell.font.sz)
        finally:
            workbook.close()
        path = self.root / "bulk-inline.xlsx"
        path.write_bytes(add_continuation_labels(content, labels, "Bulk"))
        subprocess.run(["soffice", "-env:UserInstallation=" + (self.root / "lo-bulk-inline").as_uri(),
                        "--headless", "--convert-to", "pdf", "--outdir", str(self.root), str(path)],
                       check=True, capture_output=True, timeout=120)
        printed = []
        with pymupdf.open(self.root / "bulk-inline.pdf") as document:
            for page in document:
                spans = [span for block in page.get_text("dict")["blocks"] if "lines" in block
                         for line in block["lines"] for span in line["spans"]]
                for label in [span for span in spans if "(continued" in span["text"]]:
                    self.assertEqual(label["text"], "(continued)")
                    midpoint = (label["bbox"][1] + label["bbox"][3]) / 2
                    candidates = [span for span in spans if "Bold" in span["font"]
                                  and span["bbox"][2] <= label["bbox"][0]
                                  and abs((span["bbox"][1] + span["bbox"][3]) / 2 - midpoint) < 3]
                    self.assertTrue(candidates)
                    title = min(candidates, key=lambda span: label["bbox"][0] - span["bbox"][2])
                    gap = label["bbox"][0] - title["bbox"][2]
                    self.assertGreater(gap, title["size"] * .15)
                    self.assertLess(gap, title["size"] * .5)
                    self.assertAlmostEqual(label["size"] / title["size"], .5, delta=.01)
                    self.assertNotIn("Bold", label["font"])
                    printed.append(title["text"])
        self.assertCountEqual(printed, [label["title"] for label in labels])
        self.assertIn("Wrapped Candy/Chocolate", printed)

    def test_delivery_charges_preserves_two_ordered_pages_with_green_headers(self):
        source = load_workbook(self.binder)
        try:
            content, images = delivery_sheet(source[DELIVERY_SHEET])
        finally:
            source.close()
        themed = load_workbook(io.BytesIO(content))
        try:
            sheet = themed[DELIVERY_SHEET]
            validate_delivery(sheet, images)
            self.assertEqual([item.id for item in sheet.row_breaks.brk], [49])
            self.assertEqual([image.anchor._from.row for image in sheet._images], [0, 50])
            for index, (original, result) in enumerate(zip(self.delivery_originals, images)):
                with Image.open(io.BytesIO(original)) as before, Image.open(io.BytesIO(result)) as after:
                    before, after = before.convert("RGB"), after.convert("RGB")
                    band_bottom = (496, 1340)[index]
                    self.assertIsNone(ImageChops.difference(
                        before.crop((0, band_bottom, 1700, 2200)),
                        after.crop((0, band_bottom, 1700, 2200))).getbbox())
                    self.assertIsNotNone(ImageChops.difference(
                        before.crop(DELIVERY_HEADER_BANDS[index][0]),
                        after.crop(DELIVERY_HEADER_BANDS[index][0])).getbbox())
        finally:
            themed.close()

    def test_complete_build_keeps_six_bulk_pages_with_original_and_updated_templates(self):
        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        source_checksum = sha(source.read_bytes())
        can_verify_print = bool(shutil.which("soffice"))
        first = generate(self.root, source, self.output, verify_print=can_verify_print)
        self.assertEqual(first["validations"]["photos"]["bulk_records_verified"], 399)
        booklet = self.output / "Pocket Guide.pdf"
        if can_verify_print:
            self.assertTrue(booklet.is_file())
            self.assertEqual(
                first["print_validation"]["photos"]["booklet_imposition"]["physical_sheets"] * 2,
                first["print_validation"]["photos"]["booklet_imposition"]["imposed_pdf_pages"],
            )
        updated = self.output / FILENAMES["photos"]
        book = load_workbook(updated)
        try:
            entries = read_bulk(book["Bulk"])
            self.assertEqual(len(entries), 399)
            self.assertFalse(any(
                (entry["category"], entry["product"], entry["plu"]) ==
                ("Chocolate/Yogurt", "Almonds Roasted Dark Chocolate", "DC")
                for entry in entries))
        finally:
            book.close()
        with (self.output / "BULK_EXCLUDED_RECORDS.csv").open(newline="", encoding="utf-8-sig") as file:
            exclusion, = csv.DictReader(file)
        self.assertEqual((exclusion["category"], exclusion["product"], exclusion["plu"],
                          exclusion["source_sheet"], exclusion["source_product_cell"],
                          exclusion["source_code_cell"], exclusion["status"]),
                         ("Chocolate/Yogurt", "Almonds Roasted Dark Chocolate", "DC",
                          "Bulk", "H11", "I11", "EXCLUDED"))
        second_output = self.root / "second-bulk"
        second = generate(self.root, updated, second_output)
        self.assertEqual(second["validations"]["photos"]["bulk_records_verified"], 399)
        self.assertFalse(second["validations"]["photos"]["bulk_sheets_combined"])
        first_pages = bulk_print_pages(updated)
        second_pages = bulk_print_pages(second_output / FILENAMES["photos"])
        self.assertEqual(len(first_pages), 6)
        self.assertEqual(first_pages, second_pages)
        with (second_output / "BULK_EXCLUDED_RECORDS.csv").open(newline="", encoding="utf-8-sig") as file:
            repeated_exclusion, = csv.DictReader(file)
        self.assertEqual(repeated_exclusion["status"], "ALREADY_ABSENT")
        self.assertEqual(sha(source.read_bytes()), source_checksum)

    def test_bulk_discontinuation_requires_unique_matching_source_identity(self):
        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        combined = Path(self.root) / "bulk-discontinued-source.xlsx"
        combine_bulk(source, combined)
        workbook = load_workbook(combined)
        try:
            entries = read_bulk(workbook["Bulk"])
        finally:
            workbook.close()
        filtered, audit = remove_discontinued_bulk_entries(entries)
        self.assertEqual(len(filtered), 399)
        self.assertEqual(audit[0]["status"], "EXCLUDED")
        self.assertEqual((audit[0]["source_product_cell"], audit[0]["source_code_cell"]), ("H11", "I11"))
        with self.assertRaisesRegex(ValueError, "source identity changed"):
            remove_discontinued_bulk_entries([dict(audit[0], source_product_cell="H12")])

    @unittest.skipUnless(shutil.which("soffice"), "LibreOffice Calc is required")
    def test_store_directory_prints_two_complete_original_images(self):
        import pymupdf

        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        original = load_workbook(source)
        try:
            content, images = directory_sheet(original["Store Directory"])
        finally:
            original.close()
        workbook = self.root / "directory.xlsx"
        workbook.write_bytes(content)
        subprocess.run(["soffice", "-env:UserInstallation=" + (self.root / "lo-directory").as_uri(),
                        "--headless", "--convert-to", PDF_EXPORT_FORMAT, "--outdir", str(self.root),
                        str(workbook)], check=True, capture_output=True, timeout=120)
        with pymupdf.open(self.root / "directory.pdf") as document:
            self.assertEqual(len(document), 2)
            for page, data in zip(document, images):
                drawn = page.get_image_info(xrefs=True)
                self.assertEqual(len(drawn), 1)
                with Image.open(io.BytesIO(data)) as original_image:
                    self.assertEqual((drawn[0]["width"], drawn[0]["height"]), original_image.size)
                    with Image.open(io.BytesIO(document.extract_image(drawn[0]["xref"])["image"])) as printed:
                        self.assertEqual(printed.convert("RGB").tobytes(), original_image.convert("RGB").tobytes())
                self.assertTrue(page.rect.contains(pymupdf.Rect(drawn[0]["bbox"])))
                self.assertEqual((round(page.rect.width), round(page.rect.height)), (792, 612))

    def test_directory_highlights_preserve_every_original_nonwhite_pixel(self):
        for index, path in enumerate(DIRECTORY_SOURCES):
            content = (Path(__file__).resolve().parents[1] / path).read_bytes()
            with Image.open(io.BytesIO(content)) as image:
                original = image.convert("RGB").transpose(Image.Transpose.ROTATE_270)
            with Image.open(io.BytesIO(themed_directory_image(content, index))) as image:
                themed = image.convert("RGB")
            white = [channel.point(lambda value: 255 if value == 255 else 0)
                     for channel in original.split()]
            white_mask = ImageChops.multiply(ImageChops.multiply(white[0], white[1]), white[2])
            difference = ImageChops.difference(original, themed).convert("L")
            self.assertIsNone(ImageChops.multiply(difference, ImageOps.invert(white_mask)).getbbox())
            self.assertIsNotNone(difference.getbbox())
            self.assertIsNotNone(ImageChops.difference(original.convert("L"), themed.convert("L")).getbbox())
        with self.assertRaisesRegex(ValueError, "changed"):
            themed_directory_image(b"changed", 1)

    def test_directory_list_stripes_preserve_map_and_accept_previous_theme(self):
        source = (self.root / DIRECTORY_SOURCES[0]).read_bytes()
        with Image.open(io.BytesIO(themed_directory_image(source, 0))) as image:
            self.assertEqual(image.getpixel((90, 200)), (255, 255, 255))
            self.assertEqual(image.getpixel((90, 230)), tuple(bytes.fromhex(PALE_GREEN)))
            self.assertEqual(image.getpixel((600, 85)), (255, 255, 255))
            self.assertEqual(image.getpixel((600, 115)), tuple(bytes.fromhex(PALE_GREEN)))
        workbook = load_workbook(self.binder)
        try:
            sheet = workbook["Store Directory"]
            for index, picture in enumerate(sheet._images):
                old = themed_directory_image(self.directory_originals[index], index, row_stripes=False)
                replacement = ExcelImage(io.BytesIO(old))
                replacement.anchor = picture.anchor
                sheet._images[index] = replacement
            _, images = directory_sheet(sheet)
            self.assertEqual(images, self.directory_content)
        finally:
            workbook.close()

    def test_directory_emergency_highlights_are_red_without_changing_map_ink(self):
        content = (self.root / DIRECTORY_SOURCES[1]).read_bytes()
        with Image.open(io.BytesIO(themed_directory_image(content, 1))) as image:
            self.assertEqual(image.getpixel((980, 1240)), (252, 228, 214))
            for box in DIRECTORY_EMERGENCY_HIGHLIGHTS:
                region = image.crop(box)
                colors = [color for _, color in region.getcolors(region.width * region.height)]
                self.assertIn((156, 0, 6), colors)
        workbook = load_workbook(self.binder)
        try:
            sheet = workbook["Store Directory"]
            for index, picture in enumerate(sheet._images):
                replacement = ExcelImage(io.BytesIO(themed_directory_image(
                    self.directory_originals[index], index, emergency_highlights=False, parking_labels=False)))
                replacement.anchor = picture.anchor
                sheet._images[index] = replacement
            _, images = directory_sheet(sheet)
            self.assertEqual(images, self.directory_content)
        finally:
            workbook.close()

    def test_directory_owner_parking_labels_use_green_front_and_red_back(self):
        content = (self.root / DIRECTORY_SOURCES[1]).read_bytes()
        with Image.open(io.BytesIO(themed_directory_image(content, 1))) as image:
            front = image.crop((580, 1226, 900, 1270))
            back = image.crop((1504, 620, 1644, 725))
            for label, fill in ((front, (220, 230, 218)), (back, (252, 228, 214))):
                colors = [color for _, color in label.getcolors(label.width * label.height)]
                self.assertIn(fill, colors)
                self.assertIn((0, 0, 0), colors)
        workbook = load_workbook(self.binder)
        try:
            sheet = workbook["Store Directory"]
            for index, picture in enumerate(sheet._images):
                replacement = ExcelImage(io.BytesIO(themed_directory_image(
                    self.directory_originals[index], index, parking_labels=False)))
                replacement.anchor = picture.anchor
                sheet._images[index] = replacement
            _, images = directory_sheet(sheet)
            self.assertEqual(images, self.directory_content)
        finally:
            workbook.close()

    def test_store_directory_rejects_other_content(self):
        original = load_workbook(self.binder)
        try:
            original["Store Directory"]["A1"] = "Owner data that must not be discarded"
            with self.assertRaisesRegex(ValueError, "no other content"):
                directory_sheet(original["Store Directory"])
        finally:
            original.close()

    def test_store_directory_rejects_changed_template_image(self):
        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        original = load_workbook(source)
        try:
            sheet = original["Store Directory"]
            replacement = ExcelImage(io.BytesIO(self.content))
            replacement.anchor = sheet._images[1].anchor
            sheet._images[1] = replacement
            with self.assertRaisesRegex(ValueError, "page 2 changed"):
                directory_sheet(sheet)
        finally:
            original.close()

    def test_supplier_jpeg_is_losslessly_made_docx_compatible(self):
        source = Path(__file__).resolve().parents[1] / "data/review-references/buck-brand/page-02-image-011.jpeg"
        original = source.read_bytes()
        converted = docx_compatible_image(original)
        self.assertNotEqual(converted, original)
        document = Document()
        document.add_paragraph().add_run().add_picture(io.BytesIO(converted))
        with Image.open(io.BytesIO(original)) as before, Image.open(io.BytesIO(converted)) as after:
            self.assertEqual(before.size, after.size)
            self.assertEqual(before.convert("RGB").tobytes(), after.convert("RGB").tobytes())
        self.assertEqual(source.read_bytes(), original)

    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        for folder in ("sources/plu", "sources/visual", "data"):
            (self.root / folder).mkdir(parents=True)
        (self.root / "sources/binders").mkdir(parents=True)
        bakery_source = Path(__file__).resolve().parents[1] / BAKERY_SOURCE
        shutil.copyfile(bakery_source, self.root / BAKERY_SOURCE)
        (self.root / STRETCH_SOURCE).write_bytes(
            (Path(__file__).resolve().parents[1] / STRETCH_SOURCE).read_bytes())
        for path in DIRECTORY_SOURCES:
            (self.root / path).write_bytes((Path(__file__).resolve().parents[1] / path).read_bytes())
        (self.root / CONVERSIONS_SOURCE).parent.mkdir(parents=True, exist_ok=True)
        (self.root / CONVERSIONS_SOURCE).write_bytes(
            (Path(__file__).resolve().parents[1] / CONVERSIONS_SOURCE).read_bytes())
        (self.root / WEIGHED_SOURCE).write_bytes(
            (Path(__file__).resolve().parents[1] / WEIGHED_SOURCE).read_bytes())
        self.binder = self.root / "Original.xlsx"
        image = io.BytesIO()
        Image.new("RGB", (600, 500), "green").save(image, format="PNG")
        self.content = image.getvalue()
        self.image = self.root / "data/reference.png"
        self.image.write_bytes(self.content)
        original = Workbook()
        first = original.active
        first.title = "Time To Stretch"
        first.add_image(ExcelImage(io.BytesIO(self.content)), "A1")
        first = original.create_sheet("Operational")
        first.title = "Operational"
        first["A1"] = "Do not change"
        first["A1"].font = Font(bold=True, color="FF0055")
        first.merge_cells("A1:C1")
        first["A2"], first["A3"] = 2, 3
        first["B2"] = "=SUM(A2:A3)"
        first["B2"].comment = Comment("Keep this comment", "Fixture")
        first.column_dimensions["A"].width = 28
        first.row_dimensions[2].height = 40
        first.freeze_panes = "B2"
        first.print_area = "A1:D10"
        first.print_title_rows = "1:1"
        first.add_image(ExcelImage(io.BytesIO(self.content)), "C3")
        chart = BarChart()
        chart.add_data(Reference(first, min_col=1, min_row=2, max_row=3))
        first.add_chart(chart, "E1")
        misc = original.create_sheet("Misc Codes")
        misc["A1"], misc["D1"] = "Miscellaneous items", "Bags and deposits"
        misc["A2"], misc["D2"] = "00012", "Owner supplied entry"
        misc["D2"].font = Font(bold=True, color="FF0055")
        misc.column_dimensions["A"].width = 45
        misc.column_dimensions["D"].width = 45
        misc.row_breaks.append(Break(id=1))
        misc.col_breaks.append(Break(id=2))
        bakery = original.create_sheet("Bakery")
        bakery["A1"], bakery["B1"] = "Old bakery", "Old entry"
        produce = original.create_sheet("Produce")
        produce["A1"], produce["B1"] = 42, "Old Gala"
        produce["D1"], produce["E1"] = "900042", "Organic Apple Gala"
        produce["H1"], produce["I1"] = "99999", "Removed"
        produce["A2"] = "CATEGORY"
        produce.merge_cells("A2:B2")
        produce["A3"], produce["B3"] = " 00042 ", "Apple Gala"
        produce["A4"], produce["B4"] = "00046", "Hidden Apple"
        produce["B5"] = "Description without code"
        produce.print_area = "A1:I6"
        deli = original.create_sheet("Deli")
        deli["A1"] = "Outdated deli"
        deli.print_area = "A1:B2"
        delivery = original.create_sheet(DELIVERY_SHEET)
        source_binder = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        source_workbook = load_workbook(source_binder)
        try:
            source_sheet = source_workbook[DELIVERY_SHEET]
            self.delivery_originals = [image._data() for image in sorted(
                source_sheet._images, key=lambda image: (image.anchor._from.row, image.anchor._from.col))]
        finally:
            source_workbook.close()
        self.delivery_content = [themed_delivery_image(content, index)
                                 for index, content in enumerate(self.delivery_originals)]
        for index, content in enumerate(self.delivery_originals):
            delivery.add_image(ExcelImage(io.BytesIO(content)), f"A{1 + 50 * index}")
        directory = original.create_sheet("Store Directory")
        self.directory_originals = [(self.root / path).read_bytes() for path in DIRECTORY_SOURCES]
        self.directory_content = [themed_directory_image(content, i)
                                  for i, content in enumerate(self.directory_originals)]
        directory.add_image(ExcelImage(io.BytesIO(self.directory_originals[0])), "A1")
        directory.add_image(ExcelImage(io.BytesIO(self.directory_originals[1])), "A30")
        last = original.create_sheet("Hidden operations")
        last.sheet_state = "hidden"
        last["A1"] = "=Operational!B2"
        original.save(self.binder)
        deli_document = Document()
        deli_document.add_heading("DELI PRODUCT & PLU REFERENCE GUIDE", 0)
        for category, product, plu in (("Meals", "Updated meal Organic", "00012"),
                                       ("Salads", "Updated salad", "00013")):
            deli_document.add_heading(category, 1)
            table = deli_document.add_table(rows=2, cols=2)
            table.cell(0, 0).text, table.cell(0, 1).text = "Product", "PLU"
            table.cell(1, 0).text, table.cell(1, 1).text = product, plu
        upload = self.root / "Uploaded_Deli.docx"
        deli_document.save(upload)
        convert(upload, self.root / "sources/binders")
        with zipfile.ZipFile(self.binder, "a") as archive:
            archive.writestr("customXml/untouched.xml", b"<custom>unchanged opaque object</custom>")
        self.write_source([
            ("900042", "Organic Apple Gala", True, False),
            ("00042", "Apple Gala", True, False),
            ("00042", "Different Apple", True, False),
            ("00044", "Questionable Fruit", True, False),
            ("00045", "AI Apple", True, False),
            ("00046", "Hidden Apple", False, False),
            ("00047", "Deleted Apple", True, True),
            ("00048", "Operational Code", True, False),
            ("00049", "Organic Melon New", True, False),
            ("00050", "=Unusual Apple", True, False),
        ])
        self.items, self.digest = read_picklist(self.root / "sources/plu/PickList.xml")
        write_record_baseline(
            self.root / "data/picklist_record_baseline.json", self.items, self.digest)
        (self.root / "data/working_list_decisions.json").write_text(json.dumps(dict(
            schema_version=1, baseline_source_sha256=self.digest,
            baseline=[dict(source_id=item["source_id"], plu=item["plu"], name=item["name"],
                           department=item["department"], occurrence=item["occurrence"],
                           visible=item["visible"], deleted=item["deleted"]) for item in self.items],
            decisions=[])))
        assignments, reviews = [], []
        fruit = Document()
        for number, item in enumerate(self.items):
            if number in (5, 6, 8):
                continue
            assignment = dict(
                source_id=item["source_id"], book="Fruit", section="Organic Fruit" if number == 0 else "Apples",
                image="data/reference.png", image_mode="PHOTO", reviewed_plu=item["plu"],
                reviewed_product=item["name"], legacy_entry_idx=str(number + 1),
                image_sha256=sha(self.content), visual_approval_status="NOT_APPROVED",
                photographic_origin="NOT_VERIFIED", rights_verification_status="NOT_VERIFIED", provenance={},
            )
            if number == 7:
                assignment.update(book="Non-produce", section="Misc", image="",
                                  image_mode="CODE_ONLY", code_only_approved=True)
            else:
                source = "AI Generated" if number == 4 else "Code Booklet"
                cell = fruit.add_table(rows=1, cols=1).cell(0, 0)
                cell.text = f"PLU: {item['plu']}"
                for text in (item["name"], "Category: Apple", "Image Source: " + source):
                    cell.add_paragraph(text)
                cell.add_paragraph().add_run().add_picture(io.BytesIO(self.content))
                reviews.append(dict(idx=str(number + 1), plu=item["plu"], name=item["name"],
                                    verdict="Approximate" if number == 3 else "Approved"))
            assignments.append(assignment)
        fruit.save(self.root / "sources/visual" / VISUALS["Fruit"])
        Document().save(self.root / "sources/visual" / VISUALS["Vegetable"])
        (self.root / "data/picklist_assignments.json").write_text(json.dumps(
            dict(schema_version=1, assignments=assignments)))
        with (self.root / "data/image_review.csv").open("w", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=["idx", "plu", "name", "verdict"])
            writer.writeheader()
            writer.writerows(reviews)
        self.output = self.root / "output"

    def write_source(self, values):
        root = ET.Element("PickList")
        items = ET.SubElement(root, "Items")
        for plu, name, visible, deleted in values:
            item = ET.SubElement(items, "Item", UPC=plu)
            ET.SubElement(item, "Department").text = "001"
            ET.SubElement(item, "IsVisible").text = str(visible)
            ET.SubElement(item, "DeleteThisItem").text = str(deleted)
            language = ET.SubElement(ET.SubElement(item, "Languages"), "Language", LCID="0409")
            ET.SubElement(language, "Description").text = name
        (self.root / "sources/plu/PickList.xml").write_bytes(ET.tostring(root, encoding="utf-16"))

    def test_full_workbooks_preserve_opaque_parts_formulas_objects_and_original(self):
        before = self.binder.read_bytes()
        source_before = (self.root / "sources/plu/PickList.xml").read_bytes()
        summary = generate(self.root, self.binder, self.output)
        self.assertEqual(list(FILENAMES), ["photos"])
        self.assertEqual(len(list(self.output.glob("*.xlsx"))), 1)
        self.assertFalse((self.output / "booklets").exists())
        self.assertEqual(summary["produce_records_included"], 7)
        self.assertEqual(summary["bakery_records_included"], 31)
        self.assertEqual(summary["total_source_records"], 10)
        self.assertEqual(summary["hidden_or_deleted_records"], 2)
        self.assertEqual(summary["known_non_produce_excluded"], 1)
        self.assertEqual(summary["duplicate_produce_codes"], 1)
        self.assertEqual(summary["embedded_references"], 5)
        self.assertEqual(summary["ai_illustrations_omitted"], 0)
        self.assertEqual(summary["ai_illustrations_embedded"], 1)
        self.assertEqual(summary["production_publication"], "BLOCKED")
        for edition, filename in FILENAMES.items():
            path = self.output / filename
            workbook = load_workbook(path)
            self.addCleanup(workbook.close)
            self.assertEqual(workbook.sheetnames, ["Time To Stretch", "Operational", "Misc Codes",
                                                   "Bakery", "Produce", "Deli", DELIVERY_SHEET, "Store Directory",
                                                   "Hidden operations"])
            directory = workbook["Store Directory"]
            self.assertEqual([image._data() for image in directory._images], self.directory_content)
            self.assertEqual([item.id for item in directory.row_breaks.brk], [36])
            delivery = workbook[DELIVERY_SHEET]
            validate_delivery(delivery, self.delivery_content)
            self.assertEqual(summary["validations"][edition]["delivery_pages_verified"], 2)
            misc = workbook["Misc Codes"]
            self.assertEqual(misc["A2"].value, "00012")
            self.assertEqual(misc["D2"].value, "Owner supplied entry")
            self.assertTrue(misc["D2"].font.bold)
            self.assertEqual(misc.column_dimensions["D"].width, 45)
            self.assertEqual((misc.page_setup.fitToWidth, misc.page_setup.fitToHeight), (1, 1))
            self.assertTrue(misc.sheet_properties.pageSetUpPr.fitToPage)
            self.assertFalse(misc.row_breaks.brk)
            self.assertFalse(misc.col_breaks.brk)
            stretch = workbook["Time To Stretch"]
            self.assertEqual(len(stretch._images), 1)
            self.assertEqual(stretch._images[0]._data(), (self.root / STRETCH_SOURCE).read_bytes())
            self.assertEqual(stretch.page_setup.orientation, "portrait")
            self.assertEqual(stretch.page_setup.fitToWidth, 1)
            self.assertEqual(stretch.page_setup.fitToHeight, 1)
            self.assertEqual(summary["stretch_source_sha256"], sha((self.root / STRETCH_SOURCE).read_bytes()))
            self.assertEqual(workbook["Deli"]["A5"].value, "00012")
            self.assertEqual(workbook["Deli"]["B5"].value, "Updated meal Org.")
            self.assertEqual(workbook["Deli"]["D5"].value, "00013")
            self.assertEqual(workbook["Deli"]["E5"].value, "Updated salad")
            for name in ("Bakery", "Bulk", "Deli", "Produce", "Misc Codes"):
                if name in workbook.sheetnames:
                    self.assertTrue(all(
                        display_product_name(cell.value) == cell.value
                        for row in workbook[name]
                        for cell in row
                        if isinstance(cell.value, str)
                    ))
            self.assertEqual(summary["validations"][edition]["deli_records_verified"], 2)
            _, bakery_coordinates, bakery_products = read_bakery_source(self.root)
            validate_bakery_sheet(workbook["Bakery"], bakery_products, bakery_coordinates)
            self.assertEqual(summary["validations"][edition]["bakery_records_verified"], 31)
            self.assertEqual(workbook["Operational"]["B2"].value, "=SUM(A2:A3)")
            self.assertEqual(workbook["Operational"]["B2"].comment.text, "Keep this comment")
            self.assertEqual(str(workbook["Operational"].merged_cells), "A1:C1")
            self.assertEqual(len(workbook["Operational"]._images), 1)
            self.assertEqual(len(workbook["Operational"]._charts), 1)
            self.assertEqual(workbook["Hidden operations"].sheet_state, "hidden")
            self.assertTrue(summary["validations"][edition]["unrelated_package_parts_byte_identical"])
            if edition == "photos":
                self.assertLess(summary["validations"][edition]["file_size_bytes"], MAX_GITHUB_FILE_SIZE)
            self.assertEqual(len(workbook["Produce"]._images), 5 if edition == "photos" else 0)
            if edition == "photos":
                formats = set()
                for picture in workbook["Produce"]._images:
                    with Image.open(io.BytesIO(picture._data())) as image:
                        formats.add(image.format)
                self.assertEqual(formats, {"PNG"})
            self.assertTrue(all(cell.data_type != "f" for row in workbook["Produce"] for cell in row))
            with zipfile.ZipFile(path) as archive:
                self.assertEqual(archive.read("customXml/untouched.xml"), b"<custom>unchanged opaque object</custom>")
        self.assertEqual(self.binder.read_bytes(), before)
        self.assertEqual((self.root / "sources/plu/PickList.xml").read_bytes(), source_before)
        self.assertEqual(self.image.read_bytes(), self.content)
        self.assertTrue((self.output / "PRODUCE_CHANGE_AND_VALIDATION_REPORT.txt").is_file())
        self.assertTrue((self.output / "BAKERY_LAYOUT.csv").is_file())
        with self.assertRaisesRegex(ValueError, "empty"):
            generate(self.root, self.binder, self.output)

    def test_bakery_sheet_keeps_all_exact_text_entries_in_one_page_theme(self):
        content, coordinates, products = read_bakery_source(self.root)
        self.assertEqual(len(products), 31)
        self.assertEqual([sum(row["category"] == category for row in products)
                          for category in ("Bulk Buns", "Bagels", "Muffins", "Sweets")], [10, 8, 6, 7])
        self.assertEqual([row["product"] for row in products if row["category"] == "Bagels"], [
            "Plain", "Blueberry", "Cheese", "Cinnamon Raisin", "Everything",
            "Jalapeno Cheese", "Multigrain", "Sesame Seed",
        ])
        self.assertTrue(all(entry["code_cell"][0] < entry["product_cell"][0] for entry in coordinates))
        removed = {"55854", "47898", "55935", "55865", "68342", "68343", "55756", "55861"}
        self.assertTrue(removed.isdisjoint(row["plu"] for row in products))
        self.assertEqual([row["plu"] for row in products if row["category"] == "Muffins"], [
            "65873", "42114", "66136", "2326", "73799", "42103",
        ])
        self.assertEqual(
            [(row["category"], row["plu"], row["product"]) for row in products
             if row["category"] == "Sweets"],
            [("Sweets", plu, product) for plu, product in (
                ("16799", "Brioche Swirl Chocolate"),
                ("63886", "Brioche Swirl Raisin"),
                ("22835", "Blue Danish"),
                ("22833", "Apple Danish"),
                ("60068", "Apple Fritter"),
                ("40302", "Glazed Donut Ring"),
                ("40457", "Chocolate Glazed Croisant"),
            )],
        )
        entries = {(row["category"], row["plu"], row["product"]) for row in products}
        self.assertTrue({
            ("Bulk Buns", "67721", "Kaiser/Cheese Kaiser"),
            ("Bulk Buns", "67721", "Kaisers Buns White Cheese Bulk"),
            ("Bulk Buns", "77474", "Cheese Swirl Buns"),
            ("Bulk Buns", "37769", "Cheese Sticks"),
            ("Muffins", "42114", "Carrot w/ cream cheese"),
            ("Muffins", "42103", "Blueberry w/ Greek Yogurt"),
        }.issubset(entries))
        self.assertEqual([row["plu"] for row in products if row["plu"] == "67721"], ["67721", "67721"])
        removed_bakery_plus = {
            "29684", "76255", "56126", "92655", "2298", "2329", "73798",
            "65619", "2327", "30298", "80228", "34176",
        }
        self.assertTrue(removed_bakery_plus.isdisjoint(row["plu"] for row in products))
        self.assertTrue(all("1X1EA" not in row["product"] for row in products))
        destination = self.root / "bakery-replaced.xlsx"
        replace_sheet(self.binder, content, destination, "Bakery")
        result = load_workbook(destination)
        original = load_workbook(self.binder)
        try:
            validate_bakery_sheet(result["Bakery"], products, coordinates)
            headings = [cell.value for row in result["Bakery"] for cell in row
                        if cell.value in ("Bulk Buns", "Bagels", "Muffins", "Sweets")]
            self.assertCountEqual(headings, ["Bulk Buns", "Bagels", "Muffins", "Sweets"])
            self.assertEqual(result["Produce"]["B3"].value, original["Produce"]["B3"].value)
            self.assertEqual(result["Deli"]["A1"].value, original["Deli"]["A1"].value)
            self.assertEqual(result.sheetnames, original.sheetnames)
        finally:
            result.close()
            original.close()

    def test_organic_product_word_is_abbreviated_case_insensitively(self):
        self.assertEqual(display_product_name("ORGANIC APPLES"), "Org. APPLES")
        self.assertEqual(display_product_name("Whole Roasted Chicken Organic"),
                         "Whole Roasted Chicken Org.")
        self.assertEqual(display_product_name("Organics"), "Organics")

    def test_bakery_groups_are_not_split_when_source_count_changes(self):
        _, _, products = read_bakery_source(self.root)
        content, coordinates = bakery_workbook_bytes(products[:-1])
        workbook = load_workbook(io.BytesIO(content))
        try:
            headings = [cell.value for row in workbook["Bakery"] for cell in row
                        if cell.value in ("Bulk Buns", "Bagels", "Muffins", "Sweets")]
            self.assertCountEqual(headings, ["Bulk Buns", "Bagels", "Muffins", "Sweets"])
            validate_bakery_sheet(workbook["Bakery"], products[:-1], coordinates)
        finally:
            workbook.close()

    def test_existing_binder_can_be_updated_without_losing_other_sheet_edits(self):
        generate(self.root, self.binder, self.output)
        updated_template = self.root / "Edited.xlsx"
        edited = load_workbook(self.output / FILENAMES["photos"])
        edited["Operational"]["A1"] = "Owner's updated operational sheet"
        edited["Misc Codes"]["D2"] = "Owner's updated miscellaneous entry"
        edited["Deli"]["A5"] = "Not the approved deli source"
        edited.save(updated_template)
        edited.close()
        template_bytes = updated_template.read_bytes()
        second = self.root / "second"
        summary = generate(self.root, updated_template, second)
        self.assertTrue(summary["validations"]["photos"]["unrelated_package_parts_byte_identical"])
        result = load_workbook(second / FILENAMES["photos"])
        try:
            self.assertEqual(result["Operational"]["A1"].value, "Owner's updated operational sheet")
            self.assertEqual(result["Misc Codes"]["D2"].value, "Owner's updated miscellaneous entry")
            self.assertEqual((result["Misc Codes"].page_setup.fitToWidth,
                              result["Misc Codes"].page_setup.fitToHeight), (1, 1))
            self.assertEqual([image._data() for image in result["Store Directory"]._images],
                             self.directory_content)
            self.assertEqual(result["Deli"]["A5"].value, "Updated meal")
            self.assertEqual(len(result["Produce"]._images), 5)
        finally:
            result.close()
        self.assertEqual(updated_template.read_bytes(), template_bytes)

    @unittest.skipUnless(shutil.which("soffice"), "LibreOffice Calc is required")
    def test_misc_codes_prints_as_one_page_with_all_original_entries(self):
        from pypdf import PdfReader

        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        combined = self.root / "misc-only.xlsx"
        fit_misc_sheet(source, combined)
        expected = misc_print_text(combined)
        subprocess.run(["soffice", "-env:UserInstallation=" + (self.root / "lo-misc").as_uri(),
                        "--headless", "--convert-to", "pdf", "--outdir", str(self.root),
                        str(combined)], check=True, capture_output=True, timeout=120)
        pdf = PdfReader(self.root / "misc-only.pdf")
        pages = [page for page in pdf.pages if misc_page_matches(page, expected)]
        self.assertEqual(len(pages), 1)
        self.assertIn(MISC_TITLE, pages[0].extract_text())
        self.assertEqual((round(float(pages[0].mediabox.width)), round(float(pages[0].mediabox.height))),
                         (612, 792))
        original = load_workbook(source)
        updated = load_workbook(combined)
        try:
            self.assertEqual(list(original["Misc Codes"].values), list(updated["Misc Codes"].values))
            self.assertEqual([updated["Misc Codes"].cell(27, col).fill.fgColor.rgb
                              for col in range(1, 5)], ["00305040"] * 4)
        finally:
            original.close()
            updated.close()

    def test_misc_theme_preserves_exact_cells_formats_and_unrelated_parts(self):
        source = self.root / "misc-source.xlsx"
        workbook = load_workbook(self.binder)
        workbook["Misc Codes"]["C3"] = 2.75
        workbook["Misc Codes"]["C3"].number_format = '"$"#,##0.00'
        workbook["Misc Codes"]["B4"] = "=C3*2"
        workbook.save(source)
        workbook.close()
        result = self.root / "misc-themed.xlsx"
        fit_misc_sheet(source, result)
        with zipfile.ZipFile(source) as before, zipfile.ZipFile(result) as after:
            part, _ = sheet_info(before, "Misc Codes")
            for name in before.namelist():
                if name not in (part, "xl/styles.xml"):
                    self.assertEqual(before.read(name), after.read(name), name)
            cells = []
            for archive in (before, after):
                nodes = ET.fromstring(archive.read(part)).findall(
                    ".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c")
                for node in nodes:
                    node.attrib.pop("s", None)
                cells.append({node.get("r"): ET.tostring(node) for node in nodes})
            for coordinate, cell in cells[0].items():
                self.assertEqual(cell, cells[1][coordinate])
        workbook = load_workbook(result)
        try:
            sheet = workbook["Misc Codes"]
            self.assertEqual(sheet["A1"].fill.fgColor.rgb, "00305040")
            self.assertEqual(sheet["A1"].font.color.rgb, "00FFFFFF")
            self.assertEqual(sheet["A2"].fill.fgColor.rgb, "00FFFFFF")
            self.assertEqual(sheet["A4"].fill.fgColor.rgb, "00" + PALE_GREEN)
            self.assertEqual(sheet["A1"].fill.fgColor.rgb, sheet["B1"].fill.fgColor.rgb)
            self.assertEqual(sheet["C1"].fill.fgColor.rgb, sheet["D1"].fill.fgColor.rgb)
            self.assertEqual(sheet["A2"].border.bottom.style, "hair")
            self.assertIsNone(sheet["A2"].border.left)
            self.assertEqual(sheet["A2"].font.name, "DejaVu Sans")
            self.assertEqual(sheet["A2"].value, "00012")
            self.assertEqual(sheet["C3"].number_format, '"$"#,##0.00')
            self.assertEqual(sheet["B4"].value, "=C3*2")
            self.assertEqual(sheet["A1"].alignment.horizontal, "center")
            self.assertEqual(sheet["D1"].alignment.horizontal, "center")
            self.assertFalse(sheet.sheet_view.showGridLines)
            self.assertEqual(sheet.oddHeader.center.text, MISC_TITLE)
        finally:
            workbook.close()

    def test_misc_theme_refuses_to_replace_owner_header(self):
        source = self.root / "misc-header.xlsx"
        workbook = load_workbook(self.binder)
        workbook["Misc Codes"].oddHeader.center.text = "Owner instructions"
        workbook.save(source)
        workbook.close()
        with self.assertRaisesRegex(ValueError, "existing print header"):
            fit_misc_sheet(source, self.root / "rejected-header.xlsx")

    def test_soil_and_weighed_theme_preserves_barcode_images_and_instructions(self):
        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        result = self.root / "operational-themed.xlsx"
        theme_operational_sheets(source, result)
        with zipfile.ZipFile(source) as before, zipfile.ZipFile(result) as after:
            parts = {sheet_info(before, name)[0] for name in ("Soil", "Weighed more $100")}
            for name in before.namelist():
                if name not in parts | {"xl/styles.xml"}:
                    self.assertEqual(before.read(name), after.read(name), name)
        original, themed = load_workbook(source), load_workbook(result)
        try:
            soil, updated = original["Soil"], themed["Soil"]
            self.assertEqual(updated["B1"].value, soil["E1"].value)
            self.assertEqual(updated["B1"].fill.fgColor.rgb, "00305040")
            self.assertEqual(updated["B1"].font.color.rgb, "00FFFFFF")
            self.assertIn("B1:J1", str(updated.merged_cells))
            self.assertTrue(updated.print_options.horizontalCentered)
            self.assertEqual(updated.print_area, soil.print_area)
            self.assertEqual([sha(image._data()) for image in updated._images],
                             [sha(image._data()) for image in soil._images])
            self.assertEqual(list(themed["Weighed more $100"].values),
                             list(original["Weighed more $100"].values))
            self.assertEqual(themed["Weighed more $100"]["A3"].font.color.rgb, "00305040")
        finally:
            original.close()
            themed.close()
        again = self.root / "operational-again.xlsx"
        theme_operational_sheets(result, again)
        workbook = load_workbook(again)
        try:
            self.assertEqual(len(workbook["Soil"]._images), 24)
            self.assertEqual(str(workbook["Soil"].merged_cells), "B1:J1")
        finally:
            workbook.close()

    def test_soil_theme_rejects_content_in_heading_area(self):
        source = self.root / "soil-changed.xlsx"
        workbook = Workbook()
        workbook.active.title = "Soil"
        workbook.active["E1"] = "SOIL BARCODES"
        workbook.active["F1"] = "Owner data"
        workbook.save(source)
        workbook.close()
        with self.assertRaisesRegex(ValueError, "additional data"):
            theme_operational_sheets(source, self.root / "rejected-soil.xlsx")

    def test_soil_guide_styling_preserves_images_and_barcode_positions(self):
        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        themed, styled, repeated = (self.root / name for name in
                                    ("soil-theme.xlsx", "soil-guide.xlsx", "soil-guide-again.xlsx"))
        theme_operational_sheets(source, themed)
        style_soil_guide(themed, styled)
        original, result = load_workbook(themed), load_workbook(styled)
        try:
            soil = result["Soil"]
            self.assertEqual(str(soil.print_area), "'Soil'!$A$1:$J$66")
            self.assertEqual(soil.page_setup.scale, 77)
            self.assertEqual((soil.page_margins.top, soil.page_margins.bottom), (.22, .22))
            self.assertEqual([soil[cell].value for cell in ("A2", "B2", "E2", "I2")],
                             ["#", "Product", "Barcode", "Product Image"])
            self.assertEqual([soil[f"A{row}"].value for row in range(3, 67, 8)], list(range(1, 9)))
            self.assertEqual(soil["A3"].fill.fgColor.rgb, "00305040")
            self.assertEqual(soil["B3"].fill.fgColor.rgb, "00FFFFFF")
            self.assertEqual(soil["B11"].fill.fgColor.rgb, "00FFFFFF")
            self.assertEqual([sha(image._data()) for image in original["Soil"]._images],
                             [sha(image._data()) for image in soil._images])
        finally:
            original.close()
            result.close()
        style_soil_guide(styled, repeated)
        result = load_workbook(repeated)
        try:
            self.assertEqual([result["Soil"][f"A{row}"].value for row in range(3, 67, 8)], list(range(1, 9)))
        finally:
            result.close()

    def test_misc_moves_before_soil_without_changing_sheet_parts_or_print_areas(self):
        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        result = self.root / "ordered.xlsx"
        order_misc_before_soil(source, result)
        original, updated = load_workbook(source), load_workbook(result)
        try:
            expected = original.sheetnames[:]
            expected.remove("Misc Codes")
            expected.insert(expected.index("Soil"), "Misc Codes")
            self.assertEqual(updated.sheetnames, expected)
            for name in original.sheetnames:
                self.assertEqual(str(original[name].print_area), str(updated[name].print_area))
        finally:
            original.close()
            updated.close()
        with zipfile.ZipFile(source) as before, zipfile.ZipFile(result) as after:
            for name in before.namelist():
                if name != "xl/workbook.xml":
                    self.assertEqual(before.read(name), after.read(name), name)
        repeated = self.root / "ordered-again.xlsx"
        order_misc_before_soil(result, repeated)
        self.assertEqual(result.read_bytes(), repeated.read_bytes())

    def test_soil_guide_refuses_existing_number_column_data(self):
        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        workbook = load_workbook(source)
        workbook["Soil"]["A5"] = "Owner data"
        edited = self.root / "soil-owner-data.xlsx"
        workbook.save(edited)
        workbook.close()
        with self.assertRaisesRegex(ValueError, "contains owner data"):
            style_soil_guide(edited, self.root / "rejected-soil-guide.xlsx")

    def test_soil_barcodes_preserve_codes_and_have_uniform_dimensions(self):
        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        result = self.root / "soil-barcodes.xlsx"
        drawing_parts = set(normalize_soil_barcodes(source, result))
        with zipfile.ZipFile(source) as before, zipfile.ZipFile(result) as after:
            for name in before.namelist():
                if name not in drawing_parts:
                    self.assertEqual(before.read(name), after.read(name), name)
        old, new = load_workbook(source), load_workbook(result)
        try:
            old_other = {(image.anchor._from.col, image.anchor._from.row): sha(image._data()) for image in old["Soil"]._images
                         if image.anchor._from.col == 1}
            new_other = {(image.anchor._from.col, image.anchor._from.row): sha(image._data()) for image in new["Soil"]._images
                         if image.anchor._from.col == 1}
            self.assertEqual(old_other, new_other)
            expected_products = {image.anchor._from.row: sha(soil_product_image(image._data()))
                                 for image in old["Soil"]._images if image.anchor._from.col == 8}
            actual_products = {}
            found = {}
            for image in new["Soil"]._images:
                marker = image.anchor._from
                if marker.col == 8:
                    self.assertEqual((image.anchor.ext.cx, image.anchor.ext.cy),
                                     (SOIL_PRODUCT_IMAGE_SIZE * 9525,) * 2)
                    self.assertEqual((marker.colOff, marker.rowOff), (10 * 9525, 7 * 9525))
                    actual_products[marker.row] = sha(image._data())
                if marker.col != 4:
                    continue
                self.assertEqual((image.anchor.ext.cx, image.anchor.ext.cy),
                                 (SOIL_BARCODE_WIDTH * 9525, SOIL_BARCODE_HEIGHT * 9525))
                found[marker.row] = decoded_soil_barcode(Image.open(io.BytesIO(image._data())))
            self.assertEqual(found, {row: value[0] for row, value in SOIL_BARCODE_SOURCES.items()})
            self.assertEqual(actual_products, expected_products)
        finally:
            old.close()
            new.close()
        repeated = self.root / "soil-barcodes-again.xlsx"
        normalize_soil_barcodes(result, repeated)
        with zipfile.ZipFile(repeated) as archive:
            self.assertIsNone(archive.testzip())

    def test_soil_product_seven_fills_its_display_box_without_background_padding(self):
        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        workbook = load_workbook(source)
        try:
            photo = next(image for image in workbook["Soil"]._images
                         if image.anchor._from.col == 8 and image.anchor._from.row == 50)
            content = photo._data()
        finally:
            workbook.close()
        heights = []
        for trimmed in (False, True):
            image = Image.open(io.BytesIO(soil_product_image(content, trim_padding=trimmed))).convert("RGB")
            self.assertEqual(image.size, (350, 350))
            mask = Image.new("L", image.size)
            mask.putdata([255 if max(pixel) - min(pixel) > 45 else 0
                          for pixel in image.get_flattened_data()])
            bounds = mask.getbbox()
            self.assertIsNotNone(bounds)
            heights.append(bounds[3] - bounds[1])
        self.assertGreaterEqual(heights[1], 300)
        self.assertGreater(heights[1], heights[0] * 1.1)

    def test_soil_barcodes_reject_changed_source_or_row(self):
        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        edited = self.root / "soil-altered.xlsx"
        workbook = load_workbook(source)
        next(image for image in workbook["Soil"]._images if image.anchor._from.col == 4).anchor._from.row += 1
        workbook.save(edited)
        workbook.close()
        with self.assertRaisesRegex(ValueError, "anchors changed"):
            normalize_soil_barcodes(edited, self.root / "soil-rejected.xlsx")
        generated = self.root / "soil-normalized.xlsx"
        normalize_soil_barcodes(source, generated)
        changed = self.root / "soil-changed-image.xlsx"
        with zipfile.ZipFile(generated) as before, zipfile.ZipFile(changed, "w") as after:
            for item in before.infolist():
                content = before.read(item.filename)
                if item.filename == "xl/media/soil-code128-3.png":
                    image = Image.open(io.BytesIO(content)).convert("RGB")
                    image.putpixel((0, 0), (0, 0, 0))
                    buffer = io.BytesIO()
                    image.save(buffer, format="PNG")
                    content = buffer.getvalue()
                after.writestr(item, content)
        with self.assertRaisesRegex(ValueError, "source changed"):
            normalize_soil_barcodes(changed, self.root / "soil-rejected-image.xlsx")
        workbook = load_workbook(source)
        next(image for image in workbook["Soil"]._images if image.anchor._from.col == 1).anchor._from.row += 1
        edited_label = self.root / "soil-altered-label.xlsx"
        workbook.save(edited_label)
        workbook.close()
        with self.assertRaisesRegex(ValueError, "barcode associations need review"):
            normalize_soil_barcodes(edited_label, self.root / "soil-rejected-label.xlsx")

    def test_floral_codes_remain_on_misc_and_restore_legacy_soil_layout(self):
        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        original = load_workbook(source)
        try:
            original["Misc Codes"]["A28"] = "00012"
            original_path = self.root / "original-floral.xlsx"
            original.save(original_path)
            expected = [(original["Misc Codes"][f"{col}{row}"].value,
                         original["Misc Codes"][f"{chr(ord(col) + 1)}{row}"].value)
                        for col in ("A", "C") for row in range(28, 32)]
            original.close()
            original = load_workbook(original_path)
            soil = original["Soil"]
            soil["L1"] = "Floral Codes"
            soil.merge_cells("L1:M1")
            for index, (code, name) in enumerate(expected, 3):
                soil[f"L{index}"], soil[f"M{index}"] = code, name
            for col, width in (("K", 2), ("L", 10), ("M", 30)):
                soil.column_dimensions[col].width = width
            soil.print_area = "B1:M66"
            misc = original["Misc Codes"]
            misc["B27"] = None
            for col in ("A", "B", "C", "D"):
                for row in range(28, 32):
                    misc[f"{col}{row}"] = None
            legacy = self.root / "legacy-floral.xlsx"
            original.save(legacy)
        finally:
            original.close()
        for path in (original_path, legacy):
            with self.subTest(path=path):
                restored = self.root / (path.stem + "-restored.xlsx")
                entries = restore_floral_codes(path, restored)
                self.assertEqual([(entry["code"], entry["name"]) for entry in entries], expected)
                before, after = load_workbook(path), load_workbook(restored)
                try:
                    self.assertEqual(after["Misc Codes"]["B27"].value, "Floral Codes")
                    self.assertEqual(after["Misc Codes"]["A28"].value, "00012")
                    self.assertEqual(after["Misc Codes"]["A28"].data_type, "s")
                    self.assertEqual(str(after["Soil"].print_area), "'Soil'!$B$1:$J$66")
                    self.assertIsNone(after["Soil"]["L1"].value)
                    self.assertEqual([sha(image._data()) for image in before["Soil"]._images],
                                     [sha(image._data()) for image in after["Soil"]._images])
                    for entry in entries:
                        self.assertEqual(after["Misc Codes"][entry["misc_code"]].value, entry["code"])
                        self.assertEqual(after["Misc Codes"][entry["misc_name"]].value, entry["name"])
                finally:
                    before.close()
                    after.close()
                again = self.root / (path.stem + "-again.xlsx")
                restore_floral_codes(restored, again)
                self.assertEqual(restored.read_bytes(), again.read_bytes())

    def test_floral_codes_refuse_conflicting_locations(self):
        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        workbook = load_workbook(source)
        workbook["Soil"]["L1"] = "Floral Codes"
        edited = self.root / "conflicting-floral.xlsx"
        workbook.save(edited)
        workbook.close()
        with self.assertRaisesRegex(ValueError, "both"):
            restore_floral_codes(edited, self.root / "rejected-floral.xlsx")

    def test_uploaded_conversions_image_replaces_charts_and_is_preserved(self):
        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        image = (Path(__file__).resolve().parents[1] / CONVERSIONS_SOURCE).read_bytes()
        self.assertEqual(sha(image), CONVERSIONS_SOURCE_SHA256)
        self.assertIn("$4.99", CONVERSIONS_ACCURACY_REVIEW)
        self.assertIn("$11.00/kg", CONVERSIONS_ACCURACY_REVIEW)
        result = self.root / "conversions-image.xlsx"
        theme_conversions_sheet(source, result, image)
        before = load_workbook(source)
        after = load_workbook(result)
        try:
            original, themed = before["Conversions"], after["Conversions"]
            self.assertEqual(len(themed._images), 1)
            self.assertEqual(themed._images[0]._data(), image)
            self.assertEqual(themed.sheet_properties.codeName, "UpdatedConversionsImage")
            self.assertEqual((themed.page_setup.orientation, str(themed.page_setup.paperSize)), ("portrait", "1"))
            self.assertTrue(themed.sheet_properties.pageSetUpPr.fitToPage)
            self.assertFalse(any(cell.value is not None for row in themed for cell in row))
            self.assertEqual(len(original._images), 3)
        finally:
            before.close()
            after.close()
        with zipfile.ZipFile(source) as before, zipfile.ZipFile(result) as after:
            part, _ = sheet_info(before, "Conversions")
            for name in before.namelist():
                if name not in (part, "xl/worksheets/_rels/sheet2.xml.rels",
                                "xl/styles.xml", "xl/workbook.xml", "[Content_Types].xml"):
                    self.assertEqual(before.read(name), after.read(name), name)
        again = self.root / "conversions-repeated.xlsx"
        theme_conversions_sheet(result, again, image)
        self.assertEqual(sha(result.read_bytes()), sha(again.read_bytes()))

        old_draft = self.root / "old-conversions-draft.xlsx"
        donor = Workbook()
        donor.active.title = "Conversions"
        donor.active.sheet_properties.codeName = "EditableConversions"
        donor.active["A1"] = "Previous editable draft"
        donor_bytes = io.BytesIO()
        donor.save(donor_bytes)
        donor.close()
        replace_sheet(source, donor_bytes.getvalue(), old_draft, "Conversions")
        upgraded = self.root / "upgraded-conversions-draft.xlsx"
        theme_conversions_sheet(old_draft, upgraded, image)
        upgraded_workbook = load_workbook(upgraded)
        try:
            self.assertEqual(upgraded_workbook["Conversions"]._images[0]._data(), image)
        finally:
            upgraded_workbook.close()

    def test_conversions_rejects_unexpected_owner_content(self):
        source = Path(__file__).resolve().parents[1] / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        changed = self.root / "conversions-extra.xlsx"
        workbook = load_workbook(source)
        workbook["Conversions"]["A55"] = "Do not discard"
        workbook.save(changed)
        workbook.close()
        image = (Path(__file__).resolve().parents[1] / CONVERSIONS_SOURCE).read_bytes()
        with self.assertRaisesRegex(ValueError, "additional cell content"):
            theme_conversions_sheet(changed, self.root / "rejected.xlsx", image)

    def test_missing_stretch_source_fails_without_falling_back_to_original(self):
        (self.root / STRETCH_SOURCE).unlink()
        with self.assertRaises(FileNotFoundError):
            generate(self.root, self.binder, self.output)

    @unittest.skipUnless(shutil.which("soffice"), "LibreOffice Calc is required")
    def test_stretch_image_prints_in_full_on_one_portrait_page(self):
        import pymupdf

        source = (self.root / STRETCH_SOURCE).read_bytes()
        workbook = self.root / "stretch.xlsx"
        workbook.write_bytes(stretch_sheet(source))
        subprocess.run(["soffice", "-env:UserInstallation=" + (self.root / "lo").as_uri(),
                        "--headless", "--convert-to", "pdf", "--outdir", str(self.root),
                        str(workbook)], check=True, capture_output=True, timeout=120)
        with pymupdf.open(self.root / "stretch.pdf") as document:
            self.assertEqual(len(document), 1)
            page = document[0]
            self.assertEqual((round(page.rect.width), round(page.rect.height)), (612, 792))
            images = page.get_image_info()
            self.assertEqual(len(images), 1)
            with Image.open(io.BytesIO(source)) as image:
                self.assertEqual((images[0]["width"], images[0]["height"]), image.size)
            bounds = pymupdf.Rect(images[0]["bbox"])
            self.assertTrue(page.rect.contains(bounds))
            self.assertGreater(bounds.width, 400)
            self.assertGreater(bounds.height, 600)

    def test_workbook_photo_copies_keep_transparency_and_do_not_change_source_bytes(self):
        opaque = io.BytesIO()
        Image.new("RGB", (40, 30), (40, 120, 200)).save(opaque, format="PNG")
        opaque_bytes = opaque.getvalue()
        opaque_copy = workbook_photo_copy(opaque_bytes)
        with Image.open(io.BytesIO(opaque_copy)) as image:
            self.assertEqual(image.format, "PNG")
            self.assertEqual(image.size, (40, 30))

        enlarged = io.BytesIO()
        Image.new("RGB", (1200, 600), (40, 120, 200)).save(enlarged, format="PNG")
        enlarged_bytes = enlarged.getvalue()
        resized_copy = workbook_photo_copy(enlarged_bytes)
        with Image.open(io.BytesIO(resized_copy)) as image:
            self.assertEqual(image.format, "PNG")
            self.assertEqual(image.size, (192, 96))
        self.assertEqual(enlarged.getvalue(), enlarged_bytes)

        transparent = io.BytesIO()
        Image.new("RGBA", (800, 400), (40, 120, 200, 0)).save(transparent, format="PNG")
        transparent_bytes = transparent.getvalue()
        transparent_copy = workbook_photo_copy(transparent_bytes)
        with Image.open(io.BytesIO(transparent_copy)) as image:
            self.assertEqual(image.format, "PNG")
            self.assertEqual(image.size, (192, 96))
            self.assertEqual(image.getchannel("A").getextrema(), (0, 0))
        self.assertEqual(opaque.getvalue(), opaque_bytes)
        self.assertEqual(transparent.getvalue(), transparent_bytes)

    def test_binder_photos_remove_only_border_connected_white_and_preserve_badges(self):
        image = Image.new("RGBA", (40, 30), "white")
        for x in range(8, 32):
            for y in range(5, 25):
                image.putpixel((x, y), (30, 140, 50, 255))
        image.putpixel((20, 15), (255, 255, 255, 255))
        image.putpixel((7, 10), (245, 245, 245, 255))
        image.putpixel((0, 0), (235, 70, 150, 255))
        source = io.BytesIO()
        image.save(source, format="PNG")
        with Image.open(io.BytesIO(workbook_photo_copy(source.getvalue()))) as result:
            self.assertEqual(result.getpixel((39, 29))[3], 0)
            self.assertEqual(result.getpixel((20, 15)), (255, 255, 255, 255))
            self.assertEqual(result.getpixel((10, 10)), (30, 140, 50, 255))
            self.assertEqual(result.getpixel((0, 0)), (235, 70, 150, 255))
            self.assertGreater(result.getpixel((7, 10))[3], 0)
            self.assertLess(result.getpixel((7, 10))[3], 255)

    @unittest.skipUnless(shutil.which("soffice"), "LibreOffice Calc is required")
    def test_transparent_produce_background_prints_as_the_green_row(self):
        import pymupdf

        image = Image.new("RGB", (40, 30), "white")
        for x in range(10, 30):
            for y in range(8, 22):
                image.putpixel((x, y), (190, 20, 20))
        source = io.BytesIO()
        image.save(source, format="PNG")
        rows = [dict(source_id=str(i), plu=str(i), name=f"APPLES {i}", duplicate_count=1,
                     display_section="A", display_group="Apples") for i in range(2)]
        content, _, _ = quick_sheet(rows, {"1": workbook_photo_copy(source.getvalue())}, True)
        path = self.root / "transparent-row.xlsx"
        path.write_bytes(content)
        subprocess.run(["soffice", "-env:UserInstallation=" + (self.root / "lo-alpha").as_uri(),
                        "--headless", "--convert-to", "pdf", "--outdir", str(self.root), str(path)],
                       check=True, capture_output=True, timeout=120)
        with pymupdf.open(path.with_suffix(".pdf")) as document:
            page = document[0]
            bounds = pymupdf.Rect(page.get_image_info()[0]["bbox"])
            pixmap = page.get_pixmap(matrix=pymupdf.Matrix(3, 3), colorspace=pymupdf.csRGB, alpha=False)
            background = pixmap.pixel(round((bounds.x0 + bounds.width * .1) * 3),
                                      round((bounds.y0 + bounds.height * .1) * 3))
            expected = tuple(int(PALE_GREEN[index:index + 2], 16) for index in (0, 2, 4))
            self.assertTrue(all(abs(a - b) <= 2 for a, b in zip(background, expected)),
                            (background, expected))

    def test_layout_has_only_exact_product_text_codes_categories_and_no_image_labels(self):
        generate(self.root, self.binder, self.output)
        with (self.output / "PRODUCE_LAYOUT.csv").open(encoding="utf-8-sig", newline="") as file:
            layout = list(csv.DictReader(file))
        workbook = load_workbook(self.output / FILENAMES["photos"])
        self.addCleanup(workbook.close)
        sheet = workbook["Produce"]
        for record in layout:
            code, name = sheet[record["code_cell"]], sheet[record["description_cell"]]
            self.assertEqual(code.value, record["plu"])
            self.assertEqual(code.data_type, "s")
            self.assertEqual(code.number_format, "@")
            self.assertEqual(name.value, record["product"])
            self.assertEqual(name.data_type, "s")
            self.assertIsNone(sheet[record["image_cell"]].value)
            if record["plu"] == "00042":
                self.assertEqual(code.fill.fgColor.rgb, "00FFC7CE")
        texts = [str(cell.value) for row in sheet for cell in row if cell.value is not None]
        self.assertNotIn("AI", texts)
        self.assertNotIn("REVIEW", texts)
        self.assertNotIn("CHECK", texts)
        self.assertNotIn("REF", texts)
        self.assertFalse(any("Source SHA" in text or "rights" in text.lower() for text in texts))
        self.assertEqual(sheet.page_setup.orientation, "landscape")
        self.assertEqual(sheet.page_setup.fitToWidth, 1)
        self.assertEqual(sheet.page_setup.fitToHeight, 0)

    def test_binder_uses_working_list_and_requires_fresh_image_for_merged_identity(self):
        first, second = self.items[1], self.items[2]
        registry_path = self.root / "data/working_list_decisions.json"
        registry = json.loads(registry_path.read_text())
        registry["decisions"] = [dict(
            plu=first["plu"], canonical_name="Apple Gala", mode="MERGE",
            source_ids=[first["source_id"], second["source_id"]],
            selected_source_ids=[first["source_id"], second["source_id"]],
            condition="Owner-approved duplicate merge")]
        registry_path.write_text(json.dumps(registry))
        source_before, binder_before = (
            (self.root / "sources/plu/PickList.xml").read_bytes(), self.binder.read_bytes())
        summary = generate(self.root, self.binder, self.output)
        self.assertEqual(summary["produce_records_included"], 6)
        self.assertEqual((self.root / "sources/plu/PickList.xml").read_bytes(), source_before)
        self.assertEqual(self.binder.read_bytes(), binder_before)
        with (self.output / "PRODUCE_LAYOUT.csv").open(
                encoding="utf-8-sig", newline="") as file:
            layout = list(csv.DictReader(file))
        matching = [row for row in layout if row["plu"] == "00042"]
        self.assertEqual(len(matching), 1)
        self.assertEqual(matching[0]["product"], "APPLE GALA")
        with (self.output / "PHOTO_MAPPING.csv").open(
                encoding="utf-8-sig", newline="") as file:
            photos = list(csv.DictReader(file))
        image = next(row for row in photos if row["product"] == "Apple Gala")
        self.assertEqual(image["image_status"], "MISSING")
        self.assertFalse(image["embedded"] == "True")
        self.assertIn("No prior image association was reused", image["details"])

    def test_mapping_requires_exact_identity_legacy_review_and_source_checksum(self):
        assignments = read_assignments(self.root / "data/picklist_assignments.json")
        rows, _ = produce_catalog(self.items, assignments)
        visuals = visual_records(self.root)
        images, report = map_images(self.root, rows, assignments, visuals)
        statuses = {row["plu"]: row["image_status"] for row in report}
        self.assertEqual(statuses["00045"], "AI_ILLUSTRATION_REFERENCE_REVIEW")
        self.assertEqual(statuses["00044"], "VISUAL_REVIEW_REQUIRED")
        self.assertEqual(statuses["00049"], "MISSING")
        ai = next(row for row in rows if row["plu"] == "00045")
        self.assertIn(ai["source_id"], images)
        normal = next(row for row in rows if row["plu"] == "00050")
        assignments[normal["source_id"]]["image_sha256"] = "changed"
        _, report = map_images(self.root, rows, assignments, visuals)
        self.assertEqual(next(row for row in report if row["plu"] == "00050")["image_status"],
                         "SOURCE_CHECKSUM_CHANGED")

    def test_explicit_working_list_approval_is_embedded_only_for_its_exact_identity(self):
        assignments = read_assignments(self.root / "data/picklist_assignments.json")
        rows, _ = produce_catalog(self.items, assignments)
        row = next(item for item in rows if item["name"] == "Apple Gala")
        assignment = assignments[row["source_id"]]
        assignment.update(
            image="data/reference.png", image_sha256=sha(self.content),
            reviewed_plu=row["plu"], reviewed_product=row["name"],
            visual_approval_status="APPROVED", photographic_origin="VERIFIED_PHOTO",
            rights_verification_status="VERIFIED",
            provenance=dict(
                creator="Test photographer", license="CC0",
                license_url="https://creativecommons.org/publicdomain/zero/1.0/",
                source_page_url="https://example.org/apple",
                original_image_url="https://example.org/apple.png",
                attribution="Test photographer / CC0", changes="Original image unchanged",
                rights_evidence="Test fixture rights evidence",
            ),
        )
        images, report = map_images(self.root, rows, assignments, visual_records(self.root))
        record = next(item for item in report if item["source_id"] == row["source_id"])
        self.assertEqual(record["image_status"], "APPROVED_WORKINGLIST_IMAGE")
        self.assertTrue(record["embedded"])
        self.assertEqual(images[row["source_id"]], self.content)
        row["name"] = "Different Apple"
        row["identity_changed"] = True
        images, report = map_images(self.root, rows, assignments, visual_records(self.root))
        record = next(item for item in report if item["source_id"] == row["source_id"])
        self.assertEqual(record["image_status"], "APPROVED_ASSIGNMENT_INVALID")
        self.assertNotIn(row["source_id"], images)

    def test_uncertain_old_mappings_are_reported_without_guessing(self):
        original = [
            dict(original_code="00042", original_description="Old name", cell="A1", heading="Apple", issue=""),
            dict(original_code="00046", original_description="Hidden Apple", cell="A2", heading="Apple", issue=""),
            dict(original_code="123", original_description="Apple Gala", cell="A3", heading="Apple", issue=""),
            dict(original_code="99999", original_description="Unknown", cell="A4", heading="Apple", issue=""),
        ]
        changes, _ = comparisons(original, self.items, [])
        self.assertEqual([row["status"] for row in changes], [
            "CODE_ACTIVE_DESCRIPTION_DIFFERS_REVIEW", "CODE_ONLY_HIDDEN_OR_DELETED",
            "POSSIBLE_CODE_CHANGE_EXACT_NAME_REVIEW", "NOT_IN_CURRENT_SOURCE_UNRESOLVED",
        ])
        self.assertEqual(changes[0]["active_match_count"], 2)

    def test_existing_ai_references_remain_without_cashier_labels_but_rejections_stay_blank(self):
        path = self.root / "data/image_review.csv"
        with path.open(newline="") as file:
            reviews = list(csv.DictReader(file))
        ai = next(row for row in reviews if row["plu"] == "00045")
        entries = [dict(idx=row["idx"], book="Fruit", plu=row["plu"], name=row["name"],
                        image="reference.png",
                        source="AI Generated" if row["idx"] == ai["idx"] else "Photo")
                   for row in reviews]
        entry_reviews = {row["idx"]: row for row in entries}
        with (self.root / "data/entries.csv").open("w", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=list(entries[0]))
            writer.writeheader()
            writer.writerows(entries)
        for verdict in ("Approximate", "Rejected"):
            with self.subTest(verdict=verdict):
                ai["verdict"] = verdict
                with path.open("w", newline="") as file:
                    legacy_reviews = [
                        dict(entry_reviews[row["idx"]], verdict=row["verdict"],
                             note=row.get("note", ""))
                        for row in reviews
                    ]
                    writer = csv.DictWriter(file, fieldnames=list(legacy_reviews[0]))
                    writer.writeheader()
                    writer.writerows(legacy_reviews)
                assignments = read_assignments(self.root / "data/picklist_assignments.json")
                rows, _ = produce_catalog(self.items, assignments)
                images, report = map_images(self.root, rows, assignments, visual_records(self.root))
                record = next(row for row in report if row["plu"] == "00045")
                self.assertTrue(record["embedded"])
                self.assertEqual(record["image_status"], "LEGACY_PREAPPROVED_IMAGE")
                self.assertEqual(record["image_approval_status"], "APPROVED_LEGACY")
                self.assertEqual(record["legacy_original_verdict"], verdict)
                self.assertEqual(record["image_source"], "AI Generated")
                self.assertEqual(record["visual_verdict"], "APPROVED_LEGACY")
                self.assertIn("pre-approved", record["details"])
                donor, coordinates, _ = quick_sheet(rows, images, True)
                workbook = load_workbook(io.BytesIO(donor))
                self.addCleanup(workbook.close)
                entry = next(row for row in coordinates if row["plu"] == "00045")
                self.assertIsNone(workbook["Produce"][entry["image_cell"]].value)

    def test_pagination_never_orphans_a_heading_or_loses_duplicate_records(self):
        rows = [dict(display_section="Apples" if i < 80 else "Melons", source_id=str(i))
                for i in range(130)]
        layout = layout_rows(rows)
        self.assertEqual([entry["record"]["source_id"] for entry in layout if entry["kind"] == "product"],
                         [str(i) for i in range(130)])
        for index, entry in enumerate(layout):
            if entry["kind"] == "heading":
                self.assertEqual(layout[index + 1]["kind"], "product")
                self.assertEqual(layout[index + 1]["block"], entry["block"])
        self.assertTrue(any(entry["block"] >= 3 for entry in layout))

    def test_produce_stripes_cover_merged_names_and_keep_duplicate_warnings(self):
        rows = [dict(display_section="A", source_id=str(i), plu=f"{i:05}",
                     name=f"Apple {i}", duplicate_count=2 if i == 1 else 1)
                for i in range(25)]
        content, coordinates, _ = quick_sheet(rows, {}, False)
        workbook = load_workbook(io.BytesIO(content))
        try:
            sheet = workbook["Produce"]
            for index in (0, 1, 2, 19, 20):
                entry = coordinates[index]
                product = sheet[entry["description_cell"]]
                expected = "00" + (PALE_GREEN if index in (1, 20) else "FFFFFF")
                self.assertEqual(product.fill.fgColor.rgb, expected)
                self.assertEqual(sheet[entry["image_cell"]].fill.fgColor.rgb, expected)
            self.assertEqual(sheet[coordinates[1]["code_cell"]].fill.fgColor.rgb, "00FFC7CE")
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                tree = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
                cells = {cell.get("r"): int(cell.get("s", 0)) for cell in
                         tree.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c")}
                for entry in coordinates:
                    product = entry["description_cell"]
                    next_column = chr(ord(product[0]) + 1) + product[1:]
                    self.assertEqual(workbook._cell_styles[cells[product]].fillId,
                                     workbook._cell_styles[cells[next_column]].fillId)
        finally:
            workbook.close()

    def test_binder_tables_use_matching_alternating_rows(self):
        root = Path(__file__).resolve().parents[1]
        bakery, bakery_coordinates, _ = read_bakery_source(root)
        from build_deli_sheet import read_source as read_deli_source
        deli, deli_coordinates = read_deli_source(root)
        for name, content, coordinates in (("Bakery", bakery, bakery_coordinates),
                                            ("Deli", deli, deli_coordinates)):
            workbook = load_workbook(io.BytesIO(theme_product_table(content, name, coordinates)))
            try:
                sheet = workbook[name]
                for row, color in ((5, "00FFFFFF"), (6, "00" + PALE_GREEN)):
                    self.assertEqual(sheet.cell(row, 1).fill.fgColor.rgb, color)
                    self.assertEqual(sheet.cell(row, 2).fill.fgColor.rgb, color)
                self.assertEqual(sheet["A3"].fill.fgColor.rgb, "00305040")
                for merged in sheet.merged_cells.ranges:
                    heading = sheet.cell(merged.min_row, merged.min_col)
                    if heading.value is not None:
                        self.assertEqual(heading.alignment.horizontal, "center")
            finally:
                workbook.close()

    def test_validation_detects_unrelated_package_changes(self):
        assignments = read_assignments(self.root / "data/picklist_assignments.json")
        rows, _ = produce_catalog(self.items, assignments)
        images, report = map_images(self.root, rows, assignments, visual_records(self.root))
        donor, coordinates, _ = quick_sheet(rows, images, False)
        output = self.root / "changed.xlsx"
        replace_produce(self.binder, donor, output)
        validate_package(self.binder, output, coordinates, images, False)
        corrupted = self.root / "corrupted.xlsx"
        with zipfile.ZipFile(output) as source, zipfile.ZipFile(corrupted, "w") as target:
            for entry in source.infolist():
                target.writestr(entry, b"<changed/>" if entry.filename == "customXml/untouched.xml"
                                else source.read(entry.filename))
        with self.assertRaisesRegex(ValueError, "unrelated workbook part"):
            validate_package(self.binder, corrupted, coordinates, images, False)

    def test_broken_xml_and_missing_source_never_produce_a_binder(self):
        (self.root / "sources/plu/PickList.xml").write_bytes(b"<broken>")
        with self.assertRaisesRegex(ValueError, "cannot parse"):
            generate(self.root, self.binder, self.output)
        self.assertEqual(list(self.output.glob("*.xlsx")), [])
        with self.assertRaisesRegex(ValueError, "missing original"):
            generate(self.root, self.root / "missing.xlsx", self.root / "other-output")

    def write_ai_manifest(self, panels):
        folder = self.root / "data/review-references"
        folder.mkdir(exist_ok=True)
        (folder / "sheet.png").write_bytes(self.content)
        manifest = dict(schema_version=1, source="data/review-references/sheet.png",
                        source_sha256=sha(self.content), source_size=[600, 500],
                        origin="Owner-supplied AI illustration", review_note="Review only", panels=panels)
        path = folder / "additional-36-ai.json"
        path.write_text(json.dumps(manifest))
        return path

    def ai_panel(self, number, **changes):
        item = self.items[number]
        return dict(plu=item["plu"], product=item["name"], box=[10, 10, 190, 160],
                    decision="REVIEW_REFERENCE", **changes)

    def test_new_ai_reference_replaces_blank_with_exact_crop_without_approving_registry(self):
        self.write_ai_manifest([self.ai_panel(3), self.ai_panel(8)])
        before = (self.root / "data/picklist_assignments.json").read_bytes()
        summary = generate(self.root, self.binder, self.output)
        self.assertEqual(summary["new_ai_crops_embedded"], 2)
        self.assertEqual(summary["embedded_references"], 7)
        self.assertEqual(summary["production_publication"], "BLOCKED")
        self.assertEqual((self.root / "data/picklist_assignments.json").read_bytes(), before)
        with (self.output / "PHOTO_MAPPING.csv").open(encoding="utf-8-sig", newline="") as file:
            audit = list(csv.DictReader(file))
        reference = next(row for row in audit if row["plu"] == "00044")
        self.assertEqual(reference["image_status"], "SUPPLIED_AI_CROP_REVIEW")
        self.assertEqual(reference["rights_status"], "NOT_VERIFIED")
        self.assertEqual(reference["photograph_status"], "AI_ILLUSTRATION_NOT_PHOTOGRAPH")
        self.assertEqual((reference["width"], reference["height"]), ("180", "150"))
        organic = next(row for row in audit if row["plu"] == "00049")
        self.assertEqual(organic["organic_badge"], "True")
        self.assertNotEqual(organic["embedded_sha256"], organic["original_sha256"])

    def test_ai_crops_reject_stale_identity_ambiguous_duplicates_and_rejected_panels(self):
        path = self.write_ai_manifest([self.ai_panel(3)])
        rows = [self.items[3], dict(self.items[3], source_id="another-occurrence")]
        refs, audit = load_references(self.root, rows)
        self.assertEqual(refs, {})
        self.assertEqual(audit[0]["status"], "AMBIGUOUS_DUPLICATE_IDENTITY")
        manifest = json.loads(path.read_text())
        manifest["panels"][0]["source_id"] = self.items[3]["source_id"]
        path.write_text(json.dumps(manifest))
        refs, _ = load_references(self.root, rows)
        self.assertEqual(set(refs), {self.items[3]["source_id"]})
        refs, audit = load_references(self.root, [dict(self.items[3], name="Changed product")])
        self.assertEqual(refs, {})
        self.assertEqual(audit[0]["status"], "WORKING_IDENTITY_REVIEW_REQUIRED")
        manifest["panels"][0]["decision"] = "REJECTED"
        path.write_text(json.dumps(manifest))
        refs, audit = load_references(self.root, self.items)
        self.assertEqual(refs, {})
        self.assertEqual(audit[0]["status"], "REJECTED")
        assignments = read_assignments(self.root / "data/picklist_assignments.json")
        rows, _ = produce_catalog(self.items, assignments)
        _, report = map_images(self.root, rows, assignments, visual_records(self.root))
        record = next(row for row in report if row["plu"] == self.items[3]["plu"])
        self.assertIn("New AI panel excluded:", record["details"])

    def test_working_identity_change_requires_fresh_ai_reference_review(self):
        self.write_ai_manifest([self.ai_panel(3)])
        assignments = read_assignments(self.root / "data/picklist_assignments.json")
        rows, _ = produce_catalog(self.items, assignments)
        original = next(row for row in rows if row["source_position"] == 4)
        changed = dict(original, name="WorkingList canonical name",
                       source_id="working-identity-123",
                       source_ids=[original["source_id"]], identity_changed=True)
        references, audit = load_references(self.root, [changed])
        self.assertEqual(references, {})
        self.assertEqual(audit[0]["status"], "WORKING_IDENTITY_REVIEW_REQUIRED")

    def test_small_ai_crops_keep_proportional_badge_without_changing_legacy_renderer(self):
        panel = self.ai_panel(8)
        panel["box"] = [0, 0, 60, 50]
        self.write_ai_manifest([panel])
        assignments = read_assignments(self.root / "data/picklist_assignments.json")
        rows, _ = produce_catalog(self.items, assignments)
        references, _ = load_references(self.root, rows)
        images, _ = map_images(self.root, rows, assignments, visual_records(self.root))
        source_id = self.items[8]["source_id"]
        content = references[source_id]["content"]
        self.assertEqual(images[source_id], render_watermark(content, minimum_side=1))
        self.assertNotEqual(images[source_id], render_watermark(content))
        self.assertEqual(Image.open(io.BytesIO(images[source_id])).size, (60, 50))

    def test_supplied_image_does_not_infer_ai_origin_or_production_approval(self):
        path = self.write_ai_manifest([self.ai_panel(3), self.ai_panel(8)])
        manifest = json.loads(path.read_text())
        manifest.update(image_kind="UNVERIFIED_IMAGE", origin="Owner-supplied image; origin unknown")
        path.write_text(json.dumps(manifest))
        before = (self.root / "data/picklist_assignments.json").read_bytes()
        summary = generate(self.root, self.binder, self.output)
        self.assertEqual(summary["production_publication"], "BLOCKED")
        self.assertEqual(summary["new_ai_crops_embedded"], 0)
        self.assertEqual((self.root / "data/picklist_assignments.json").read_bytes(), before)
        with (self.output / "PHOTO_MAPPING.csv").open(encoding="utf-8-sig", newline="") as file:
            audit = list(csv.DictReader(file))
        for plu in ("00044", "00049"):
            record = next(row for row in audit if row["plu"] == plu)
            self.assertEqual(record["image_status"], "SUPPLIED_IMAGE_REVIEW")
            self.assertEqual(record["image_source"], "Owner Supplied")
            self.assertEqual(record["photograph_status"], "NOT_VERIFIED")
            self.assertEqual(record["rights_status"], "NOT_VERIFIED")
            self.assertNotIn("AI illustration", record["details"])
            self.assertNotIn("below 500", record["details"])
        organic = next(row for row in audit if row["plu"] == "00049")
        self.assertEqual(organic["organic_badge"], "True")
        self.assertNotEqual(organic["embedded_sha256"], organic["original_sha256"])
        manifest["image_kind"] = "APPROVED"
        path.write_text(json.dumps(manifest))
        with self.assertRaisesRegex(ValueError, "unsupported review-reference image kind"):
            load_references(self.root, self.items)

    def test_fresh_working_reference_requires_exact_snapshot_and_original_identities(self):
        original = self.items[3]
        changed = dict(original, name="Reviewed canonical name", source_id="working-id",
                       source_ids=[original["source_id"]], record_sha256="reviewed-record",
                       identity_changed=True)
        snapshot = {key: changed[key] for key in ("source_id", "source_ids", "record_sha256")}
        panel = self.ai_panel(3, working_identity=snapshot)
        panel["product"] = changed["name"]
        path = self.write_ai_manifest([panel])
        references, _ = load_references(self.root, [changed])
        self.assertEqual(set(references), {changed["source_id"]})
        for values in (dict(name="Changed again"), dict(source_id="new-working-id"),
                       dict(source_ids=["other-source"]), dict(record_sha256="new-record"),
                       dict(source_record_changed=True)):
            with self.subTest(values=values):
                references, audit = load_references(self.root, [dict(changed, **values)])
                self.assertEqual(references, {})
                self.assertEqual(audit[0]["status"], "WORKING_IDENTITY_REVIEW_REQUIRED")
        references, audit = load_references(self.root, [changed, dict(changed)])
        self.assertEqual(references, {})
        self.assertEqual(audit[0]["status"], "AMBIGUOUS_DUPLICATE_IDENTITY")
        manifest = json.loads(path.read_text())
        manifest["panels"][0]["decision"] = "REJECTED"
        path.write_text(json.dumps(manifest))
        references, audit = load_references(self.root, [changed])
        self.assertEqual(references, {})
        self.assertEqual(audit[0]["status"], "REJECTED")

    def test_ai_crop_checksums_bounds_and_exact_text_are_enforced(self):
        path = self.write_ai_manifest([self.ai_panel(3)])
        original = json.loads(path.read_text())
        for field, value, message in (
                ("source_sha256", "changed", "checksum"),
                ("source_size", [1, 1], "dimensions")):
            with self.subTest(field=field):
                manifest = dict(original, **{field: value})
                path.write_text(json.dumps(manifest))
                with self.assertRaisesRegex(ValueError, message):
                    load_references(self.root, self.items)
        for box in ([-1, 0, 100, 100], [0, 0, 601, 100], [100, 0, 0, 100], [0, 0, 1.5, 2]):
            manifest = json.loads(json.dumps(original))
            manifest["panels"][0]["box"] = box
            path.write_text(json.dumps(manifest))
            with self.assertRaisesRegex(ValueError, "crop box"):
                load_references(self.root, self.items)
        path.write_text(json.dumps(original))
        refs, audit = load_references(self.root, [dict(self.items[3], plu="44")])
        self.assertEqual(refs, {})
        self.assertEqual(audit[0]["status"], "NO_ACTIVE_EXACT_IDENTITY")

    def test_review_booklets_include_exact_entries_and_images_without_cashier_metadata(self):
        assignments = read_assignments(self.root / "data/picklist_assignments.json")
        rows, _ = produce_catalog(self.items, assignments)
        self.write_ai_manifest([self.ai_panel(3)])
        images, _ = map_images(self.root, rows, assignments, visual_records(self.root))
        non_produce = dict(self.items[7], book="Non-produce", section="Misc",
                           organic=False, duplicate_count=1)
        vegetable = dict(rows[0], source_id="vegetable-fixture", book="Vegetable", plu="00071",
                         name="Vegetable", section="Vegetables", organic=False, duplicate_count=1)
        results = build_review_booklets(
            self.root / "booklets", rows + [non_produce, vegetable], images, assignments, self.digest)
        self.assertEqual(results["Fruit"]["records"], len(rows))
        self.assertEqual(results["Non-produce"]["images"], 0)
        self.assertTrue(all(result["cashier_image_labels_absent"] for result in results.values()))
        self.assertEqual(len(list((self.root / "booklets").glob("*.docx"))), 3)
        self.assertEqual(len(list((self.root / "booklets").glob("*.xlsx"))), 3)

    def test_standalone_review_reports_available_binder_and_rejects_changed_original(self):
        registry_path = self.root / "data/picklist_assignments.json"
        registry = json.loads(registry_path.read_text())
        registry["assignments"][1].update(book="Vegetable", section="Vegetables")
        registry_path.write_text(json.dumps(registry))
        reports = self.root / "outputs/review-booklets"
        reports.mkdir(parents=True)
        for name in ("DUPLICATE_PLU_CLASSIFICATIONS.csv", "DUPLICATE_PLU_CLASSIFICATION_REVIEW.md"):
            (reports / name).write_text("Fixture decision report\n")
        missing = build_working_list_review(self.root, self.output)
        self.assertEqual(missing["complete_binder_regeneration"]["status"], "BLOCKED_MISSING_ORIGINAL")
        binder_path = self.root / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"
        binder_path.parent.mkdir(exist_ok=True)
        binder_path.write_bytes(self.binder.read_bytes())
        with patch("review_booklets.ORIGINAL_BINDER_SHA256", sha(self.binder.read_bytes())):
            output = self.root / "available-review"
            available = build_working_list_review(self.root, output)
            self.assertEqual(available["complete_binder_regeneration"]["status"], "ORIGINAL_AVAILABLE_NOT_RUN")
            self.assertIn("available and checksum-verified", (output / "BUILD_NOTES.txt").read_text())
            self.assertNotIn("workbook is missing", (output / "BUILD_NOTES.txt").read_text())
            self.assertIn(str(binder_path.relative_to(self.root)), (output / "SHA256SUMS.txt").read_text())
            binder_path.write_bytes(b"Changed original")
            with self.assertRaisesRegex(ValueError, "binder checksum differs"):
                build_working_list_review(self.root, self.root / "changed-review")
        self.assertFalse((self.root / "changed-review").exists())

    def test_supplier_photos_override_ai_with_permission_and_checksum_evidence_only(self):
        self.write_ai_manifest([self.ai_panel(3)])
        folder = self.root / "data/review-references/buck-brand"
        folder.mkdir()
        (folder / "test.png").write_bytes(self.content)
        (folder / "IMAGE_INVENTORY.json").write_text(json.dumps([
            dict(filename="test.png", width=600, height=500, sha256=sha(self.content))]))
        path = folder.parent / "buck-brand.json"
        manifest = dict(schema_version=1, source_url="https://example.test/source.pdf",
                        pdf_sha256="source-digest", attribution="Supplier copyright",
                        rights_status="OWNER_CONFIRMED_PERMISSION",
                        rights_evidence="Owner confirmed permission for binders and repository",
                        note="Review only", mappings=[
                            dict(plu=self.items[3]["plu"], product=self.items[3]["name"],
                                 label="Exact product", asset="test.png")])
        path.write_text(json.dumps(manifest))
        assignments = read_assignments(self.root / "data/picklist_assignments.json")
        before = (self.root / "data/picklist_assignments.json").read_bytes()
        rows, _ = produce_catalog(self.items, assignments)
        images, report = map_images(self.root, rows, assignments, visual_records(self.root))
        record = next(row for row in report if row["plu"] == self.items[3]["plu"])
        self.assertEqual(record["image_status"], "SUPPLIER_REFERENCE_REVIEW")
        self.assertEqual(record["rights_status"], "OWNER_CONFIRMED_PERMISSION")
        self.assertEqual(record["photograph_status"], "SUPPLIER_PUBLISHED_PHOTO")
        self.assertEqual(images[self.items[3]["source_id"]], self.content)
        self.assertEqual(before, (self.root / "data/picklist_assignments.json").read_bytes())
        refs, _ = load_references(self.root, [dict(self.items[3], name="New identity")])
        self.assertNotIn(self.items[3]["source_id"], refs)
        _, audit = load_references(self.root, [dict(self.items[3], name="New identity")])
        self.assertEqual(audit[0]["status"], "WORKING_IDENTITY_REVIEW_REQUIRED")
        manifest["rights_status"] = "UNKNOWN"
        path.write_text(json.dumps(manifest))
        with self.assertRaisesRegex(ValueError, "permission"):
            load_references(self.root, rows)
        manifest["rights_status"] = "OWNER_CONFIRMED_PERMISSION"
        path.write_text(json.dumps(manifest))
        jpeg = io.BytesIO()
        Image.open(io.BytesIO(self.content)).save(jpeg, format="JPEG")
        # PDF-extracted JPEGs may lack the JFIF header required by python-docx.
        original = jpeg.getvalue()
        segment_end = 4 + int.from_bytes(original[4:6], "big")
        content = original[:2] + original[segment_end:]
        (folder / "test.png").write_bytes(content)
        (folder / "IMAGE_INVENTORY.json").write_text(json.dumps([
            dict(filename="test.png", width=600, height=500, sha256=sha(content))]))
        refs, _ = load_references(self.root, rows)
        from docx.image.image import Image as DocxImage
        decoded = refs[self.items[3]["source_id"]]["content"]
        self.assertEqual(DocxImage.from_file(io.BytesIO(decoded)).px_width, 600)
        self.assertEqual(Image.open(io.BytesIO(decoded)).tobytes(),
                         Image.open(io.BytesIO(content)).tobytes())
        (folder / "test.png").write_bytes(b"changed")
        with self.assertRaisesRegex(ValueError, "checksum"):
            load_references(self.root, rows)


if __name__ == "__main__":
    unittest.main()
