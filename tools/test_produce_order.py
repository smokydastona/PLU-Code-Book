import unittest
import io
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pymupdf
from openpyxl import load_workbook
from PIL import Image

from build_till_binders import BLOCKS, PAGE_ROWS, layout_rows, normalize_printed_produce_text, quick_sheet
from produce_order import ASIAN_SECTION, display_unique_rows, order_produce, product_key
from text_fit import text_points


def row(source_id, plu, name, counterparts=""):
    return dict(source_id=source_id, source_ids=["original-" + source_id], plu=plu,
                name=name, organic="organic" in name.lower(),
                counterpart_source_ids=counterparts)


class ProduceOrderTests(unittest.TestCase):
    def test_orders_by_main_item_and_variety_across_categories(self):
        rows = [
            row("1", "3", "Organic Royal Gala Apples"),
            row("2", "2", "Apples Gala Small"),
            row("3", "1", "Apples Ambrosia"),
            row("4", "4", "Apricots"),
            row("5", "5", "Avocado Large Hass"),
            row("6", "6", "Banana Baby"),
            row("7", "7", "Bananas"),
            row("8", "8", "Organic Bananas"),
        ]
        ordered = order_produce(rows)
        self.assertEqual([r["source_id"] for r in ordered], ["3", "2", "1", "4", "5", "7", "6", "8"])
        self.assertEqual([r["display_section"] for r in ordered], ["A"] * 5 + ["B"] * 3)
        self.assertEqual(ordered[0]["display_name"], "APPLES AMBROSIA")
        self.assertEqual(rows[0]["name"], "Organic Royal Gala Apples")

    def test_subgroups_by_type_with_asian_produce_last(self):
        rows = [row("1", "1", "Bok Choy"), row("2", "2", "Herbs Basil"),
                row("3", "3", "Eggplant Italian"), row("4", "4", "Apples Fuji"),
                row("5", "5", "Mint"), row("6", "6", "Apples Gala"),
                row("7", "7", "Zucchini"), row("8", "8", "Napa Cabbage")]
        ordered = order_produce(rows)
        self.assertEqual([(r["display_section"], r["display_group"], r["name"]) for r in ordered], [
            ("A", "Apples", "Apples Fuji"), ("A", "Apples", "Apples Gala"),
            ("E", "Eggplant", "Eggplant Italian"), ("H", "Herbs", "Herbs Basil"),
            ("H", "Herbs", "Mint"), ("Z", "Zucchini", "Zucchini"),
            ("ASIAN PRODUCE", "", "Bok Choy"), ("ASIAN PRODUCE", "", "Napa Cabbage")])

    def test_organic_follows_its_conventional_variety_within_subgroup(self):
        rows = [row("1", "94131", "Organic Apples Fuji", "3"), row("2", "4133", "Apples Gala"),
                row("3", "4131", "Apples Fuji"), row("4", "94133", "Organic Apples Gala", "2")]
        self.assertEqual([r["source_id"] for r in order_produce(rows)], ["3", "1", "2", "4"])

    def test_specific_names_and_reviewed_pairing(self):
        self.assertEqual(product_key("Melon Cantaloupe Large")[0], "CANTALOUPE")
        self.assertEqual(product_key("Organic Oranges Mango")[0], "ORANGE")
        self.assertEqual(product_key("Pepper Sweet Banana")[0], "PEPPER")
        self.assertEqual(product_key("Squash Acorn/Pepper")[0], "SQUASH")
        rows = [row("1", "01", "Lettuce Red Leaf"),
                row("2", "02", "Organic Leaf Lettuce Red", "1"),
                row("3", "03", "Lettuce Romaine")]
        self.assertEqual([r["source_id"] for r in order_produce(rows)], ["1", "2", "3"])
        self.assertFalse(product_key("Unrecognized Produce")[2])

    def test_exact_duplicates_are_suppressed_but_identity_and_conflicts_remain(self):
        rows = [row("1", "00042", "Apple Gala"), row("2", "00042", "Apple Gala"),
                row("3", "00042", "Apple Fuji"), row("4", "42", "Apple Gala")]
        displayed, duplicates = display_unique_rows(rows)
        self.assertEqual([r["source_id"] for r in displayed], ["1", "3", "4"])
        self.assertEqual(duplicates, [dict(
            source_id="2", source_ids="original-2", plu="00042", name="Apple Gala",
            displayed_source_id="1", displayed_source_ids="original-1")])

    def test_sheet_uses_letter_heading_code_first_and_uppercase(self):
        rows = [dict(r, duplicate_count=1) for r in order_produce([
            row("1", "00042", "Organic Apple Gala"),
            row("2", "42", "Apple Gala"),
            row("3", "50", "Bananas"),
        ])]
        content, coordinates, _ = quick_sheet(rows, {}, False)
        workbook = load_workbook(io.BytesIO(content))
        self.addCleanup(workbook.close)
        sheet = workbook["Produce"]
        self.assertEqual(sheet["A4"].value, "A")
        self.assertEqual(sheet["A5"].value, "Apples")
        self.assertEqual(sheet["A4"].alignment.horizontal, "center")
        self.assertEqual(sheet["A5"].alignment.horizontal, "center")
        self.assertEqual(sheet["A5"].alignment.indent, 0)
        self.assertEqual([(sheet[e["code_cell"]].value, sheet[e["description_cell"]].value)
                          for e in coordinates],
                         [("42", "APPLE GALA"), ("00042", "ORGANIC APPLE GALA"),
                          ("50", "BANANAS")])

    def test_wrapped_hyphenated_names_remain_verifiable(self):
        self.assertIn("MEYER-BUCKBRAND",
                      normalize_printed_produce_text("ORGANIC LEMONS MEYER-\nBUCKBRAND"))

    def test_continued_labels_repeat_across_columns_and_pages_without_moving_title_center(self):
        rows = [dict(r, duplicate_count=1) for r in order_produce(
            [row(str(i), str(i), f"Apples {i}") for i in range(60)] +
            [row("new" + str(i), str(i), f"Apricots {i}") for i in range(25)])]
        content, coordinates, _ = quick_sheet(rows, {}, False)
        workbook = load_workbook(io.BytesIO(content))
        self.addCleanup(workbook.close)
        sheet = workbook["Produce"]
        xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        drawingml = "http://schemas.openxmlformats.org/drawingml/2006/main"
        with zipfile.ZipFile(io.BytesIO(content)) as archive:
            drawing = ET.fromstring(archive.read("xl/drawings/produce-continuations.xml"))
        annotations = {
            (int(anchor.find(f"{{{xdr}}}from/{{{xdr}}}col").text),
             int(anchor.find(f"{{{xdr}}}from/{{{xdr}}}row").text)): anchor
            for anchor in drawing
        }
        seen, repeated = set(), []
        for entry in layout_rows(rows):
            if entry["kind"] != "heading":
                continue
            expected = entry["level"] == 2 and entry["text"] in seen
            self.assertEqual(entry["continued"], expected)
            seen.add(entry["text"])
            base = BLOCKS[entry["block"] % 3]
            number = (entry["block"] // 3) * PAGE_ROWS + 4 + entry["line"]
            title = sheet.cell(number, base)
            self.assertEqual(title.value, entry["text"])
            self.assertEqual(title.alignment.horizontal, "center")
            self.assertIsNone(sheet.cell(number, base + 3).value)
            if expected:
                repeated.append(entry)
                annotation = annotations.pop((base + 1, number - 1))
                properties = annotation.find(f".//{{{drawingml}}}rPr")
                self.assertEqual(int(properties.get("sz")), title.font.sz * 50)
                self.assertEqual(properties.get("b"), "0")
                self.assertEqual(annotation.find(f".//{{{drawingml}}}t").text, "(continued)")
                offset = int(annotation.find(f"{{{xdr}}}from/{{{xdr}}}colOff").text) / 12700
                gap = offset - text_points(title.value, title.font.sz, bold=True) / 2
                self.assertAlmostEqual(gap, text_points(" ", title.font.sz, bold=True), places=3)
        self.assertFalse(annotations)
        self.assertTrue(any(entry["block"] >= 3 for entry in repeated))
        self.assertEqual(len(coordinates), len(rows))

    def test_native_labels_preserve_existing_photo_drawing(self):
        image = io.BytesIO()
        Image.new("RGB", (10, 10), "green").save(image, format="PNG")
        rows = [dict(r, duplicate_count=1) for r in order_produce(
            [row(str(i), str(i), f"Apples {i}") for i in range(50)])]
        content, _, _ = quick_sheet(rows, {rows[0]["source_id"]: image.getvalue()}, True)
        workbook = load_workbook(io.BytesIO(content))
        self.addCleanup(workbook.close)
        self.assertEqual(len(workbook["Produce"]._images), 1)
        with zipfile.ZipFile(io.BytesIO(content)) as archive:
            drawing = ET.fromstring(archive.read("xl/drawings/drawing1.xml"))
        xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
        self.assertEqual(len(drawing.findall(f".//{{{xdr}}}pic")), 1)
        self.assertEqual(len(drawing.findall(f".//{{{xdr}}}sp")), 2)

    @unittest.skipUnless(shutil.which("soffice"), "LibreOffice Calc is required")
    def test_printed_label_is_one_word_space_from_unchanged_centered_heading(self):
        rows = [dict(r, duplicate_count=1) for r in order_produce(
            [row(str(i), str(i), f"Apples {i}") for i in range(50)])]
        content, _, _ = quick_sheet(rows, {}, False)
        with tempfile.TemporaryDirectory(prefix="produce-continuation-") as directory:
            folder = Path(directory)
            source = folder / "Produce.xlsx"
            source.write_bytes(content)
            subprocess.run(
                ["soffice", "-env:UserInstallation=" + (folder / "profile").as_uri(),
                 "--headless", "--convert-to", "pdf", "--outdir", str(folder), str(source)],
                check=True, capture_output=True, timeout=90)
            with pymupdf.open(folder / "Produce.pdf") as document:
                spans = [span for block in document[0].get_text("dict")["blocks"] if "lines" in block
                         for line in block["lines"] for span in line["spans"]]
        letters = sorted((s["bbox"][0] + s["bbox"][2]) / 2 for s in spans if s["text"] == "A")
        titles = sorted((s for s in spans if s["text"] == "Apples"), key=lambda s: s["bbox"][0])
        labels = sorted((s for s in spans if s["text"] == "(continued)"), key=lambda s: s["bbox"][0])
        self.assertEqual(len(titles), 3)
        self.assertEqual(len(labels), 2)
        for letter_center, title in zip(letters, titles):
            self.assertAlmostEqual((title["bbox"][0] + title["bbox"][2]) / 2, letter_center, delta=.05)
        for title, label in zip(titles[1:], labels):
            gap = label["bbox"][0] - title["bbox"][2]
            self.assertGreater(gap, title["size"] * .15)
            self.assertLess(gap, title["size"] * .5)
            self.assertAlmostEqual(label["size"] / title["size"], .5, delta=.01)
            self.assertNotIn("Bold", label["font"])

    def test_special_asian_subgroup_also_marks_only_repeated_headings(self):
        rows = [dict(display_section=ASIAN_SECTION, display_group="") for _ in range(45)]
        headings = [entry for entry in layout_rows(rows) if entry["kind"] == "heading"]
        self.assertEqual([entry["continued"] for entry in headings], [False, True, True])

    def test_asian_produce_stays_on_one_page_when_it_fits(self):
        for preceding, asian_count, expected_page in ((56, 15, 1), (20, 15, 0), (56, 45, 1)):
            with self.subTest(preceding=preceding, asian_count=asian_count):
                rows = ([dict(display_section="Z", display_group="") for _ in range(preceding)] +
                        [dict(display_section=ASIAN_SECTION, display_group="") for _ in range(asian_count)])
                layout = layout_rows(rows)
                asian = [entry for entry in layout if
                         (entry["kind"] == "product" and
                          entry["record"]["display_section"] == ASIAN_SECTION)]
                self.assertEqual(len(asian), asian_count)
                self.assertEqual({entry["block"] // 3 for entry in asian}, {expected_page})
                self.assertEqual(sum(entry["kind"] == "product" for entry in layout), len(rows))

    def test_oversized_asian_produce_section_still_paginates(self):
        rows = [dict(display_section=ASIAN_SECTION, display_group="") for _ in range(60)]
        products = [entry for entry in layout_rows(rows) if entry["kind"] == "product"]
        self.assertEqual(len(products), 60)
        self.assertEqual({entry["block"] // 3 for entry in products}, {0, 1})


if __name__ == "__main__":
    unittest.main()
