import io
import unittest
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image, ImageChops

import build_weighed_sheet as guide
import prepare_raincheck_image as raincheck_prep

ROOT = Path(__file__).resolve().parents[1]
BINDER = ROOT / "sources/binders/Complete_Till_Binder_-_Nov_2025.xlsx"


class WeighedGuideTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.upload = (ROOT / guide.SOURCE).read_bytes()
        cls.themed = guide.themed_image(cls.upload)
        cls.rain_upload = (ROOT / guide.RAIN_CHECK_UPLOAD).read_bytes()
        cls.rain_original_pdf = (ROOT / guide.RAIN_CHECK_ORIGINAL_PDF).read_bytes()
        cls.rain_source = (ROOT / guide.RAIN_CHECK_SOURCE).read_bytes()
        cls.rain = guide.rain_check_image(cls.rain_source)

    def test_theme_changes_only_green_colors_and_owner_corrected_price_digit(self):
        with Image.open(io.BytesIO(self.upload)) as source, Image.open(io.BytesIO(self.themed)) as themed:
            original = source.convert("RGB")
            self.assertEqual(original.size, themed.size)
            label_difference = ImageChops.difference(original, themed)
            self.assertIsNotNone(label_difference.crop(guide.PRICE_CORRECTION_BOX).getbbox())
            label_difference.paste((0, 0, 0), guide.PRICE_CORRECTION_BOX)
            self.assertIsNone(label_difference.crop((578, 205, 990, 636)).getbbox())
            changed = 0
            for index, (before, after) in enumerate(zip(
                    original.get_flattened_data(), themed.get_flattened_data())):
                x, y = index % original.width, index // original.width
                left, top, right, bottom = guide.PRICE_CORRECTION_BOX
                if left <= x < right and top <= y < bottom:
                    continue
                red, green, blue = before
                if not (green > red + 8 and green > blue + 4):
                    self.assertEqual(before, after)
                changed += before != after
            self.assertGreater(changed, 100000)
        with self.assertRaisesRegex(ValueError, "upload changed"):
            guide.themed_image(b"unreviewed upload")

    def test_original_and_generated_templates_produce_same_two_page_guide(self):
        original_bytes = BINDER.read_bytes()
        workbook = load_workbook(BINDER)
        try:
            content = guide.workbook_bytes(workbook[guide.SHEET], self.themed, self.rain)
        finally:
            workbook.close()
        generated = load_workbook(io.BytesIO(content))
        try:
            guide.validate_sheet(generated[guide.SHEET], self.themed, self.rain)
        finally:
            generated.close()
        generated = load_workbook(io.BytesIO(content))
        try:
            repeated = guide.workbook_bytes(generated[guide.SHEET], self.themed, self.rain)
        finally:
            generated.close()
        generated = load_workbook(io.BytesIO(repeated))
        try:
            guide.validate_sheet(generated[guide.SHEET], self.themed, self.rain)
        finally:
            generated.close()
        self.assertEqual(original_bytes, BINDER.read_bytes())

    def test_changed_instructions_or_label_block_replacement(self):
        for change in ("instructions", "label"):
            with self.subTest(change=change):
                workbook = load_workbook(BINDER)
                try:
                    sheet = workbook[guide.SHEET]
                    if change == "instructions":
                        sheet["A39"] = "Changed rounding example"
                    else:
                        sheet._images.clear()
                    with self.assertRaisesRegex(ValueError, "instructions or label changed"):
                        guide.workbook_bytes(sheet, self.themed, self.rain)
                finally:
                    workbook.close()

    def test_corrected_rain_check_sources_and_page_placement_are_verified(self):
        self.assertEqual(guide.digest(self.rain_upload), guide.RAIN_CHECK_UPLOAD_SHA256)
        self.assertEqual(guide.digest(self.rain_original_pdf), guide.RAIN_CHECK_ORIGINAL_PDF_SHA256)
        self.assertEqual(guide.digest(self.rain_source), guide.RAIN_CHECK_SHA256)
        self.assertEqual(raincheck_prep.prepare_image(self.rain_upload), self.rain_source)
        with Image.open(io.BytesIO(self.rain)) as corrected:
            self.assertEqual(corrected.format, "PNG")
            self.assertEqual(corrected.size, (1024, 1536))
        with self.assertRaisesRegex(ValueError, "Corrected Rain Check image changed"):
            guide.rain_check_image(self.rain_source + b"changed")
        workbook = load_workbook(BINDER)
        try:
            content = guide.workbook_bytes(workbook[guide.SHEET], self.themed, self.rain)
        finally:
            workbook.close()
        generated = load_workbook(io.BytesIO(content))
        try:
            sheet = generated[guide.SHEET]
            sheet._images[1].anchor._from.colOff = sheet._images[0].anchor._from.colOff
            with self.assertRaisesRegex(ValueError, "two-page"):
                guide.validate_sheet(sheet, self.themed, self.rain)
        finally:
            generated.close()


if __name__ == "__main__":
    unittest.main()
