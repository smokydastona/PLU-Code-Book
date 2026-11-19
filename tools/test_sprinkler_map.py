import hashlib
import io
from pathlib import Path
import tempfile
import unittest

from PIL import Image
import pymupdf

from build_sprinkler_map import generate, point, rectangle, ROOT, LABEL_FILL, OUTLINE, SPRINKLER_FILL
from build_till_binders import (DIRECTORY_EMERGENCY_HIGHLIGHTS, DIRECTORY_HIGHLIGHTS,
                                DIRECTORY_INTERCOM_BOX, DIRECTORY_SOURCES)


class SprinklerMapTests(unittest.TestCase):
    def test_reference_markings_without_intercom_or_source_changes(self):
        source = ROOT / DIRECTORY_SOURCES[1]
        before = hashlib.sha256(source.read_bytes()).hexdigest()
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "map.pdf"
            generate(path)
            with pymupdf.open(path) as document:
                self.assertEqual(len(document), 1)
                self.assertEqual(document[0].rect, pymupdf.Rect(0, 0, 792, 612))
                self.assertEqual(document.metadata["title"], "Sprinkler Room Map")
                page = document[0]
                text = page.get_text()
                for expected in ("Sprinkler Room Map", "Side A", "Side B",
                                 "Sprinkler Room A", "Sprinkler Room B",
                                 "(Exterior of building)"):
                    self.assertIn(expected, text)
                image = document.extract_image(page.get_images()[0][0])
                with Image.open(io.BytesIO(image["image"])) as artwork:
                    self.assertEqual(artwork.crop(DIRECTORY_INTERCOM_BOX).getextrema(),
                                     ((255, 255),) * 3)
                divider = [drawing for drawing in page.get_drawings()
                           if any(item[0] == "l"
                                  and abs(item[1] - point(940, 190)) < 0.001
                                  and abs(item[2] - point(940, 1150)) < 0.001
                                  for item in drawing["items"])]
                self.assertEqual(len(divider), 1)
                self.assertGreaterEqual(divider[0]["width"], 3)
                self.assertEqual(divider[0]["dashes"], "[ 8 4 ] 0")
                drawings = page.get_drawings()
                labels = [drawing for drawing in drawings
                          if drawing["fill"] is not None
                          and drawing["color"] is not None
                          and all(abs(a - b) < 0.001
                                  for a, b in zip(drawing["color"], OUTLINE))]
                self.assertEqual(len(labels), 4)
                for text, fill in (("Side A", LABEL_FILL), ("Side B", LABEL_FILL),
                                   ("Sprinkler Room A", SPRINKLER_FILL),
                                   ("Sprinkler Room B", SPRINKLER_FILL)):
                    text_rect = page.search_for(text)[0]
                    matches = [drawing for drawing in labels
                               if drawing["rect"].contains((text_rect.tl + text_rect.br) / 2)]
                    self.assertEqual(len(matches), 1)
                    for actual, expected in zip(matches[0]["fill"], fill):
                        self.assertAlmostEqual(actual, expected, places=3)

                existing_labels = [rectangle(box) for box in (
                    *DIRECTORY_HIGHLIGHTS[1], *DIRECTORY_EMERGENCY_HIGHLIGHTS,
                    (580, 1226, 900, 1270), (1504, 620, 1644, 725))]
                for i, drawing in enumerate(labels):
                    area = drawing["rect"] + (-6, -6, 6, 6)
                    for other in existing_labels + [item["rect"] for item in labels[i + 1:]]:
                        self.assertFalse(area.intersects(other), (area, other))
                    self.assertFalse(area.intersects(page.search_for("Sprinkler Room Map")[0]))
                for needle in ("Emergency Intercom", "Code Adam", "Cashier 20", "Code Zero"):
                    self.assertNotIn(needle, text)
        self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), before)

    def test_changed_map_source_is_rejected_before_writing(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / DIRECTORY_SOURCES[1]
            source.parent.mkdir(parents=True)
            source.write_bytes(b"changed map")
            output = root / "map.pdf"
            with self.assertRaisesRegex(ValueError, "changed; review"):
                generate(output, root)
            self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
