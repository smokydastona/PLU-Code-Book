import io
import json
import tempfile
import unittest
from pathlib import Path

from PIL import Image, PngImagePlugin

from consolidate_image_library import consolidate
from library_assets import ALIASES, resolve_image
from used_image_library import check_used_library, jpeg_copy, record_images


class ConsolidationChecks(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.contents = []
        for compression in (0, 9):
            buffer = io.BytesIO()
            Image.new("RGB", (40, 40), "red").save(
                buffer, format="PNG", compress_level=compression)
            self.contents.append(buffer.getvalue())

    def test_identical_pixels_share_storage_without_merging_source_metadata(self):
        paths = [jpeg_copy(self.root, content)[0] for content in self.contents]
        self.assertNotEqual(paths[0], paths[1])
        self.assertEqual(paths[0].read_bytes(), paths[1].read_bytes())
        sidecars = {path.with_suffix(".jpg.json"): path.with_suffix(".jpg.json").read_bytes()
                    for path in paths}
        record_images(self.root, "WORKING_LIST_IMAGES.json", {
            "first": self.contents[0], "second": self.contents[1],
        }, [dict(source_id="first", plu="00042", product="Apple", library_approval_status="APPROVED"),
            dict(source_id="second", plu="42", product="Different Apple",
                 library_approval_status="NOT_APPROVED")])
        index = self.root / "sources/image-library/WORKING_LIST_IMAGES.json"
        original_index = index.read_bytes()
        result = consolidate(self.root)
        self.assertEqual((result["logical_images"], result["stored_images"]), (2, 1))
        self.assertEqual(result["removed_files"], 1)
        self.assertEqual(index.read_bytes(), original_index)
        for path, content in sidecars.items():
            self.assertEqual(path.read_bytes(), content)
        for path in paths:
            self.assertTrue(resolve_image(path).is_file())
        self.assertEqual(len(check_used_library(self.root)), 1)
        for content in self.contents:
            _, _, reused = jpeg_copy(self.root, content)
            self.assertTrue(reused)
        record_images(self.root, "WORKING_LIST_IMAGES.json", {"second": self.contents[1]},
                      [dict(source_id="second", plu="42", product="Different Apple")])
        self.assertEqual(len(check_used_library(self.root)), 1)
        self.assertEqual(consolidate(self.root)["removed_files"], 0)

    def test_tampered_canonical_image_fails_without_deleting_metadata(self):
        for content in self.contents:
            jpeg_copy(self.root, content)
        consolidate(self.root)
        library = self.root / "sources/image-library"
        before = {path.name: path.read_bytes() for path in library.glob("*.jpg.json")}
        next(library.glob("*.jpg")).write_bytes(b"tampered")
        with self.assertRaisesRegex(ValueError, "changed or missing|checksum mismatch"):
            consolidate(self.root)
        self.assertEqual(before, {path.name: path.read_bytes() for path in library.glob("*.jpg.json")})

    def test_organic_badge_is_not_merged_with_the_conventional_image(self):
        conventional, _, _ = jpeg_copy(self.root, self.contents[0])
        organic, _, _ = jpeg_copy(self.root, self.contents[0], organic=True)
        result = consolidate(self.root)
        self.assertEqual(result["stored_images"], 2)
        self.assertTrue(conventional.is_file())
        self.assertTrue(organic.is_file())

    def test_new_binder_cache_copy_is_automatically_consolidated(self):
        jpeg_copy(self.root, self.contents[0])
        consolidate(self.root)
        buffer = io.BytesIO()
        metadata = PngImagePlugin.PngInfo()
        metadata.add_text("source", "Different original, identical displayed pixels")
        Image.new("RGB", (40, 40), "red").save(buffer, format="PNG", pnginfo=metadata)
        record_images(self.root, "WORKING_LIST_IMAGES.json", {"new-id": buffer.getvalue()},
                      [dict(source_id="new-id", plu="00042", product="Apple")])
        self.assertEqual(len(list((self.root / "sources/image-library").glob("*.jpg"))), 1)
        self.assertEqual(len(check_used_library(self.root)), 1)

    def test_alias_cannot_escape_the_library(self):
        path, _, _ = jpeg_copy(self.root, self.contents[0])
        (path.parent / ALIASES).write_text(json.dumps({
            path.name: {"image": "../outside.jpg", "sha256": "0" * 64}
        }))
        with self.assertRaisesRegex(ValueError, "invalid image alias"):
            check_used_library(self.root)


if __name__ == "__main__":
    unittest.main()
