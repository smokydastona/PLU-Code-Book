import io
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image

from used_image_library import (
    OWNER_CANDIDATES, absorb_used_library, check_used_library, digest, jpeg_copy,
    local_inventory, record_images,
)


class UsedImageLibraryChecks(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        original = io.BytesIO()
        Image.new("RGBA", (100, 50), (255, 0, 0, 128)).save(original, format="PNG")
        self.content = original.getvalue()

    def test_jpeg_is_white_padded_and_reused_without_encoding_or_rewriting(self):
        path, content, reused = jpeg_copy(self.root, self.content)
        self.assertFalse(reused)
        modified = path.stat().st_mtime_ns
        with Image.open(io.BytesIO(content)) as image:
            self.assertEqual((image.format, image.mode, image.size), ("JPEG", "RGB", (500, 500)))
            self.assertEqual(image.getpixel((0, 0)), (255, 255, 255))
            red, green, blue = image.getpixel((250, 250))
            self.assertGreater(red, 240)
            self.assertTrue(115 < green < 140 and 115 < blue < 140)
        with patch("used_image_library.display_image", side_effect=AssertionError("must reuse")):
            self.assertEqual(jpeg_copy(self.root, self.content), (path, content, True))
        self.assertEqual(path.stat().st_mtime_ns, modified)

    def test_changed_source_and_organic_rendering_have_distinct_copies(self):
        first, _, _ = jpeg_copy(self.root, self.content)
        changed = io.BytesIO()
        Image.new("RGB", (100, 50), "blue").save(changed, format="PNG")
        second, _, reused = jpeg_copy(self.root, changed.getvalue())
        organic, _, _ = jpeg_copy(self.root, self.content, organic=True)
        self.assertFalse(reused)
        self.assertEqual(len({first, second, organic}), 3)
        self.assertEqual(len(check_used_library(self.root)), 3)

    def test_tampered_or_incomplete_cache_is_reported_not_repaired(self):
        path, _, _ = jpeg_copy(self.root, self.content)
        path.write_bytes(b"changed")
        with self.assertRaisesRegex(ValueError, "checksum mismatch"):
            jpeg_copy(self.root, self.content)
        with self.assertRaisesRegex(ValueError, "checksum mismatch"):
            check_used_library(self.root)
        path.unlink()
        with self.assertRaisesRegex(ValueError, "incomplete"):
            jpeg_copy(self.root, self.content)

    def test_unapproved_associations_remain_separate_and_missing_images_stay_blank(self):
        record = dict(source_id="exact-identity", plu="00042", product="AI Apple",
                      image_status="SUPPLIED_AI_CROP_REVIEW", rights_status="NOT_VERIFIED")
        images = {"exact-identity": self.content}
        records = [record, dict(record, source_id="rejected-identity", image_status="REJECTED")]
        result = record_images(self.root, "WORKING_LIST_IMAGES.json", images, records)
        self.assertEqual(result, dict(images=1, reused=0))
        index = self.root / "sources/image-library/WORKING_LIST_IMAGES.json"
        associations = json.loads(index.read_text())
        self.assertEqual(len(associations), 1)
        self.assertEqual(associations[0]["plu"], "00042")
        self.assertEqual(associations[0]["image_status"], record["image_status"])
        self.assertEqual(associations[0]["rights_status"], "NOT_VERIFIED")
        self.assertEqual(associations[0]["library_approval_status"], "NOT_APPROVED")
        modified = index.stat().st_mtime_ns
        self.assertEqual(record_images(self.root, index.name, images, records)["reused"], 1)
        self.assertEqual(index.stat().st_mtime_ns, modified)
        self.assertEqual(len(check_used_library(self.root)), 1)
        associations[0]["library_sha256"] = "altered"
        index.write_text(json.dumps(associations))
        with self.assertRaisesRegex(ValueError, "association checksum mismatch"):
            check_used_library(self.root)

    def test_preapproved_owner_candidate_requires_one_exact_included_identity(self):
        source = self.root / "data/review-references/owner-alternates/4500-baby-eggplant.jpg"
        source.parent.mkdir(parents=True)
        source.write_bytes(self.content)
        source_sha = digest(self.content)
        (source.with_suffix(source.suffix + ".json")).write_text(json.dumps(dict(
            plu="4500", product="Baby Eggplant", sha256=source_sha,
            visual_approval_status="APPROVED",
            photographic_origin="OWNER_CONFIRMED_PHOTO",
            rights_status="OWNER_CONFIRMED_PERMISSION",
        )))
        path, rendered, _ = jpeg_copy(self.root, self.content)
        metadata_path = path.with_suffix(".jpg.json")
        metadata = json.loads(metadata_path.read_text())
        metadata.update(
            owner_candidate_id="owner-4500-baby-eggplant",
            candidate_plu="4500", candidate_product="Baby Eggplant",
            visual_approval_status="APPROVED",
            photographic_origin_status="OWNER_CONFIRMED_PHOTO",
            rights_status="OWNER_CONFIRMED_PERMISSION",
        )
        metadata_path.write_text(json.dumps(metadata))
        library = self.root / "sources/image-library"
        row = dict(
            candidate_id="owner-4500-baby-eggplant", plu="4500", product="Baby Eggplant",
            source_image=str(source.relative_to(self.root)), source_sha256=source_sha,
            original_source_image=str(source.relative_to(self.root)),
            original_source_sha256=source_sha, organic=False,
            library_image=str(path.relative_to(self.root)), library_sha256=digest(rendered),
            library_source_sha256=source_sha,
            visual_approval_status="APPROVED",
            photographic_origin_status="OWNER_CONFIRMED_PHOTO",
            rights_status="OWNER_CONFIRMED_PERMISSION",
            preapproval_status="APPROVED_FOR_EXACT_PLU_PRODUCT",
            library_approval_status="PREAPPROVED",
            production_status="PREAPPROVED_PENDING_PICKLIST_IDENTITY",
        )
        (library / OWNER_CANDIDATES).write_text(json.dumps([row]))
        (library / "WORKING_LIST_IMAGES.json").write_text("[]")
        working_list = self.root / "sources/plu/WorkingList.xml"
        working_list.parent.mkdir(parents=True)
        working_list.write_text("<WorkingList />")

        matching = dict(source_id="new-source-id", plu="4500", name="Baby Eggplant", included=True)
        with patch("working_list.read_working_list", return_value=([matching], "source-hash")):
            inventory = local_inventory(self.root)
        candidates = [item for item in inventory if item.get("candidate_id") == row["candidate_id"]]
        self.assertEqual(len(candidates), 1)
        self.assertEqual(candidates[0]["source_id"], "new-source-id")
        self.assertEqual(candidates[0]["tier"], "APPROVED")
        self.assertEqual(candidates[0]["library_approval_status"], "APPROVED")

        different = dict(matching, name="Eggplant Baby")
        with patch("working_list.read_working_list", return_value=([different], "source-hash")):
            self.assertFalse(any(
                item.get("candidate_id") == row["candidate_id"] for item in local_inventory(self.root)
            ))
        ambiguous = [matching, dict(matching, source_id="duplicate-source-id")]
        with patch("working_list.read_working_list", return_value=(ambiguous, "source-hash")):
            self.assertFalse(any(
                item.get("candidate_id") == row["candidate_id"] for item in local_inventory(self.root)
            ))
        row["source_identity_id"] = matching["source_id"]
        (library / OWNER_CANDIDATES).write_text(json.dumps([row]))
        with patch("working_list.read_working_list", return_value=(ambiguous, "source-hash")):
            pinned = [
                item for item in local_inventory(self.root)
                if item.get("candidate_id") == row["candidate_id"]
            ]
        self.assertEqual(len(pinned), 1)
        self.assertEqual(pinned[0]["source_id"], matching["source_id"])
        (library / "WORKING_LIST_IMAGES.json").unlink()
        with patch("working_list.read_working_list", return_value=([matching], "source-hash")):
            self.assertEqual(local_inventory(self.root)[0]["source_id"], matching["source_id"])
        self.assertEqual(len(check_used_library(self.root)), 1)

    def test_working_identity_checked_before_returning_local_images(self):
        record_images(self.root, "WORKING_LIST_IMAGES.json", {"old-id": self.content}, [
            dict(source_id="old-id", plu="4500", product="Baby Eggplant")
        ])
        working = self.root / "sources/plu/WorkingList.xml"
        working.parent.mkdir(parents=True)
        working.write_text("<WorkingList />")
        with patch("working_list.read_working_list", return_value=([], "source-hash")):
            with self.assertRaisesRegex(ValueError, "stale local working image identity"):
                local_inventory(self.root)

    def test_rejected_legacy_image_is_not_offered_even_in_old_index(self):
        record_images(self.root, "LEGACY_BOOKLET_IMAGES.json", {"legacy:1": self.content}, [
            dict(source_id="legacy:1", plu="00042", name="Apple",
                 image_status="LEGACY_PREAPPROVED_IMAGE",
                 legacy_review={"verdict": "Rejected"})
        ])
        self.assertEqual(local_inventory(self.root), [])

    def test_rejected_working_assignment_is_not_offered_even_in_old_index(self):
        record_images(self.root, "WORKING_LIST_IMAGES.json", {"identity": self.content}, [
            dict(source_id="identity", plu="00042", product="Apple",
                 library_approval_status="APPROVED",
                 assignment={"visual_approval_status": "REVOKED"})
        ])
        self.assertEqual(local_inventory(self.root), [])

    def previous_library(self):
        record_images(self.root, "WORKING_LIST_IMAGES.json", {"identity": self.content}, [
            dict(source_id="identity", plu="00042", rights_status="NOT_VERIFIED")
        ])
        current = self.root / "sources/image-library"
        previous = self.root / "sources/used-image-library"
        current.rename(previous)
        index = previous / "WORKING_LIST_IMAGES.json"
        records = json.loads(index.read_text())
        records[0]["library_image"] = "sources/used-image-library/" + Path(records[0]["library_image"]).name
        index.write_text(json.dumps(records))
        return previous, records[0]

    def test_absorption_preserves_bytes_and_review_status_and_removes_second_directory(self):
        previous, record = self.previous_library()
        filename = Path(record["library_image"]).name
        original = (previous / filename).read_bytes()
        absorb_used_library(self.root)
        self.assertFalse(previous.exists())
        library = self.root / "sources/image-library"
        self.assertEqual((library / filename).read_bytes(), original)
        association = json.loads((library / "WORKING_LIST_IMAGES.json").read_text())[0]
        self.assertEqual(association["library_image"], "sources/image-library/" + filename)
        self.assertEqual(association["rights_status"], "NOT_VERIFIED")
        self.assertEqual(len(check_used_library(self.root)), 1)
        absorb_used_library(self.root)

    def test_absorption_refuses_collisions_without_removing_previous_assets(self):
        previous, record = self.previous_library()
        library = self.root / "sources/image-library"
        library.mkdir()
        filename = Path(record["library_image"]).name
        (library / filename).write_bytes(b"unmanaged")
        with self.assertRaisesRegex(ValueError, "merge collision"):
            absorb_used_library(self.root)
        self.assertTrue((previous / filename).exists())
        self.assertEqual((library / filename).read_bytes(), b"unmanaged")


if __name__ == "__main__":
    unittest.main()
