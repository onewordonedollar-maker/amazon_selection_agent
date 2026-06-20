import tempfile
import unittest
from pathlib import Path


class ProductAnnotationTests(unittest.TestCase):
    def test_favorites_are_deduped_by_asin_and_keep_batch_sources(self):
        try:
            from src.product_annotations import load_favorites, save_favorites, upsert_favorite
        except ImportError as exc:
            self.fail(f"product annotation helpers are missing: {exc}")

        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "favorites.json"
            favorites = {}
            product = {"asin": "B001TEST", "title": "First title", "category_path": "Home > Bath"}

            upsert_favorite(favorites, product, source_label="batch one", source_url="https://example.com/one", now="2026-06-20T10:00:00")
            upsert_favorite(favorites, {**product, "title": "Updated title"}, source_label="batch two", source_url="https://example.com/two", now="2026-06-20T11:00:00")
            save_favorites(path, favorites)

            loaded = load_favorites(path)
            self.assertEqual(["B001TEST"], list(loaded))
            favorite = loaded["B001TEST"]
            self.assertEqual("Updated title", favorite["product"]["title"])
            self.assertEqual("2026-06-20T10:00:00", favorite["favorited_at"])
            self.assertEqual(2, len(favorite["sources"]))
            self.assertEqual("batch two", favorite["sources"][-1]["label"])

    def test_notes_are_persisted_by_asin_and_blank_note_removes_entry(self):
        try:
            from src.product_annotations import load_notes, save_notes, set_product_note
        except ImportError as exc:
            self.fail(f"product annotation helpers are missing: {exc}")

        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "notes.json"
            notes = {}

            set_product_note(notes, "B001TEST", "Ask supplier for MOQ")
            save_notes(path, notes)
            self.assertEqual({"B001TEST": "Ask supplier for MOQ"}, load_notes(path))

            set_product_note(notes, "B001TEST", "   ")
            save_notes(path, notes)
            self.assertEqual({}, load_notes(path))

