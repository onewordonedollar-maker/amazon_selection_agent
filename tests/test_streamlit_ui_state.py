from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


class StreamlitUiStateTests(unittest.TestCase):
    def test_history_load_buttons_close_category_dialog_before_rerun(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")

        self.assertIn("def close_category_dialog_state", source)
        self.assertIn('key="load_history_record_button"', source)
        self.assertIn('key="load_last_raw_button"', source)
        self.assertGreaterEqual(source.count("on_click=close_category_dialog_state"), 2)

    def test_category_row_selection_uses_compact_state_model(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")

        self.assertIn("toggle_compact_category_selection", source)
        self.assertNotIn("def set_category_branch_selected", source)
        self.assertNotIn("def sync_category_select_all_state", source)
        self.assertNotIn("def sync_category_ancestors", source)


if __name__ == "__main__":
    unittest.main()
