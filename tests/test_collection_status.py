import unittest

from src.collection_status import build_collection_total_status_text


class CollectionStatusTests(unittest.TestCase):
    def test_in_progress_entry_is_not_counted_as_completed(self):
        text = build_collection_total_status_text(
            total_raw=372,
            completed=4,
            total_seeds=231,
            current_seed=5,
            current_label="Artificial Shrubs",
            warning=3,
        )

        self.assertIn("已完成小类入口：**4/231**", text)
        self.assertIn("正在处理：**5/231**", text)
        self.assertNotIn("已完成小类入口：**5/231**", text)
        self.assertIn("数量偏少警告：**3 个**", text)

    def test_completed_status_has_no_current_entry(self):
        text = build_collection_total_status_text(
            total_raw=96,
            completed=1,
            total_seeds=1,
        )

        self.assertIn("已完成小类入口：**1/1**", text)
        self.assertNotIn("正在处理", text)


if __name__ == "__main__":
    unittest.main()
