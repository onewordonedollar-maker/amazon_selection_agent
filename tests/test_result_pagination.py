import unittest

from src.result_pagination import (
    PAGE_SIZE_OPTIONS,
    clamp_page,
    normalize_page_size,
    page_count,
    page_range_label,
    page_start_index,
    page_slice,
)


class ResultPaginationTests(unittest.TestCase):
    def test_page_size_options_start_with_50_then_100(self):
        self.assertEqual(PAGE_SIZE_OPTIONS[:2], [50, 100])
        self.assertEqual(normalize_page_size(None), 50)
        self.assertEqual(normalize_page_size(100), 100)
        self.assertEqual(normalize_page_size(999), 50)

    def test_page_count_has_at_least_one_page(self):
        self.assertEqual(page_count(0, 50), 1)
        self.assertEqual(page_count(1, 50), 1)
        self.assertEqual(page_count(50, 50), 1)
        self.assertEqual(page_count(51, 50), 2)

    def test_clamp_page_keeps_page_inside_available_range(self):
        self.assertEqual(clamp_page(0, 120, 50), 1)
        self.assertEqual(clamp_page(1, 120, 50), 1)
        self.assertEqual(clamp_page(3, 120, 50), 3)
        self.assertEqual(clamp_page(9, 120, 50), 3)

    def test_page_slice_returns_current_page_items(self):
        items = list(range(1, 121))

        self.assertEqual(page_slice(items, 1, 50), list(range(1, 51)))
        self.assertEqual(page_slice(items, 2, 50), list(range(51, 101)))
        self.assertEqual(page_slice(items, 3, 50), list(range(101, 121)))
        self.assertEqual(page_slice(items, 9, 50), list(range(101, 121)))

    def test_page_start_index_describes_display_number_offset(self):
        self.assertEqual(page_start_index(0, 1, 50), 0)
        self.assertEqual(page_start_index(120, 1, 50), 0)
        self.assertEqual(page_start_index(120, 2, 50), 50)
        self.assertEqual(page_start_index(120, 3, 50), 100)
        self.assertEqual(page_start_index(120, 9, 50), 100)

    def test_page_range_label_describes_visible_range(self):
        self.assertEqual(page_range_label(0, 1, 50), "显示 0 / 共 0 条")
        self.assertEqual(page_range_label(120, 1, 50), "显示 1-50 / 共 120 条")
        self.assertEqual(page_range_label(120, 3, 50), "显示 101-120 / 共 120 条")
        self.assertEqual(page_range_label(120, 9, 50), "显示 101-120 / 共 120 条")


if __name__ == "__main__":
    unittest.main()
