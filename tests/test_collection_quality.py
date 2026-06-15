import unittest
from types import SimpleNamespace

from src.collection_quality import (
    completed_collection_page_count,
    evaluate_sellersprite_collection_quality,
    is_collection_quality_warning,
    is_sellersprite_load_failure,
)


EMPTY_MESSAGE = "Amazon 明确显示该类目暂无热门新品。"


def result(products: int, hydrated: int, message: str = "", ok: bool = True):
    return SimpleNamespace(
        ok=ok,
        product_count=products,
        hydrated_count=hydrated,
        message=message,
    )


class CollectionQualityTests(unittest.TestCase):
    def evaluate(self, unique_count: int, results: list, list_type: str = "Best Sellers"):
        return evaluate_sellersprite_collection_quality(
            unique_count,
            results,
            empty_message=EMPTY_MESSAGE,
            expected_products_per_page=50,
            min_products_per_page=45,
            min_products_two_pages=95,
            list_type=list_type,
        )

    def test_complete_46_plus_46_is_warning_not_failure(self):
        ok, message = self.evaluate(92, [result(46, 46), result(46, 46)])

        self.assertTrue(ok)
        self.assertTrue(is_collection_quality_warning(message))
        self.assertIn("不计为失败", message)

    def test_incomplete_page_is_still_failure(self):
        ok, message = self.evaluate(82, [result(36, 36), result(46, 46)])

        self.assertFalse(ok)
        self.assertFalse(is_collection_quality_warning(message))
        self.assertIn("页面产品数偏少", message)

    def test_incomplete_plugin_fields_are_still_failure(self):
        ok, message = self.evaluate(92, [result(46, 40), result(46, 46)])

        self.assertFalse(ok)
        self.assertFalse(is_collection_quality_warning(message))
        self.assertIn("卖家精灵父体月销量字段未加载完整", message)

    def test_explicit_sellersprite_timeout_is_preserved(self):
        timeout_message = "已保存当前页面数据，但卖家精灵补充字段仍未完全加载。"
        ok, message = self.evaluate(28, [result(46, 28, timeout_message, ok=False)])

        self.assertFalse(ok)
        self.assertEqual(message, timeout_message)
        self.assertTrue(is_sellersprite_load_failure(message))

    def test_normal_two_pages_are_success(self):
        ok, message = self.evaluate(96, [result(46, 46), result(50, 50)])

        self.assertTrue(ok)
        self.assertFalse(is_collection_quality_warning(message))
        self.assertIn("两页采集正常", message)


    def test_new_releases_single_complete_page_is_success(self):
        ok, message = self.evaluate(
            23,
            [result(23, 23)],
            list_type="New Releases",
        )

        self.assertTrue(ok)
        self.assertTrue(is_collection_quality_warning(message))

    def test_new_releases_missing_next_button_after_complete_page_is_success(self):
        ok, message = self.evaluate(
            23,
            [
                result(23, 23),
                result(0, 0, "未找到可点击的下一页按钮。", ok=False),
            ],
            list_type="New Releases",
        )

        self.assertTrue(ok)
        self.assertTrue(is_collection_quality_warning(message))

    def test_new_releases_complete_smaller_second_page_is_success(self):
        ok, message = self.evaluate(
            78,
            [result(46, 46), result(32, 32)],
            list_type="New Releases",
        )

        self.assertTrue(ok)
        self.assertTrue(is_collection_quality_warning(message))

    def test_new_releases_incomplete_plugin_fields_are_failure(self):
        ok, message = self.evaluate(
            23,
            [result(23, 18)],
            list_type="New Releases",
        )

        self.assertFalse(ok)
        self.assertFalse(is_collection_quality_warning(message))

    def test_best_sellers_single_page_remains_failure(self):
        ok, message = self.evaluate(23, [result(23, 23)])

        self.assertFalse(ok)
        self.assertFalse(is_collection_quality_warning(message))

    def test_best_sellers_missing_next_button_remains_failure(self):
        ok, message = self.evaluate(
            23,
            [
                result(23, 23),
                result(0, 0, "未找到可点击的下一页按钮。", ok=False),
            ],
        )

        self.assertFalse(ok)
        self.assertFalse(is_collection_quality_warning(message))

    def test_completed_page_count_ignores_missing_next_button_diagnostic(self):
        results = [
            result(23, 23),
            result(0, 0, "未找到可点击的下一页按钮。", ok=False),
        ]

        self.assertEqual(completed_collection_page_count(results), 1)

    def test_completed_page_count_keeps_empty_category_page(self):
        results = [result(0, 0, EMPTY_MESSAGE)]

        self.assertEqual(completed_collection_page_count(results), 1)


if __name__ == "__main__":
    unittest.main()
