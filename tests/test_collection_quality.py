import unittest
from types import SimpleNamespace

from src.collection_quality import (
    evaluate_sellersprite_collection_quality,
    is_collection_quality_warning,
)


EMPTY_MESSAGE = "Amazon 明确显示该类目暂无热门新品。"


def result(products: int, hydrated: int, message: str = ""):
    return SimpleNamespace(
        product_count=products,
        hydrated_count=hydrated,
        message=message,
    )


class CollectionQualityTests(unittest.TestCase):
    def evaluate(self, unique_count: int, results: list):
        return evaluate_sellersprite_collection_quality(
            unique_count,
            results,
            empty_message=EMPTY_MESSAGE,
            expected_products_per_page=50,
            min_products_per_page=45,
            min_products_two_pages=95,
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

    def test_normal_two_pages_are_success(self):
        ok, message = self.evaluate(96, [result(46, 46), result(50, 50)])

        self.assertTrue(ok)
        self.assertFalse(is_collection_quality_warning(message))
        self.assertIn("两页采集正常", message)


if __name__ == "__main__":
    unittest.main()
