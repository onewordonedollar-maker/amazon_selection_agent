import unittest

from src.chrome_cdp import (
    ChromeRefreshResult,
    SELLERSPRITE_INCOMPLETE_MESSAGE,
    should_reload_incomplete_sellersprite_page,
)


def result(ok: bool, products: int, hydrated: int, message: str):
    return ChromeRefreshResult(
        ok=ok,
        product_count=products,
        hydrated_count=hydrated,
        image_count=products,
        source_url="https://www.amazon.com/gp/bestsellers/pet-supplies/3024171011",
        message=message,
    )


class SellerSpriteRecoveryTests(unittest.TestCase):
    def test_incomplete_page_gets_one_refresh_attempt(self):
        page = result(False, 46, 31, SELLERSPRITE_INCOMPLETE_MESSAGE)

        self.assertTrue(should_reload_incomplete_sellersprite_page(page, False))
        self.assertFalse(should_reload_incomplete_sellersprite_page(page, True))

    def test_complete_page_is_not_refreshed(self):
        page = result(True, 46, 46, "刷新完成。")

        self.assertFalse(should_reload_incomplete_sellersprite_page(page, False))

    def test_other_failures_are_not_refreshed_as_plugin_timeouts(self):
        page = result(False, 0, 0, "榜单入口校验失败")

        self.assertFalse(should_reload_incomplete_sellersprite_page(page, False))


if __name__ == "__main__":
    unittest.main()
