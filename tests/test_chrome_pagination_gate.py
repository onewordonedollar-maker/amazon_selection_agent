import unittest

from src.chrome_cdp import (
    ChromeRefreshResult,
    EMPTY_NEW_RELEASES_MESSAGE,
    SELLERSPRITE_INCOMPLETE_MESSAGE,
    should_advance_to_next_page,
)


def refresh_result(
    ok: bool,
    message: str,
    source_url: str = "https://www.amazon.com/gp/bestsellers/pet-supplies/3024125011",
    next_page_url: str = "https://www.amazon.com/Best-Sellers/zgbs/pet-supplies/3024125011?pg=2",
):
    return ChromeRefreshResult(
        ok=ok,
        product_count=46,
        hydrated_count=46 if ok else 28,
        image_count=46,
        source_url=source_url,
        message=message,
        next_page_url=next_page_url,
    )


class ChromePaginationGateTests(unittest.TestCase):
    def test_incomplete_sellersprite_page_must_not_advance(self):
        self.assertFalse(
            should_advance_to_next_page(
                refresh_result(False, SELLERSPRITE_INCOMPLETE_MESSAGE),
                page=1,
                page_count=2,
            )
        )

    def test_complete_first_page_can_advance(self):
        self.assertTrue(
            should_advance_to_next_page(
                refresh_result(True, "刷新完成。"),
                page=1,
                page_count=2,
            )
        )

    def test_empty_new_releases_does_not_advance(self):
        self.assertFalse(
            should_advance_to_next_page(
                refresh_result(True, EMPTY_NEW_RELEASES_MESSAGE),
                page=1,
                page_count=2,
            )
        )

    def test_new_releases_complete_single_page_without_next_does_not_advance(self):
        self.assertFalse(
            should_advance_to_next_page(
                refresh_result(
                    True,
                    "刷新完成。",
                    source_url="https://www.amazon.com/gp/new-releases/pet-supplies/2975329011",
                    next_page_url="",
                ),
                page=1,
                page_count=2,
            )
        )

    def test_last_requested_page_does_not_advance(self):
        self.assertFalse(
            should_advance_to_next_page(
                refresh_result(True, "刷新完成。"),
                page=2,
                page_count=2,
            )
        )


if __name__ == "__main__":
    unittest.main()
