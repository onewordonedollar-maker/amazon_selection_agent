import unittest

from src.chrome_cdp import rank_category_identity, validate_rank_category_page


EXPECTED = "https://www.amazon.com/gp/new-releases/home-garden/17659096011"


class RankPageValidationTests(unittest.TestCase):
    def test_identity_ignores_ref_and_query(self):
        self.assertEqual(
            rank_category_identity(
                "https://www.amazon.com/gp/new-releases/home-garden/17659096011/ref=zg_nr_pg_2?pg=2"
            ),
            ("home-garden", "17659096011"),
        )

    def test_matching_category_is_accepted(self):
        valid, message = validate_rank_category_page(
            EXPECTED,
            {
                "url": EXPECTED + "/ref=zg_nr_pg_1",
                "selectedText": "Air Fryers",
                "unavailableText": "",
            },
        )
        self.assertTrue(valid)
        self.assertEqual(message, "")

    def test_same_node_with_amazon_department_alias_is_accepted(self):
        valid, message = validate_rank_category_page(
            EXPECTED,
            {
                "url": (
                    "https://www.amazon.com/gp/new-releases/kitchen/"
                    "17659096011/ref=zg_bsnr_pg_2_kitchen?pg=2"
                ),
                "selectedText": "Air Fryers",
                "unavailableText": "",
            },
        )
        self.assertTrue(valid)
        self.assertEqual(message, "")

    def test_same_node_on_different_ranking_type_is_rejected(self):
        valid, message = validate_rank_category_page(
            EXPECTED,
            {
                "url": "https://www.amazon.com/gp/bestsellers/home-garden/17659096011",
                "selectedText": "Air Fryers",
                "unavailableText": "",
            },
        )
        self.assertFalse(valid)
        self.assertIn("bestsellers", message)

    def test_any_department_fallback_is_rejected(self):
        valid, message = validate_rank_category_page(
            EXPECTED,
            {
                "url": EXPECTED,
                "selectedText": "Any Department\n(Current)",
                "unavailableText": "",
            },
        )
        self.assertFalse(valid)
        self.assertIn("Any Department", message)

    def test_node_redirect_is_rejected(self):
        valid, message = validate_rank_category_page(
            EXPECTED,
            {
                "url": "https://www.amazon.com/gp/new-releases/home-garden/12345678901",
                "selectedText": "Other category",
                "unavailableText": "",
            },
        )
        self.assertFalse(valid)
        self.assertIn("页面跳离目标类目", message)

    def test_unavailable_page_is_rejected(self):
        valid, message = validate_rank_category_page(
            EXPECTED,
            {
                "url": EXPECTED,
                "selectedText": "",
                "unavailableText": "page not found",
            },
        )
        self.assertFalse(valid)
        self.assertIn("page not found", message)


if __name__ == "__main__":
    unittest.main()
