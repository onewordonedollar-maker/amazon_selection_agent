import unittest

from src.chrome_cdp import is_rank_category_url


class CategoryDiscoveryTests(unittest.TestCase):
    def test_rank_category_url_requires_node(self):
        self.assertTrue(
            is_rank_category_url(
                "https://www.amazon.com/gp/bestsellers/home-garden/17659096011"
            )
        )
        self.assertFalse(
            is_rank_category_url(
                "https://www.amazon.com/gp/bestsellers/home-garden/"
            )
        )


if __name__ == "__main__":
    unittest.main()
