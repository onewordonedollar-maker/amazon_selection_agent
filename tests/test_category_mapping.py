import unittest

from src.category_mapping import (
    AIR_FRYERS_PARENT_PATH,
    AIR_FRYERS_PATH,
    AIR_FRYERS_URL,
    category_url_matches_path,
    clean_home_category_entries,
    normalized_category_url,
)


class CategoryMappingTests(unittest.TestCase):
    def test_department_mismatch_is_rejected(self):
        self.assertTrue(
            category_url_matches_path(
                "Home & Kitchen > Bath",
                "https://www.amazon.com/gp/bestsellers/home-garden/1063236",
            )
        )
        self.assertFalse(
            category_url_matches_path(
                "Home & Kitchen > Bath",
                "https://www.amazon.com/gp/bestsellers/pet-supplies/2975312011",
            )
        )

    def test_cleaner_removes_legacy_and_duplicate_paths(self):
        categories = {
            "Home & Kitchen > Kitchen & Dining": {
                "url": "https://www.amazon.com/gp/bestsellers/home-garden/284507"
            },
            "Home & Kitchen > Kitchen & Dining > Coffee, Tea & Espresso": {
                "url": "https://www.amazon.com/gp/bestsellers/home-garden/915194"
            },
            "Home & Kitchen > Kitchen & Dining > Coffee, Tea & Espresso > Coffee Makers": {
                "url": "https://www.amazon.com/gp/bestsellers/home-garden/7740213011"
            },
            "Home & Kitchen > Kitchen & Dining > Coffee Machines > Coffee Makers": {
                "url": "https://www.amazon.com/gp/bestsellers/home-garden/7740213011"
            },
            "Home & Kitchen > Kitchen & Dining > Food Storage > Dogs": {
                "url": "https://www.amazon.com/gp/bestsellers/pet-supplies/2975312011"
            },
        }

        cleaned, report = clean_home_category_entries(categories)

        self.assertIn(
            "Home & Kitchen > Kitchen & Dining > Coffee, Tea & Espresso > Coffee Makers",
            cleaned,
        )
        self.assertNotIn(
            "Home & Kitchen > Kitchen & Dining > Coffee Machines > Coffee Makers",
            cleaned,
        )
        self.assertNotIn("Home & Kitchen > Kitchen & Dining > Food Storage > Dogs", cleaned)
        self.assertEqual(cleaned[AIR_FRYERS_PATH]["url"], AIR_FRYERS_URL)
        self.assertIn(AIR_FRYERS_PARENT_PATH, cleaned)
        self.assertGreaterEqual(report["legacy_paths_removed"], 1)
        self.assertGreaterEqual(report["cross_department_removed"], 1)

    def test_url_normalization_removes_ref_and_query(self):
        self.assertEqual(
            normalized_category_url(
                "https://www.amazon.com/gp/bestsellers/home-garden/17659096011/ref=zg_test?x=1"
            ),
            AIR_FRYERS_URL,
        )


if __name__ == "__main__":
    unittest.main()
