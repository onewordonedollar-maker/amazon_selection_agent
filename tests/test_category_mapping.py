import unittest

from src.category_mapping import (
    AIR_FRYERS_PARENT_PATH,
    AIR_FRYERS_PATH,
    AIR_FRYERS_URL,
    clean_category_entries,
    category_url_matches_path,
    clean_home_category_entries,
    mapping_audit_report,
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

    def test_safe_cleaner_removes_only_unscoped_aliases_and_canonicalizes_urls(self):
        categories = {
            "Cats": {
                "title": "Cats",
                "url": "https://www.amazon.com/gp/bestsellers/pet-supplies/2975241011",
            },
            "Pet Supplies > Cats": {
                "title": "Cats",
                "url": (
                    "https://www.amazon.com/gp/new-releases/pet-supplies/"
                    "2975241011/ref=zg_bsnr_test?x=1"
                ),
                "node": "2975241011",
            },
            "Home & Kitchen > Bath": {
                "title": "Bath",
                "url": "https://www.amazon.com/gp/bestsellers/home-garden/1063236",
                "node": "1063236",
            },
        }

        cleaned, report = clean_category_entries(categories)

        self.assertNotIn("Cats", cleaned)
        self.assertIn("Pet Supplies > Cats", cleaned)
        self.assertEqual(
            cleaned["Pet Supplies > Cats"]["url"],
            "https://www.amazon.com/gp/bestsellers/pet-supplies/2975241011",
        )
        self.assertIn("Home & Kitchen > Bath", cleaned)
        self.assertEqual(report["unscoped_aliases_removed"], 1)
        self.assertEqual(report["urls_canonicalized"], 1)

    def test_audit_reports_unscoped_aliases_without_marking_valid_paths(self):
        categories = {
            "Dogs": {
                "url": "https://www.amazon.com/gp/bestsellers/pet-supplies/2975312011"
            },
            "Pet Supplies > Dogs": {
                "url": "https://www.amazon.com/gp/bestsellers/pet-supplies/2975312011",
                "node": "2975312011",
            },
        }

        report = mapping_audit_report(categories)

        self.assertEqual(report["total_records"], 2)
        self.assertEqual(report["unscoped_aliases"], ["Dogs"])
        self.assertEqual(report["missing_node"], ["Dogs"])
        self.assertEqual(report["duplicate_url_groups"], 1)


if __name__ == "__main__":
    unittest.main()
