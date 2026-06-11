import unittest

import rebuild_appliances_categories as rebuild


class AppliancesRebuildSafetyTests(unittest.TestCase):
    def test_safety_thresholds_are_not_trivial(self):
        self.assertGreaterEqual(rebuild.MIN_APPLIANCES_CATEGORIES, 25)
        self.assertGreaterEqual(rebuild.MIN_ROOT_CHILDREN, 10)
        self.assertGreaterEqual(rebuild.MAX_RETRIES, 2)


if __name__ == "__main__":
    unittest.main()
