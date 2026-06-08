import unittest

import rebuild_home_categories


class HomeRebuildSafetyTests(unittest.TestCase):
    def test_safety_thresholds_are_not_trivial(self):
        self.assertGreaterEqual(rebuild_home_categories.MIN_ROOT_CHILDREN, 10)
        self.assertGreaterEqual(rebuild_home_categories.MIN_HOME_CATEGORIES, 500)


if __name__ == "__main__":
    unittest.main()
