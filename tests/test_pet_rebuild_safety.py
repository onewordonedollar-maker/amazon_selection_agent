import unittest

import rebuild_pet_categories as rebuild


class PetRebuildSafetyTests(unittest.TestCase):
    def test_safety_thresholds_are_not_trivial(self):
        self.assertGreaterEqual(rebuild.MIN_PET_CATEGORIES, 250)
        self.assertGreaterEqual(rebuild.MIN_ROOT_CHILDREN, 6)
        self.assertGreaterEqual(rebuild.MAX_RETRIES, 2)


if __name__ == "__main__":
    unittest.main()
