import unittest

from src.category_selection import (
    category_path_selected,
    selected_compact_category_paths,
    toggle_compact_category_selection,
)


class CategorySelectionTests(unittest.TestCase):
    def setUp(self):
        self.tree = {
            "Home & Kitchen": {
                "children": {
                    "Kitchen & Dining": {
                        "children": {
                            "Air Fryers": {"children": {}},
                            "Coffee Machines": {"children": {}},
                        },
                    },
                    "Home Decor": {
                        "children": {
                            "Artificial Plants": {"children": {}},
                        },
                    },
                },
            },
            "Pet Supplies": {
                "children": {
                    "Dogs": {"children": {}},
                },
            },
        }

    def test_checked_parent_is_returned_without_expanding_descendants(self):
        tree = {
            "Home & Kitchen": {
                "children": {
                    "Kitchen & Dining": {
                        "children": {
                            "Air Fryers": {"children": {}},
                        },
                    },
                },
            },
        }
        visited = []

        def is_selected(path):
            visited.append(path)
            return path == "Home & Kitchen"

        self.assertEqual(
            selected_compact_category_paths(tree, is_selected),
            ["Home & Kitchen"],
        )
        self.assertEqual(visited, ["Home & Kitchen"])

    def test_child_selection_is_preserved_when_parent_is_not_selected(self):
        tree = {
            "Home & Kitchen": {
                "children": {
                    "Kitchen & Dining": {
                        "children": {
                            "Air Fryers": {"children": {}},
                        },
                    },
                },
            },
        }
        selected = {"Home & Kitchen > Kitchen & Dining > Air Fryers"}

        self.assertEqual(
            selected_compact_category_paths(tree, selected.__contains__),
            ["Home & Kitchen > Kitchen & Dining > Air Fryers"],
        )

    def test_selecting_parent_stores_only_parent_path(self):
        self.assertEqual(
            toggle_compact_category_selection(self.tree, [], "Home & Kitchen", True),
            ["Home & Kitchen"],
        )

    def test_parent_selection_covers_descendants(self):
        selected = ["Home & Kitchen"]

        self.assertTrue(category_path_selected("Home & Kitchen", selected))
        self.assertTrue(category_path_selected("Home & Kitchen > Kitchen & Dining", selected))
        self.assertTrue(category_path_selected("Home & Kitchen > Kitchen & Dining > Air Fryers", selected))
        self.assertFalse(category_path_selected("Pet Supplies", selected))

    def test_unselecting_parent_removes_branch(self):
        self.assertEqual(
            toggle_compact_category_selection(self.tree, ["Home & Kitchen"], "Home & Kitchen", False),
            [],
        )

    def test_unselecting_child_under_selected_parent_keeps_sibling_branches_compact(self):
        self.assertEqual(
            toggle_compact_category_selection(
                self.tree,
                ["Home & Kitchen"],
                "Home & Kitchen > Kitchen & Dining > Air Fryers",
                False,
            ),
            [
                "Home & Kitchen > Home Decor",
                "Home & Kitchen > Kitchen & Dining > Coffee Machines",
            ],
        )


if __name__ == "__main__":
    unittest.main()
