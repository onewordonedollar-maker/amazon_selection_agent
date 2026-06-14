import unittest

from src.collection_retry import CollectionRetryQueue


class CollectionRetryQueueTests(unittest.TestCase):
    def test_deferred_entries_keep_first_pass_order(self):
        queue = CollectionRetryQueue()
        queue.defer(
            label="Ladders",
            url="https://example.test/ladders",
            pages=1,
            products=["A"],
            error="timeout",
        )
        queue.defer(
            label="Nests",
            url="https://example.test/nests",
            pages=2,
            products=["B"],
            error="timeout",
        )

        self.assertEqual([item.label for item in queue], ["Ladders", "Nests"])
        self.assertEqual(len(queue), 2)

    def test_deferred_products_are_snapshotted(self):
        products = ["A"]
        queue = CollectionRetryQueue()
        queue.defer(
            label="Ladders",
            url="https://example.test/ladders",
            pages=1,
            products=products,
            error="timeout",
        )
        products.append("B")

        self.assertEqual(next(iter(queue)).products, ["A"])


if __name__ == "__main__":
    unittest.main()
