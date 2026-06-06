import unittest

from src.chrome_cdp import count_hydrated_products
from src.sellersprite_parser import parse_sellersprite_text


class SellerSpriteParserTests(unittest.TestCase):
    def test_repeated_asin_keeps_later_enriched_fields(self):
        text = (
            "\n#1\nFirst title\nASIN:B0GSTH883W\n$59.98\n"
            "4.5 out of 5 stars\n58\n"
            "\n#1\nFirst title\nASIN:B0GSTH883W\n"
            "\u54c1\u724c: Frida Mom\n"
            "\u5356\u5bb6: Amazon\n"
            "\u8fd130\u5929\u9500\u91cf(\u7236\u4f53): 1,234\n"
            "\u8fd130\u5929\u9500\u91cf(\u5b50\u4f53): 1,000+\n"
            "\u9500\u552e\u989d: $74,015\n"
            "FBA\u8d39\u7528: $8.50\n"
            "#25 in Home & Kitchen\n"
            "$59.98\n4.5 out of 5 stars\n58\n"
            "\n#2\nSecond title\nASIN:B000000002\n$10.00\n"
        )

        products = parse_sellersprite_text(text, limit=50)

        self.assertEqual(2, len(products))
        product = products[0]
        self.assertEqual("B0GSTH883W", product.asin)
        self.assertEqual(1234, product.parent_monthly_sales)
        self.assertEqual(1000, product.child_monthly_sales)
        self.assertEqual(74015.0, product.sales_amount)
        self.assertEqual(8.5, product.fba_fee)
        self.assertEqual("Frida Mom", product.brand)

    def test_hydrated_count_deduplicates_snapshots(self):
        text = (
            "\n#1\nASIN:B0GSTH883W\n"
            "\u8fd130\u5929\u9500\u91cf(\u7236\u4f53): 860\n"
            "\n#1\nASIN:B0GSTH883W\n"
            "\u9500\u552e\u989d: $51,583\n"
        )

        self.assertEqual(1, count_hydrated_products(text))


if __name__ == "__main__":
    unittest.main()
