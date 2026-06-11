# New Releases Entry Validation

Amazon does not guarantee that every Best Sellers node has a matching New
Releases page. Some converted URLs can fall back to `Any Department`, another
node, or an unavailable page.

The collection flow now validates each page after the existing initial
top-of-page wait:

- The final department and node must match the requested category.
- The selected category must not be `Any Department`.
- The page must not contain Amazon's explicit unavailable-page messages.
- Invalid entries are skipped and included in the failed-entry report.
- Valid entries continue through the existing wait, scroll, pagination,
  SellerSprite loading, deduplication, and filtering flow unchanged.
