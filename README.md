# Amazon Selection Agent

Amazon product collection, filtering, selection, and Excel export tool.

The application uses a dedicated local Chrome profile with the SellerSprite
browser extension. It opens Amazon ranking pages, waits for extension data,
parses products, applies filters, and exports selected products.

## Start

Double-click:

```text
一键启动工具.bat
```

The launcher:

1. Starts Streamlit at `http://localhost:8501/`.
2. Opens the dedicated collection Chrome.
3. Uses `D:\amazon_selection_agent\chrome_profile` as its profile.
4. Enables Chrome remote debugging on port `9222`.

First-time setup in that Chrome profile:

1. Log in to Amazon.
2. Install and log in to SellerSprite.
3. Keep this Chrome window open while collecting.

## Workflow

1. Select Best Sellers or New Releases.
2. Select one or more mapped categories.
3. Set the product filters.
4. Click Start Collection.
5. Wait while the application opens Amazon pages and reads SellerSprite data.
6. Review and select the filtered products.
7. Export selected products to Excel.

## Filters

Blank values mean unlimited.

- Price: minimum and maximum
- Review count: minimum and maximum
- Monthly sales: parent monthly sales minimum and maximum
- Child sales: child monthly sales minimum and maximum
- BSR: main category rank minimum and maximum
- Launch date: unrestricted or a selected time range

The initial UI values are:

- Minimum price: `24.99`
- Maximum review count: `300`
- Minimum monthly sales: `100`

These are UI prefilled values, not hidden backend defaults.

## Category Mappings

The repository includes the validated category mapping file:

```text
outputs/category_links_learned.json
```

Appliances, Home & Kitchen, and Pet Supplies mappings are built from live
Amazon Best Sellers category navigation. Runtime collection results,
checkpoints, backups, and browser data remain excluded from Git.

## Export

The application exports `.xlsx` files. The image column uses Excel's
`IMAGE()` formula and requires a modern Microsoft 365 Excel version.

## Local Data

The following folders are intentionally not committed:

- `chrome_profile/`: Amazon login and extension state
- Runtime files in `outputs/`, except the validated category mapping
- Python caches and virtual environments

## Maintenance Notes

Before changing collection, filtering, category mapping, or UI behavior, read:

```text
docs/lessons-learned/
```

It contains the project guardrails, known failure modes, tool handoff rules,
and the stable-release checklist.
