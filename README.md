# Amazon Selection Agent

Amazon product selection and export tool.

The current production workflow uses the SellerSprite browser extension rendered inside a dedicated local Chrome profile. The app opens Amazon pages, waits for SellerSprite data to load, parses the rendered page, applies filters, and exports selected products to Excel.

## Start

Double click:

```text
一键启动工具.bat
```

The launcher will:

1. Start Streamlit at `http://localhost:8501/`.
2. Open the dedicated collection Chrome.
3. Use `D:\amazon_selection_agent\chrome_profile` as the Chrome profile.
4. Enable Chrome remote debugging on port `9222`.

First-time setup in that Chrome profile:

1. Log in to Amazon.
2. Install and log in to SellerSprite.
3. Keep this Chrome window open while collecting.

## Main Workflow

1. Choose data source: `卖家精灵插件`.
2. Select categories, or paste an Amazon Best Sellers URL.
3. Set filters.
4. Click `开始采集`.
5. The app opens Amazon and waits for SellerSprite data.
6. Products are parsed, filtered, deduplicated by ASIN, and displayed.
7. Select products and export Excel.

## Batch Category Collection

When `大类批量采集` is checked, the app:

1. Uses selected category links when available.
2. Discovers child Best Sellers category links from the Amazon page.
3. Opens child categories one by one.
4. Waits for SellerSprite data.
5. Filters each category immediately.
6. Deduplicates by ASIN.
7. Shows progress and logs.

If a selected category has no mapped Amazon Best Sellers URL, the UI will warn you instead of silently using a wrong link.

Single-page collection URL priority:

1. Manually pasted Amazon URL.
2. Selected category mapped URL.
3. Default test URL.

## Filters

Blank values mean unlimited.

- Price: min / max
- Review count: min / max
- Monthly sales: parent monthly sales min / max
- Child sales: child monthly sales min / max
- BSR: main category rank min / max
- Launch date: any, last 30 days, last 60 days, last 3 months, last 6 months, last year, last 2 years, 1-2 years

## Export

The app exports `.xlsx` files.

The image column uses the Excel `IMAGE()` formula. It requires a modern Microsoft 365 Excel version to render images directly.

## Important Notes

- The main data source is SellerSprite extension data rendered on Amazon pages.
- Failed real collection does not fall back to sample data.
- `chrome_profile` contains local browser login and extension state. It is runtime data, not application source code.
