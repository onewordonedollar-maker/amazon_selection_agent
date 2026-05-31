from __future__ import annotations

import re
from dataclasses import dataclass
from html import unescape
from urllib.request import Request, urlopen


DEFAULT_TEST_URL = "https://www.amazon.com/gp/bestsellers/pet-supplies/2975424011/ref=pd_zg_hrsr_pet-supplies"


@dataclass
class ScrapedProduct:
    rank: int
    title: str
    asin: str
    amazon_url: str
    image_url: str
    price: float
    rating: float
    review_count: int
    monthly_bought: int
    brand: str


def fetch_best_sellers(url: str = DEFAULT_TEST_URL, limit: int = 30, timeout: int = 20) -> list[ScrapedProduct]:
    html = fetch_html(url, timeout=timeout)
    products = parse_best_sellers(html, limit=limit)
    if not products:
        raise ValueError("No products parsed from Amazon best sellers page.")
    return products


def fetch_html(url: str, timeout: int = 20) -> str:
    request = Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/125.0.0.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
            "Cache-Control": "no-cache",
        },
    )
    with urlopen(request, timeout=timeout) as response:
        charset = response.headers.get_content_charset() or "utf-8"
        return response.read().decode(charset, errors="replace")


def parse_best_sellers(html: str, limit: int = 30) -> list[ScrapedProduct]:
    products: list[ScrapedProduct] = []
    seen: set[str] = set()
    for block in iter_product_blocks(html):
        asin = extract_asin(block)
        if not asin or asin in seen:
            continue
        seen.add(asin)
        title = extract_title(block)
        if not title:
            continue
        rank = extract_rank(block) or len(products) + 1
        monthly_bought = extract_monthly_bought(block)
        if not monthly_bought:
            monthly_bought = extract_monthly_bought(find_asin_neighborhood(html, asin))
        products.append(
            ScrapedProduct(
                rank=rank,
                title=title,
                asin=asin,
                amazon_url=f"https://www.amazon.com/dp/{asin}",
                image_url=extract_image_url(block),
                price=extract_price(block),
                rating=extract_rating(block),
                review_count=extract_review_count(block),
                monthly_bought=monthly_bought,
                brand=guess_brand(title),
            )
        )
        if len(products) >= limit:
            break
    return products


def iter_product_blocks(html: str):
    markers = [match.start() for match in re.finditer(r'data-asin="[^"]{10}"', html)]
    for index, start in enumerate(markers):
        end = markers[index + 1] if index + 1 < len(markers) else start + 12000
        yield html[max(0, start - 8000) : min(len(html), end + 8000)]


def find_asin_neighborhood(html: str, asin: str) -> str:
    positions = [
        match.start()
        for match in re.finditer(re.escape(asin), html)
    ]
    if not positions:
        return ""
    chunks = []
    for position in positions[:5]:
        chunks.append(html[max(0, position - 16000) : min(len(html), position + 16000)])
    return "\n".join(chunks)


def extract_asin(block: str) -> str:
    match = re.search(r'data-asin="([A-Z0-9]{10})"', block)
    if match:
        return match.group(1)
    match = re.search(r"/dp/([A-Z0-9]{10})", block)
    return match.group(1) if match else ""


def extract_rank(block: str) -> int:
    matches = re.findall(r">#([0-9,]+)<", block)
    return int(matches[-1].replace(",", "")) if matches else 0


def extract_title(block: str) -> str:
    patterns = [
        r'<div[^>]+_cDEzb_p13n-sc-css-line-clamp[^>]*>(.*?)</div>',
        r'<img[^>]+alt="([^"]+)"',
        r'aria-label="([^"]+)"',
    ]
    for pattern in patterns:
        match = re.search(pattern, block, flags=re.S)
        if match:
            title = clean_text(match.group(1))
            if title and "out of 5 stars" not in title.lower():
                return title
    return ""


def extract_image_url(block: str) -> str:
    match = re.search(r'<img[^>]+src="([^"]+)"', block)
    return unescape(match.group(1)) if match else ""


def extract_price(block: str) -> float:
    match = re.search(r'\$([0-9]+(?:,[0-9]{3})*(?:\.[0-9]{2})?)', block)
    return float(match.group(1).replace(",", "")) if match else 0.0


def extract_rating(block: str) -> float:
    match = re.search(r'([0-9.]+) out of 5 stars', block)
    return float(match.group(1)) if match else 0.0


def extract_review_count(block: str) -> int:
    candidates = re.findall(r'aria-label="([0-9,]+)"', block)
    values = [int(candidate.replace(",", "")) for candidate in candidates]
    return max(values) if values else 0


def extract_monthly_bought(block: str) -> int:
    text = clean_text(block).replace("&plus;", "+")
    patterns = [
        r'([0-9]+(?:[,.][0-9]+)?)([KkMm]?)\s*\+?\s*bought\s+in\s+(?:the\s+)?past\s+month',
        r'([0-9]+(?:[,.][0-9]+)?)([KkMm]?)\s*\+?\s*purchased\s+in\s+(?:the\s+)?past\s+month',
        r'([0-9]+(?:[,.][0-9]+)?)([KkMm]?)\s*\+?\s*bought',
    ]
    values = []
    for pattern in patterns:
        for match in re.finditer(pattern, text, flags=re.I):
            values.append(parse_compact_number(match.group(1), match.group(2)))
    return max(values) if values else 0


def parse_compact_number(value: str, suffix: str = "") -> int:
    normalized = value.replace(",", "")
    number = float(normalized)
    suffix = suffix.upper()
    if suffix == "K":
        number *= 1000
    elif suffix == "M":
        number *= 1000000
    return int(number)


def guess_brand(title: str) -> str:
    for separator in (" - ", " | ", ","):
        if separator in title:
            title = title.split(separator, 1)[0]
            break
    return title.split()[0][:40] if title.split() else "Unknown"


def clean_text(value: str) -> str:
    value = re.sub(r"<[^>]+>", " ", value)
    value = unescape(value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()
