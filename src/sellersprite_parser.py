from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path


DOM_CACHE_PATH = Path(__file__).resolve().parents[1] / "outputs" / "sellersprite_dom.txt"


@dataclass
class SellerSpriteProduct:
    rank: int
    title: str
    asin: str
    brand: str
    seller: str
    fulfillment: str
    seller_count: int
    bsr_rank: int
    bsr_category: str
    sub_rank: int
    sub_category: str
    parent_monthly_sales: int
    child_monthly_sales: int
    child_monthly_sales_label: str
    sales_amount: float
    fba_fee: float
    margin_rate: str
    variant_count: int
    price: float
    rating: float
    review_count: int
    package_weight_lb: float
    package_dimensions: str
    launched_at: str
    parent_monthly_sales_loaded: bool = False


def load_cached_sellersprite_products(path: Path = DOM_CACHE_PATH, limit: int = 50) -> list[SellerSpriteProduct]:
    if not path.exists():
        raise FileNotFoundError(f"SellerSprite DOM cache not found: {path}")
    return parse_sellersprite_text(path.read_text(encoding="utf-8", errors="replace"), limit=limit)


def parse_sellersprite_text(text: str, limit: int = 50) -> list[SellerSpriteProduct]:
    text = normalize_text(text)
    starts = [m.start() for m in re.finditer(r"\n#\d+\n", text)]
    products_by_asin: dict[str, SellerSpriteProduct] = {}
    for index, start in enumerate(starts):
        end = starts[index + 1] if index + 1 < len(starts) else len(text)
        block = text[start:end].strip()
        asin = find_first(r"ASIN:([A-Z0-9]{10})", block)
        if not asin:
            continue
        candidate = parse_product_block(block, asin)
        existing = products_by_asin.get(asin)
        products_by_asin[asin] = (
            merge_sellersprite_products(existing, candidate)
            if existing
            else candidate
        )
    return list(products_by_asin.values())[:limit]


def sellersprite_product_completeness(product: SellerSpriteProduct) -> int:
    """Score plugin enrichment, not Amazon's basic card fields."""
    values = (
        product.parent_monthly_sales,
        product.child_monthly_sales,
        product.sales_amount,
        product.fba_fee,
        product.bsr_rank,
        product.sub_rank,
        product.seller_count,
        product.variant_count,
        product.brand,
        product.seller,
        product.fulfillment,
        product.margin_rate,
        product.launched_at,
        product.package_dimensions,
        product.package_weight_lb,
    )
    return sum(bool(value) for value in values)


def sellersprite_product_hydrated(product: SellerSpriteProduct) -> bool:
    return product.parent_monthly_sales_loaded


def merge_sellersprite_products(
    existing: SellerSpriteProduct,
    candidate: SellerSpriteProduct,
) -> SellerSpriteProduct:
    """Merge repeated snapshots of one ASIN, preferring enriched values."""
    existing_score = sellersprite_product_completeness(existing)
    candidate_score = sellersprite_product_completeness(candidate)
    primary, secondary = (
        (candidate, existing)
        if candidate_score >= existing_score
        else (existing, candidate)
    )
    values = {}
    for field_name in SellerSpriteProduct.__dataclass_fields__:
        primary_value = getattr(primary, field_name)
        secondary_value = getattr(secondary, field_name)
        values[field_name] = primary_value if primary_value not in ("", 0, 0.0, None) else secondary_value
    return SellerSpriteProduct(**values)


def parse_product_block(block: str, asin: str) -> SellerSpriteProduct:
    lines = [line.strip() for line in block.splitlines() if line.strip()]
    rank = parse_int(find_first(r"^#([0-9,]+)", block, flags=re.M))
    asin_index = lines.index(f"ASIN:{asin}") if f"ASIN:{asin}" in lines else -1
    title = extract_title(lines, asin_index)
    rating, reviews = extract_rating_reviews(lines)
    bsr_rank, bsr_category = extract_rank_category(block, 0)
    sub_rank, sub_category = extract_rank_category(block, 1)
    child_label = find_first(r"近30天销量\(子体\):\s*([^\n]+)", block)
    parent_sales_label = find_first(r"近30天销量\(父体\):\s*([^\n]+)", block)
    parent_monthly_sales = parse_compact_int(parent_sales_label)
    sales_amount = parse_money(find_first(r"销售额:\s*\$?([0-9,]+(?:\.[0-9]+)?)", block))
    labeled_price = parse_money(find_first(r"价格:\s*\$?([0-9,]+(?:\.[0-9]+)?)", block))
    card_price = extract_price(lines, asin_index)
    inferred_price = round(sales_amount / parent_monthly_sales, 2) if sales_amount and parent_monthly_sales else 0.0
    return SellerSpriteProduct(
        rank=rank,
        title=title,
        asin=asin,
        brand=find_first(r"品牌:\s*\n?([^\n]+)", block),
        seller=find_first(r"卖家:\s*([^\n]+)", block),
        fulfillment=find_first(r"配送:\s*([A-Z]+)", block),
        seller_count=parse_int(find_any([r"配送:\s*[A-Z]+卖家:\s*([0-9,]+)", r"卖家:\s*([0-9,]+)\s*家"], block)),
        bsr_rank=bsr_rank,
        bsr_category=bsr_category,
        sub_rank=sub_rank,
        sub_category=sub_category,
        parent_monthly_sales=parent_monthly_sales,
        child_monthly_sales=parse_compact_int(child_label),
        child_monthly_sales_label=child_label,
        sales_amount=sales_amount,
        fba_fee=parse_money(find_first(r"FBA费用:\s*\n?\$?([0-9,]+(?:\.[0-9]+)?)", block)),
        margin_rate=find_first(r"毛利率:\s*([^\n]+)", block),
        variant_count=parse_int(find_first(r"变体数:\s*([0-9,]+)", block)),
        price=labeled_price or card_price or inferred_price,
        rating=rating,
        review_count=reviews,
        package_weight_lb=parse_weight_lb(find_first(r"包装重量:\s*([^\n]+)", block)),
        package_dimensions=find_first(r"包装尺寸:\s*([^\n]+)", block),
        launched_at=find_first(r"上架时间:\s*([0-9]{4}-[0-9]{2}-[0-9]{2})", block),
        parent_monthly_sales_loaded=bool(parent_sales_label),
    )


def extract_title(lines: list[str], asin_index: int) -> str:
    if asin_index <= 0:
        return ""
    for i in range(asin_index - 1, max(-1, asin_index - 8), -1):
        line = lines[i]
        if line.startswith("$") or "out of 5 stars" in line or re.fullmatch(r"[0-9,]+", line):
            continue
        if re.search(r"\boffers?\s+from\s+\$", line, flags=re.I):
            continue
        if line.startswith("#"):
            continue
        return line
    return ""


def extract_rating_reviews(lines: list[str]) -> tuple[float, int]:
    rating = 0.0
    reviews = 0
    for index, line in enumerate(lines):
        match = re.search(r"([0-9.]+) out of 5 stars", line)
        if match:
            rating = float(match.group(1))
            if index + 1 < len(lines):
                reviews = parse_int(lines[index + 1])
            break
    return rating, reviews


def extract_price(lines: list[str], asin_index: int = -1) -> float:
    candidate_lines = lines[:asin_index] if asin_index > 0 else []
    for line in reversed(candidate_lines):
        if re.fullmatch(r"\$[0-9,]+(?:\.[0-9]+)?", line):
            return parse_money(line)
        offer_match = re.search(r"\boffers?\s+from\s+\$([0-9,]+(?:\.[0-9]+)?)", line, flags=re.I)
        if offer_match:
            return parse_money(offer_match.group(1))
    return 0.0


def extract_rank_category(block: str, offset: int) -> tuple[int, str]:
    matches = re.findall(r"#([0-9,]+)\s+in\s+([^\n]+)", block)
    if len(matches) <= offset:
        return 0, ""
    return parse_int(matches[offset][0]), matches[offset][1].strip()


def find_first(pattern: str, text: str, flags: int = 0) -> str:
    match = re.search(pattern, text, flags)
    return match.group(1).strip() if match else ""


def find_any(patterns: list[str], text: str, flags: int = 0) -> str:
    for pattern in patterns:
        value = find_first(pattern, text, flags)
        if value:
            return value
    return ""


def parse_int(value: str) -> int:
    if not value:
        return 0
    value = re.sub(r"[^0-9]", "", value)
    return int(value) if value else 0


def parse_money(value: str) -> float:
    if not value:
        return 0.0
    value = value.replace("$", "").replace(",", "").strip()
    try:
        return float(value)
    except ValueError:
        return 0.0


def parse_compact_int(value: str) -> int:
    if not value:
        return 0
    value = value.replace(",", "").replace(" ", "").strip()
    value = re.sub(r"^[<>≤≥~]+", "", value)
    match = re.match(r"([0-9.]+)([KkMm]?)(\+?)", value)
    if not match:
        return 0
    number = float(match.group(1))
    suffix = match.group(2).upper()
    if suffix == "K":
        number *= 1000
    elif suffix == "M":
        number *= 1000000
    return int(number)


def parse_weight_lb(value: str) -> float:
    if not value:
        return 0.0
    lb_match = re.search(r"([0-9.]+)\s*pounds?", value, flags=re.I)
    if lb_match:
        return float(lb_match.group(1))
    oz_match = re.search(r"([0-9.]+)\s*ounces?", value, flags=re.I)
    if oz_match:
        return round(float(oz_match.group(1)) / 16, 2)
    return 0.0


def normalize_text(text: str) -> str:
    text = text.replace("\u2009", "")
    text = re.sub(r"\r\n?", "\n", text)
    return text
