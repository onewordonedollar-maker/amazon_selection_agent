from collections.abc import Sequence
from math import ceil
from typing import TypeVar


PAGE_SIZE_OPTIONS = [50, 100, 200]
DEFAULT_PAGE_SIZE = PAGE_SIZE_OPTIONS[0]

T = TypeVar("T")


def normalize_page_size(value: int | None) -> int:
    return value if value in PAGE_SIZE_OPTIONS else DEFAULT_PAGE_SIZE


def page_count(total_items: int, page_size: int) -> int:
    normalized_size = normalize_page_size(page_size)
    return max(1, ceil(max(0, total_items) / normalized_size))


def clamp_page(page: int | None, total_items: int, page_size: int) -> int:
    try:
        requested_page = int(page or 1)
    except (TypeError, ValueError):
        requested_page = 1
    return min(max(1, requested_page), page_count(total_items, page_size))


def page_bounds(total_items: int, page: int, page_size: int) -> tuple[int, int]:
    normalized_size = normalize_page_size(page_size)
    safe_page = clamp_page(page, total_items, normalized_size)
    start = (safe_page - 1) * normalized_size
    end = min(start + normalized_size, max(0, total_items))
    return start, end


def page_start_index(total_items: int, page: int, page_size: int) -> int:
    start, _ = page_bounds(total_items, page, page_size)
    return start


def page_slice(items: Sequence[T], page: int, page_size: int) -> list[T]:
    start, end = page_bounds(len(items), page, page_size)
    return list(items[start:end])


def page_range_label(total_items: int, page: int, page_size: int) -> str:
    if total_items <= 0:
        return "显示 0 / 共 0 条"
    start, end = page_bounds(total_items, page, page_size)
    return f"显示 {start + 1:,}-{end:,} / 共 {total_items:,} 条"
