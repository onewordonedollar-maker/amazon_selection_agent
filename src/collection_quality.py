from __future__ import annotations


QUANTITY_WARNING_PREFIX = "数量偏少警告"


def is_collection_quality_warning(message: str) -> bool:
    return str(message or "").startswith(QUANTITY_WARNING_PREFIX)


def is_sellersprite_load_failure(message: str) -> bool:
    text = str(message or "")
    return "卖家精灵" in text and ("未完全加载" in text or "未加载完整" in text)


def evaluate_sellersprite_collection_quality(
    unique_product_count: int,
    refresh_results: list,
    *,
    empty_message: str,
    expected_products_per_page: int,
    min_products_per_page: int,
    min_products_two_pages: int,
) -> tuple[bool, str]:
    if not refresh_results:
        return False, "没有读取到页面"
    if any(result.message == empty_message for result in refresh_results):
        return True, empty_message
    failed_results = [result for result in refresh_results if not bool(getattr(result, "ok", True))]
    if failed_results:
        return False, str(getattr(failed_results[-1], "message", "") or "页面采集未完成")

    page_counts = [int(getattr(result, "product_count", 0) or 0) for result in refresh_results]
    hydrated_counts = [int(getattr(result, "hydrated_count", 0) or 0) for result in refresh_results]
    page_detail = "；".join(
        f"第{i + 1}页：页面产品 {int(getattr(result, 'product_count', 0) or 0)} 条，"
        f"卖家精灵字段完整 {int(getattr(result, 'hydrated_count', 0) or 0)} 条"
        for i, result in enumerate(refresh_results)
    )
    if len(refresh_results) < 2:
        return False, f"只读取到 {len(refresh_results)} 页；{page_detail}。需要第 1 页和第 2 页都完成，才算一个入口完整。"

    weak_pages = [
        f"第{i + 1}页 {count} 条"
        for i, count in enumerate(page_counts[:2])
        if count < min_products_per_page
    ]
    if weak_pages:
        return False, (
            f"页面产品数偏少：{'，'.join(weak_pages)}；预期每页接近 {expected_products_per_page} 条。"
            f"{page_detail}。"
        )

    weak_plugin_pages = [
        f"第{i + 1}页 {hydrated_count}/{product_count} 条"
        for i, (product_count, hydrated_count) in enumerate(zip(page_counts[:2], hydrated_counts[:2]))
        if hydrated_count < min(product_count, min_products_per_page)
    ]
    if weak_plugin_pages:
        return False, (
            f"卖家精灵父体月销量字段未加载完整：{'，'.join(weak_plugin_pages)}。"
            f"{page_detail}。"
        )

    if unique_product_count < min_products_two_pages:
        return True, (
            f"{QUANTITY_WARNING_PREFIX}：两页均已完成且卖家精灵字段完整，"
            f"原始去重 {unique_product_count} 条，低于常见的接近 100 条。"
            f"{page_detail}。数据已正常保留，不计为失败。"
        )

    return True, f"两页采集正常：原始去重 {unique_product_count} 条。{page_detail}。"
