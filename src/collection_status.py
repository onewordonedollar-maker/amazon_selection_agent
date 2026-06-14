from __future__ import annotations


def build_collection_total_status_text(
    *,
    total_raw: int,
    completed: int,
    total_seeds: int,
    current_seed: int = 0,
    current_label: str = "",
    empty: int = 0,
    warning: int = 0,
    failed: int = 0,
) -> str:
    if not total_seeds:
        return f"本轮所有已选类目累计原始产品：**{total_raw:,} 条**"

    current_text = ""
    if current_seed > completed:
        current_text = f"｜正在处理：**{current_seed}/{total_seeds}**"
        if current_label:
            current_text += f"（{current_label}）"

    empty_text = f"｜空榜入口：**{empty} 个**" if empty else ""
    warning_text = f"｜数量偏少警告：**{warning} 个**" if warning else ""
    failure_text = f"｜失败入口：**{failed} 个**" if failed else ""
    return (
        f"本轮所有已选类目累计原始产品：**{total_raw:,} 条**｜"
        f"已完成小类入口：**{completed}/{total_seeds}**"
        f"{current_text}{empty_text}{warning_text}{failure_text}"
    )
