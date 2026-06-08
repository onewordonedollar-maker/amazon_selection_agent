from __future__ import annotations

import json
import sys
import time
from collections import deque
from datetime import datetime
from pathlib import Path

from src.category_mapping import (
    AIR_FRYERS_PARENT_PATH,
    AIR_FRYERS_PATH,
    AIR_FRYERS_URL,
    category_url_matches_path,
    normalized_category_url,
)
from src.chrome_cdp import discover_bestseller_category_links


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "outputs" / "category_links_learned.json"
CHECKPOINT = ROOT / "outputs" / "home_category_rebuild_checkpoint.json"
ROOT_PATH = "Home & Kitchen"
ROOT_URL = "https://www.amazon.com/gp/bestsellers/home-garden/ref=zg_bs_nav_home-garden_0"
MAX_CATEGORIES = 2500
MAX_RETRIES = 2
MIN_HOME_CATEGORIES = 500
MIN_ROOT_CHILDREN = 10

try:
    sys.stdout.reconfigure(errors="backslashreplace")
    sys.stderr.reconfigure(errors="backslashreplace")
except (AttributeError, ValueError):
    pass


def save_checkpoint(queue, discovered, visited, failures):
    CHECKPOINT.write_text(
        json.dumps(
            {
                "queue": list(queue),
                "discovered": discovered,
                "visited": sorted(visited),
                "failures": failures,
                "updated_at": datetime.now().isoformat(timespec="seconds"),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def load_checkpoint():
    if not CHECKPOINT.exists():
        return deque([(ROOT_PATH, ROOT_URL)]), {}, set(), {}
    data = json.loads(CHECKPOINT.read_text(encoding="utf-8"))
    return (
        deque((str(path), str(url)) for path, url in data.get("queue", [])),
        dict(data.get("discovered", {})),
        set(data.get("visited", [])),
        dict(data.get("failures", {})),
    )


def main():
    queue, discovered, visited, failures = load_checkpoint()
    queued_urls = {normalized_category_url(url) for _, url in queue}
    print("开始重建 Home & Kitchen 类目映射。可随时按 Ctrl+C，之后再次运行会继续。")

    try:
        while queue and len(discovered) < MAX_CATEGORIES:
            parent_path, parent_url = queue.popleft()
            normalized_parent = normalized_category_url(parent_url)
            queued_urls.discard(normalized_parent)
            if normalized_parent in visited:
                continue

            print(
                f"[{len(visited) + 1}] {parent_path} "
                f"| 待处理 {len(queue)} | 已发现 {len(discovered)}"
            )
            links = None
            last_error = ""
            for attempt in range(1, MAX_RETRIES + 1):
                try:
                    links = discover_bestseller_category_links(parent_url, max_links=200)
                    break
                except Exception as exc:
                    last_error = str(exc)
                    print(f"  第 {attempt} 次读取失败：{last_error}")
                    time.sleep(2)

            if links is None:
                failures[parent_path] = last_error or "未知错误"
                visited.add(normalized_parent)
                save_checkpoint(queue, discovered, visited, failures)
                continue
            if parent_path == ROOT_PATH and len(links) < MIN_ROOT_CHILDREN:
                raise RuntimeError(
                    f"Home & Kitchen 根页只识别到 {len(links)} 个直属子类，"
                    f"低于安全阈值 {MIN_ROOT_CHILDREN}；未覆盖正式映射。"
                )

            visited.add(normalized_parent)
            failures.pop(parent_path, None)
            if parent_path in discovered:
                discovered[parent_path]["is_leaf"] = not bool(links)
            for link in links:
                child_path = f"{parent_path} > {link.title}"
                child_url = normalized_category_url(link.url)
                if not category_url_matches_path(child_path, child_url):
                    continue
                if child_url in visited or child_url in queued_urls:
                    continue
                discovered[child_path] = {
                    "title": link.title,
                    "url": child_url,
                    "source": parent_path,
                    "is_leaf": bool(link.is_leaf),
                    "node": str(link.node or ""),
                    "depth": child_path.count(" > "),
                    "updated_at": datetime.now().isoformat(timespec="seconds"),
                }
                queue.append((child_path, child_url))
                queued_urls.add(child_url)

            save_checkpoint(queue, discovered, visited, failures)
    except KeyboardInterrupt:
        save_checkpoint(queue, discovered, visited, failures)
        print("\n已保存进度。再次运行本脚本可继续。")
        return

    if queue:
        raise RuntimeError(f"类目超过安全上限 {MAX_CATEGORIES}，未覆盖正式映射。")
    if failures:
        save_checkpoint(queue, discovered, visited, failures)
        raise RuntimeError(f"仍有 {len(failures)} 个页面失败，未覆盖正式映射。")
    if len(discovered) < MIN_HOME_CATEGORIES:
        save_checkpoint(queue, discovered, visited, failures)
        raise RuntimeError(
            f"只发现 {len(discovered)} 个 Home & Kitchen 类目，"
            f"低于安全阈值 {MIN_HOME_CATEGORIES}；未覆盖正式映射。"
        )

    discovered.setdefault(
        AIR_FRYERS_PARENT_PATH,
        {
            "title": "Fryers",
            "url": "https://www.amazon.com/gp/bestsellers/home-garden/17659095011",
            "source": "Home & Kitchen > Kitchen & Dining > Small Appliances",
            "is_leaf": False,
            "node": "17659095011",
            "depth": 3,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        },
    )
    discovered[AIR_FRYERS_PATH] = {
        "title": "Air Fryers",
        "url": AIR_FRYERS_URL,
        "source": AIR_FRYERS_PARENT_PATH,
        "is_leaf": True,
        "node": "17659096011",
        "depth": 4,
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }

    data = json.loads(OUTPUT.read_text(encoding="utf-8")) if OUTPUT.exists() else {}
    categories = {
        path: payload
        for path, payload in data.get("categories", {}).items()
        if not path.startswith("Home & Kitchen")
    }
    categories.update(discovered)
    data["categories"] = categories
    backup = OUTPUT.with_name(
        f"category_links_learned.before_home_rebuild_{datetime.now():%Y%m%d_%H%M%S}.json"
    )
    if OUTPUT.exists():
        backup.write_text(OUTPUT.read_text(encoding="utf-8"), encoding="utf-8")
    OUTPUT.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    CHECKPOINT.unlink(missing_ok=True)
    print(f"完成：写入 {len(discovered)} 个 Home & Kitchen 类目。备份：{backup.name}")


if __name__ == "__main__":
    main()
