from __future__ import annotations

import json
import sys
import time
from collections import deque
from datetime import datetime
from pathlib import Path

from src.category_mapping import category_url_matches_path, normalized_category_url
from src.chrome_cdp import discover_bestseller_category_links


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "outputs" / "category_links_learned.json"
CHECKPOINT = ROOT / "outputs" / "appliances_category_rebuild_checkpoint.json"
ROOT_PATH = "Appliances"
ROOT_URL = "https://www.amazon.com/gp/bestsellers/appliances/ref=zg_bs_nav_appliances_0"
MAX_CATEGORIES = 1000
MAX_RETRIES = 3
MIN_APPLIANCES_CATEGORIES = 25
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
    print(
        "Starting Appliances category rebuild. "
        "Press Ctrl+C to save progress; rerun to continue."
    )

    try:
        while queue and len(discovered) < MAX_CATEGORIES:
            parent_path, parent_url = queue.popleft()
            normalized_parent = normalized_category_url(parent_url)
            queued_urls.discard(normalized_parent)
            if normalized_parent in visited:
                continue

            print(
                f"[{len(visited) + 1}] {parent_path} "
                f"| queued {len(queue)} | discovered {len(discovered)}"
            )
            links = None
            last_error = ""
            for attempt in range(1, MAX_RETRIES + 1):
                try:
                    links = discover_bestseller_category_links(parent_url, max_links=200)
                    break
                except Exception as exc:
                    last_error = str(exc)
                    print(f"  attempt {attempt} failed: {last_error}")
                    time.sleep(2)

            if links is None:
                failures[parent_path] = last_error or "unknown error"
                save_checkpoint(queue, discovered, visited, failures)
                raise RuntimeError(
                    f"Failed to read {parent_path}; checkpoint saved and formal mapping unchanged."
                )

            if parent_path == ROOT_PATH and len(links) < MIN_ROOT_CHILDREN:
                save_checkpoint(queue, discovered, visited, failures)
                raise RuntimeError(
                    f"Appliances root exposed only {len(links)} direct children, below "
                    f"the safety threshold {MIN_ROOT_CHILDREN}; formal mapping unchanged."
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
        print("\nProgress saved. Run this script again to continue.")
        return

    if queue:
        save_checkpoint(queue, discovered, visited, failures)
        raise RuntimeError(
            f"Category count exceeded the safety limit {MAX_CATEGORIES}; formal mapping unchanged."
        )
    if failures:
        save_checkpoint(queue, discovered, visited, failures)
        raise RuntimeError(
            f"{len(failures)} pages still failed; formal mapping unchanged."
        )
    if len(discovered) < MIN_APPLIANCES_CATEGORIES:
        save_checkpoint(queue, discovered, visited, failures)
        raise RuntimeError(
            f"Only {len(discovered)} Appliances categories were discovered, below "
            f"the safety threshold {MIN_APPLIANCES_CATEGORIES}; formal mapping unchanged."
        )

    data = json.loads(OUTPUT.read_text(encoding="utf-8")) if OUTPUT.exists() else {}
    categories = {
        path: payload
        for path, payload in data.get("categories", {}).items()
        if path != ROOT_PATH and not path.startswith(f"{ROOT_PATH} > ")
    }
    categories.update(discovered)
    data["categories"] = categories

    backup = OUTPUT.with_name(
        f"category_links_learned.before_appliances_rebuild_{datetime.now():%Y%m%d_%H%M%S}.json"
    )
    if OUTPUT.exists():
        backup.write_text(OUTPUT.read_text(encoding="utf-8"), encoding="utf-8")
    OUTPUT.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    CHECKPOINT.unlink(missing_ok=True)
    print(
        f"Completed: wrote {len(discovered)} Appliances categories. "
        f"Backup: {backup.name}"
    )


if __name__ == "__main__":
    main()
