from __future__ import annotations

from collections import OrderedDict
from collections import defaultdict
import re
from urllib.parse import urlsplit, urlunsplit


DEPARTMENT_SLUG_BY_ROOT = {
    "Appliances": "appliances",
    "Home & Kitchen": "home-garden",
    "Pet Supplies": "pet-supplies",
}

# These were early demo-only labels. They do not match Amazon's real hierarchy
# and caused discovered sibling categories to be stored below the wrong parent.
LEGACY_HOME_PREFIXES = (
    "Home & Kitchen > Kitchen & Dining > Coffee Machines",
    "Home & Kitchen > Kitchen & Dining > Air Fryers",
    "Home & Kitchen > Kitchen & Dining > Food Storage",
    "Home & Kitchen > Kitchen & Dining > Kitchen Gadgets",
    "Home & Kitchen > Storage & Organization > Closet Systems",
    "Home & Kitchen > Storage & Organization > Laundry Storage",
    "Home & Kitchen > Storage & Organization > Kitchen Storage",
)

AIR_FRYERS_PARENT_PATH = "Home & Kitchen > Kitchen & Dining > Small Appliances > Fryers"
AIR_FRYERS_PARENT_URL = "https://www.amazon.com/gp/bestsellers/home-garden/17659095011"
AIR_FRYERS_PATH = f"{AIR_FRYERS_PARENT_PATH} > Air Fryers"
AIR_FRYERS_URL = "https://www.amazon.com/gp/bestsellers/home-garden/17659096011"


def category_root(path: str) -> str:
    return str(path or "").split(" > ", 1)[0].strip()


def normalized_category_url(url: str) -> str:
    value = str(url or "").strip()
    if not value:
        return ""
    parts = urlsplit(value)
    path = parts.path.split("/ref=", 1)[0].rstrip("/")
    return urlunsplit((parts.scheme.lower(), parts.netloc.lower(), path, "", ""))


def canonical_category_url(url: str) -> str:
    normalized = normalized_category_url(url)
    return re.sub(
        r"/gp/new-releases/",
        "/gp/bestsellers/",
        normalized,
        count=1,
    )


def category_url_matches_path(path: str, url: str) -> bool:
    slug = DEPARTMENT_SLUG_BY_ROOT.get(category_root(path))
    if not slug:
        return True
    normalized = normalized_category_url(url)
    return f"/gp/bestsellers/{slug}/" in normalized or f"/gp/new-releases/{slug}/" in normalized


def is_legacy_home_path(path: str) -> bool:
    return any(path == prefix or path.startswith(f"{prefix} > ") for prefix in LEGACY_HOME_PREFIXES)


def mapping_audit_report(categories: dict) -> dict:
    allowed_roots = tuple(DEPARTMENT_SLUG_BY_ROOT)
    unscoped_aliases = []
    missing_node = []
    department_mismatches = []
    urls_by_canonical: defaultdict[str, list[str]] = defaultdict(list)

    for path, payload in (categories or {}).items():
        payload = payload or {}
        if not any(path == root or path.startswith(f"{root} > ") for root in allowed_roots):
            unscoped_aliases.append(path)
        if not str(payload.get("node") or "").strip():
            missing_node.append(path)
        url = str(payload.get("url") or "")
        if not category_url_matches_path(path, url):
            department_mismatches.append(path)
        canonical_url = canonical_category_url(url)
        if canonical_url:
            urls_by_canonical[canonical_url].append(path)

    duplicate_groups = {
        url: paths
        for url, paths in urls_by_canonical.items()
        if len(paths) > 1
    }
    return {
        "total_records": len(categories or {}),
        "unscoped_aliases": sorted(unscoped_aliases),
        "missing_node": sorted(missing_node),
        "department_mismatches": sorted(department_mismatches),
        "duplicate_url_groups": len(duplicate_groups),
        "duplicate_urls": duplicate_groups,
    }


def clean_category_entries(categories: dict) -> tuple[dict, dict]:
    allowed_roots = tuple(DEPARTMENT_SLUG_BY_ROOT)
    cleaned = OrderedDict()
    report = {
        "records_before": len(categories or {}),
        "records_after": 0,
        "unscoped_aliases_removed": 0,
        "department_mismatches_removed": 0,
        "legacy_paths_removed": 0,
        "urls_canonicalized": 0,
    }

    for path, payload in (categories or {}).items():
        if not any(path == root or path.startswith(f"{root} > ") for root in allowed_roots):
            report["unscoped_aliases_removed"] += 1
            continue

        payload = dict(payload or {})
        raw_url = str(payload.get("url") or "")
        if not category_url_matches_path(path, raw_url):
            report["department_mismatches_removed"] += 1
            continue
        if is_legacy_home_path(path):
            report["legacy_paths_removed"] += 1
            continue

        canonical_url = canonical_category_url(raw_url)
        if canonical_url != normalized_category_url(raw_url):
            report["urls_canonicalized"] += 1
        payload["url"] = canonical_url
        cleaned[path] = payload

    report["records_after"] = len(cleaned)
    return dict(cleaned), report


def clean_home_category_entries(categories: dict) -> tuple[dict, dict]:
    cleaned = OrderedDict()
    home_by_url: OrderedDict[str, tuple[str, dict]] = OrderedDict()
    report = {
        "home_before": 0,
        "home_after": 0,
        "cross_department_removed": 0,
        "legacy_paths_removed": 0,
        "duplicate_urls_removed": 0,
        "orphan_paths_removed": 0,
    }

    for path, payload in categories.items():
        if not path.startswith("Home & Kitchen"):
            cleaned[path] = payload
            continue

        report["home_before"] += 1
        url = str((payload or {}).get("url") or "")
        if path == "Home & Kitchen":
            continue
        if not category_url_matches_path(path, url):
            report["cross_department_removed"] += 1
            continue
        if is_legacy_home_path(path):
            report["legacy_paths_removed"] += 1
            continue

        normalized_url = normalized_category_url(url)
        if normalized_url in home_by_url:
            report["duplicate_urls_removed"] += 1
            continue
        normalized_payload = dict(payload or {})
        normalized_payload["url"] = normalized_url
        home_by_url[normalized_url] = (path, normalized_payload)

    known_paths = {path for path, _ in home_by_url.values()}
    stable_home = OrderedDict()
    for url, (path, payload) in home_by_url.items():
        parts = path.split(" > ")
        missing_parent = any(
            " > ".join(parts[:depth]) not in known_paths
            for depth in range(2, len(parts))
        )
        if missing_parent:
            report["orphan_paths_removed"] += 1
            continue
        stable_home[path] = payload

    stable_home.setdefault(
        AIR_FRYERS_PARENT_PATH,
        {
            "title": "Fryers",
            "url": AIR_FRYERS_PARENT_URL,
            "source": "Home & Kitchen > Kitchen & Dining > Small Appliances",
            "is_leaf": False,
            "node": "17659095011",
            "depth": 3,
        },
    )
    stable_home[AIR_FRYERS_PATH] = {
        "title": "Air Fryers",
        "url": AIR_FRYERS_URL,
        "source": AIR_FRYERS_PARENT_PATH,
        "is_leaf": True,
        "node": "17659096011",
        "depth": 4,
    }

    cleaned.update(stable_home)
    report["home_after"] = len(stable_home)
    return dict(cleaned), report
