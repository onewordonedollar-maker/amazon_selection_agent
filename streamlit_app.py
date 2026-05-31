import csv
import json
import random
import re
import threading
import time
from copy import deepcopy
from dataclasses import MISSING, asdict, dataclass, fields
from datetime import datetime
from functools import lru_cache
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from html import escape
from io import BytesIO, StringIO
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from src.amazon_scraper import DEFAULT_TEST_URL, fetch_best_sellers
from src.chrome_cdp import (
    chrome_debugger_available,
    discover_bestseller_category_links,
    is_rank_category_url,
    refresh_sellersprite_cache_pages,
)
from src.sellersprite_parser import load_cached_sellersprite_products


OUTPUT_DIR = Path(__file__).resolve().parent / "outputs"
SELLERSPRITE_DOM_CACHE = OUTPUT_DIR / "sellersprite_dom.txt"
SELLERSPRITE_IMAGE_CACHE = OUTPUT_DIR / "sellersprite_images.json"
SELLERSPRITE_META_CACHE = OUTPUT_DIR / "sellersprite_cache_meta.json"
LEARNED_CATEGORY_LINKS = OUTPUT_DIR / "category_links_learned.json"
RAW_PRODUCTS_CACHE = OUTPUT_DIR / "last_raw_products.json"
RAW_PRODUCTS_HISTORY_DIR = OUTPUT_DIR / "raw_products"
RAW_PRODUCTS_HISTORY_INDEX = RAW_PRODUCTS_HISTORY_DIR / "index.json"
RAW_PRODUCTS_HISTORY_LIMIT = 5
STOP_COLLECTION_FLAG = OUTPUT_DIR / "stop_collection.flag"
STOP_COLLECTION_PORT = 8765
SELLERSPRITE_PLUGIN_FIELDS_TEXT = "价格、评分、评分数、排名、销量、销售额、FBA费用、毛利率、变体数、卖家数、包装信息"
SELLERSPRITE_EXPECTED_PRODUCTS_PER_PAGE = 50
SELLERSPRITE_MIN_PRODUCTS_PER_PAGE = 45
SELLERSPRITE_MIN_PRODUCTS_TWO_PAGES = 95
SELLERSPRITE_CATEGORY_RETRY_LIMIT = 1

SELLERSPRITE_EXPORT_COLUMNS = [
    ("image_preview_formula", "主图"),
    ("rank", "#"),
    ("title", "产品信息"),
    ("asin", "ASIN"),
    ("brand", "品牌"),
    ("seller_name", "卖家"),
    ("prime_fba", "配送"),
    ("seller_count", "卖家数"),
    ("bsr_rank", "大类BSR"),
    ("bsr_category", "大类目"),
    ("sub_rank", "子类排名"),
    ("sub_category", "子类目"),
    ("monthly_bought", "近30天销量(父体)"),
    ("child_monthly_sales_label", "近30天销量(子体)"),
    ("sales_amount", "销售额"),
    ("variant_count", "变体数"),
    ("price", "价格"),
    ("review_count", "评分数"),
    ("rating", "评分"),
    ("fba_fee", "FBA费用"),
    ("margin_rate", "毛利率"),
    ("launched_at", "上架时间"),
    ("package_weight_lb", "包装重量(lb)"),
    ("package_dimensions", "包装尺寸"),
    ("review_status", "审核状态"),
    ("note", "备注"),
    ("potential_score", "AI分"),
    ("potential_level", "评级"),
    ("reason", "推荐理由"),
    ("risk_tags", "风险词"),
    ("category_path", "浏览同类目"),
    ("amazon_url", "Amazon链接"),
    ("image_url", "主图链接"),
    ("scraped_at", "抓取时间"),
    ("status", "状态"),
    ("error", "错误"),
]


st.set_page_config(
    page_title="Amazon Selection Agent",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded",
)


CATEGORIES = {
    "Appliances": {
        "zh": "家电",
        "count": 80335,
        "children": {
            "Appliance Warranties": {"zh": "家电保修", "count": 2, "children": {}},
            "Dishwashers": {
                "zh": "洗碗机",
                "count": 1314,
                "children": {
                    "Built-In Dishwashers": {"zh": "嵌入式洗碗机", "count": 437},
                    "Countertop Dishwashers": {"zh": "台式洗碗机", "count": 542},
                    "Portable Dishwashers": {"zh": "便携式洗碗机", "count": 103},
                },
            },
            "Small Appliances": {
                "zh": "小家电",
                "count": 18640,
                "children": {
                    "Humidifiers": {"zh": "加湿器", "count": 3182},
                    "Portable Fans": {"zh": "便携风扇", "count": 2710},
                    "Electric Kettles": {"zh": "电热水壶", "count": 1788},
                    "Vacuums": {"zh": "吸尘器", "count": 4960},
                },
            },
        },
    },
    "Home & Kitchen": {
        "zh": "家居厨房",
        "count": 245980,
        "children": {
            "Kitchen & Dining": {
                "zh": "厨房餐厨",
                "count": 68220,
                "children": {
                    "Coffee Machines": {"zh": "咖啡机", "count": 2534},
                    "Air Fryers": {"zh": "空气炸锅", "count": 1876},
                    "Food Storage": {"zh": "食品收纳", "count": 5210},
                    "Kitchen Gadgets": {"zh": "厨房工具", "count": 14680},
                },
            },
            "Storage & Organization": {
                "zh": "收纳整理",
                "count": 54310,
                "children": {
                    "Closet Systems": {"zh": "衣柜系统", "count": 4280},
                    "Laundry Storage": {"zh": "洗衣收纳", "count": 3610},
                    "Kitchen Storage": {"zh": "厨房收纳", "count": 9212},
                },
            },
        },
    },
    "Pet Supplies": {
        "zh": "宠物用品",
        "count": 115420,
        "children": {
            "Dogs": {
                "zh": "狗用品",
                "count": 58400,
                "children": {
                    "Dog Carriers": {"zh": "狗包/狗笼", "count": 3180},
                    "Dog Grooming": {"zh": "狗狗美容", "count": 4260},
                    "Dog Feeding": {"zh": "狗狗喂食", "count": 6170},
                    "Dog Toys": {"zh": "狗玩具", "count": 9150},
                },
            },
            "Cats": {
                "zh": "猫用品",
                "count": 36200,
                "children": {
                    "Cat Litter": {"zh": "猫砂", "count": 1980},
                    "Cat Trees": {"zh": "猫爬架", "count": 2780},
                    "Cat Grooming": {"zh": "猫咪美容", "count": 1640},
                },
            },
        },
    },
}

CATEGORY_BESTSELLER_URLS = {
    "Appliances": "https://www.amazon.com/gp/bestsellers/appliances/ref=zg_bs_nav_appliances_0",
    "Home & Kitchen": "https://www.amazon.com/gp/bestsellers/home-garden/ref=zg_bs_nav_home-garden_0",
    "Pet Supplies": "https://www.amazon.com/gp/bestsellers/pet-supplies/ref=zg_bs_nav_pet-supplies_0",
    "Pet Supplies > Dogs": "https://www.amazon.com/gp/bestsellers/pet-supplies/2975312011",
    "Pet Supplies > Dogs > Dog Carriers": "https://www.amazon.com/gp/bestsellers/pet-supplies/2975333011",
    "Pet Supplies > Dogs > Dog Grooming": "https://www.amazon.com/gp/bestsellers/pet-supplies/2975362011",
    "Pet Supplies > Dogs > Dog Feeding": "https://www.amazon.com/gp/bestsellers/pet-supplies/2975351011",
    "Pet Supplies > Dogs > Dog Toys": "https://www.amazon.com/gp/bestsellers/pet-supplies/2975413011",
}

CATEGORY_TITLE_ALIASES = {
    "Pet Supplies > Dogs > Dog Carriers": ["Carriers & Travel Products", "Dog Carriers & Travel Products"],
    "Pet Supplies > Dogs > Dog Grooming": ["Grooming", "Dog Grooming Supplies"],
    "Pet Supplies > Dogs > Dog Feeding": ["Feeding & Watering Supplies", "Dog Feeding & Watering Supplies"],
    "Pet Supplies > Dogs > Dog Toys": ["Toys", "Dog Toys"],
    "Pet Supplies > Cats > Cat Litter": ["Litter & Housebreaking", "Cat Litter"],
    "Pet Supplies > Cats > Cat Trees": ["Trees", "Cat Trees"],
    "Pet Supplies > Cats > Cat Grooming": ["Grooming", "Cat Grooming Supplies"],
}


TEXT = {
    "English": {
        "title": "Amazon Selection Agent",
        "caption": "Amazon product collection, filtering, selection, and Excel export.",
        "task": "Task",
        "ui_language": "UI language",
        "list_type": "List type",
        "categories": "Categories",
        "all_categories": "All categories",
        "filters": "Filters",
        "price": "Price",
        "reviews": "Reviews",
        "monthly_sales": "Monthly sales",
        "child_sales": "Child sales",
        "bsr": "BSR",
        "launched_at": "Launch date",
        "pages": "Pages per category",
        "custom_url": "Amazon URL (optional)",
        "run": "Start Collection",
        "clear": "Clear Results",
        "cards": "Product Cards",
        "table": "Table",
        "log": "Log",
        "open_amazon": "Open Amazon",
        "review_status": "Review status",
        "note": "Note",
        "include": "Include in export",
    },
    "中文": {
        "title": "亚马逊选品 AI 智能体",
        "caption": "亚马逊产品采集、筛选、勾选与 Excel 导出。",
        "task": "任务",
        "ui_language": "界面语言",
        "list_type": "榜单类型",
        "categories": "类目",
        "all_categories": "全部类目",
        "filters": "筛选条件",
        "price": "价格",
        "reviews": "评分数",
        "monthly_sales": "月销量",
        "child_sales": "子体销量",
        "bsr": "BSR",
        "launched_at": "上架时间",
        "pages": "每个类目页数",
        "custom_url": "指定 Amazon 链接（可选）",
        "run": "开始采集",
        "clear": "清空结果",
        "cards": "产品卡片",
        "table": "表格",
        "log": "日志",
        "open_amazon": "打开 Amazon",
        "review_status": "审核状态",
        "note": "备注",
        "include": "加入导出",
    },
}

BLOCKED_KEYWORDS = {
    "supplement",
    "vitamin",
    "adult",
    "book",
    "shoes",
    "clothing",
    "jewelry",
    "video game",
    "amazon basics",
    "kindle",
}


@dataclass
class Product:
    selected: bool
    review_status: str
    note: str
    list_type: str
    category_path: str
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
    prime_fba: str
    delivery: str
    variant_count: int
    launched_at: str
    package_dimensions: str
    package_weight_lb: float
    fba_fee: float
    potential_score: int
    potential_level: str
    reason: str
    risk_tags: str
    scraped_at: str
    status: str
    error: str
    seller_name: str = ""
    seller_count: int = 0
    sales_amount: float = 0.0
    child_monthly_sales: int = 0
    child_monthly_sales_label: str = ""
    bsr_rank: int = 0
    bsr_category: str = ""
    sub_rank: int = 0
    sub_category: str = ""
    margin_rate: str = ""


class CollectionStopped(RuntimeError):
    pass


def clear_stop_collection_flag() -> None:
    try:
        STOP_COLLECTION_FLAG.unlink(missing_ok=True)
    except OSError:
        pass


def stop_collection_requested() -> bool:
    return STOP_COLLECTION_FLAG.exists()


def raise_if_stop_requested() -> None:
    if stop_collection_requested():
        raise CollectionStopped("用户请求停止采集，已保留当前已完成的采集结果。")


class StopCollectionHandler(BaseHTTPRequestHandler):
    def _send(self, status: int = 200, body: str = "ok") -> None:
        data = body.encode("utf-8")
        origin = self.headers.get("Origin") or ""
        allowed_origins = {"http://localhost:8501", "http://127.0.0.1:8501", "null"}
        allowed_origin = origin if origin in allowed_origins else "http://localhost:8501"
        self.send_response(status)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", allowed_origin)
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS, GET")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_OPTIONS(self):
        self._send(204, "")

    def do_GET(self):
        if self.path == "/status":
            self._send(200, "stopping" if stop_collection_requested() else "running")
            return
        self._send(404, "not found")

    def do_POST(self):
        if self.path == "/stop":
            origin = self.headers.get("Origin") or ""
            if origin and origin not in {"http://localhost:8501", "http://127.0.0.1:8501", "null"}:
                self._send(403, "forbidden")
                return
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            STOP_COLLECTION_FLAG.write_text("stop", encoding="utf-8")
            self._send(200, "stop requested")
            return
        self._send(404, "not found")

    def log_message(self, format, *args):
        return


@st.cache_resource
def ensure_stop_collection_server() -> bool:
    try:
        server = ThreadingHTTPServer(("127.0.0.1", STOP_COLLECTION_PORT), StopCollectionHandler)
    except OSError:
        return False
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    return True


def request_stop_collection() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    STOP_COLLECTION_FLAG.write_text("stop", encoding="utf-8")
    st.session_state.last_collection_summary = "已请求停止采集：程序会在当前页面/小类结束后保留已完成数据。"
    log("Stop collection requested from UI.")


def render_stop_collection_button() -> None:
    server_ready = ensure_stop_collection_server()
    disabled_attr = "" if server_ready else "disabled"
    components.html(
        f"""
        <button id="stop-collection-btn" {disabled_attr} style="
            width: 100%;
            height: 40px;
            border: 1px solid #ff4b4b;
            border-radius: 8px;
            background: #fff;
            color: #ff4b4b;
            font-size: 15px;
            font-weight: 600;
            cursor: pointer;
        ">停止采集</button>
        <script>
        const btn = document.getElementById("stop-collection-btn");
        if (btn) {{
          btn.addEventListener("click", async () => {{
            btn.disabled = true;
            btn.textContent = "停止中...";
            try {{
              const res = await fetch("http://127.0.0.1:{STOP_COLLECTION_PORT}/stop", {{ method: "POST" }});
              btn.textContent = res.ok ? "已请求停止" : "停止失败";
            }} catch (error) {{
              btn.textContent = "停止失败";
            }}
          }});
        }}
        </script>
        """,
        height=44,
    )


def ensure_state():
    if "raw_products" not in st.session_state:
        st.session_state.raw_products = []
    if "products" not in st.session_state:
        st.session_state.products = []
    if "run_log" not in st.session_state:
        st.session_state.run_log = ["Ready."]
    if "confirmed_category_paths" not in st.session_state:
        st.session_state.confirmed_category_paths = []
    if "show_category_dialog" not in st.session_state:
        st.session_state.show_category_dialog = False
    if "last_cache_refresh_message" not in st.session_state:
        st.session_state.last_cache_refresh_message = ""
    if "last_category_mapping_message" not in st.session_state:
        st.session_state.last_category_mapping_message = ""
    if "last_collection_summary" not in st.session_state:
        st.session_state.last_collection_summary = ""
    if "last_raw_products_message" not in st.session_state:
        st.session_state.last_raw_products_message = ""
    if "category_search" not in st.session_state:
        st.session_state.category_search = ""


def log(message: str):
    st.session_state.run_log.append(f"{datetime.now().strftime('%H:%M:%S')}  {message}")


def sync_product_selection_from_widgets(products):
    for product in products:
        key = f"row_include_{product.asin}"
        if key in st.session_state:
            product.selected = bool(st.session_state[key])


def set_all_product_selection(products, selected: bool):
    for product in products:
        product.selected = selected
        st.session_state[f"row_include_{product.asin}"] = selected


def handle_select_all_products_change():
    set_all_product_selection(
        st.session_state.products,
        bool(st.session_state.get("select_all_products", False)),
    )


def product_from_dict(data: dict) -> Product:
    product_fields = {field.name: field for field in fields(Product)}
    values = {}
    for name, field in product_fields.items():
        if name in data:
            values[name] = data[name]
        elif field.default is not MISSING:
            values[name] = field.default
        else:
            values[name] = None
    product = Product(**values)
    product.selected = False
    return product


def safe_slug(value: str) -> str:
    slug = re.sub(r"[^A-Za-z0-9]+", "-", value or "").strip("-").lower()
    return slug[:60] or "collection"


def raw_products_payload(products: list[Product]) -> dict:
    return {
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "count": len(products),
        "products": [asdict(product) for product in products],
    }


def load_raw_products_payload(path: Path) -> tuple[list[Product], dict, str]:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as exc:
        return [], {}, f"读取采集结果失败：{exc}"
    products = [product_from_dict(item) for item in payload.get("products", []) if isinstance(item, dict)]
    return products, payload, ""


def load_raw_history_index() -> list[dict]:
    if not RAW_PRODUCTS_HISTORY_INDEX.exists():
        return []
    try:
        data = json.loads(RAW_PRODUCTS_HISTORY_INDEX.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return []
    if isinstance(data, dict):
        data = data.get("records", [])
    return [record for record in data if isinstance(record, dict)]


def write_raw_history_index(records: list[dict]) -> None:
    RAW_PRODUCTS_HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    RAW_PRODUCTS_HISTORY_INDEX.write_text(
        json.dumps({"records": records[:RAW_PRODUCTS_HISTORY_LIMIT]}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def save_raw_products(products: list[Product], label: str = "", source_url: str = "") -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    RAW_PRODUCTS_HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    payload = raw_products_payload(products)
    payload["label"] = label or "未命名采集"
    payload["source_url"] = source_url or ""
    RAW_PRODUCTS_CACHE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    history_name = f"{timestamp}_{safe_slug(label)}.json"
    history_path = RAW_PRODUCTS_HISTORY_DIR / history_name
    history_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    records = load_raw_history_index()
    records.insert(0, {
        "saved_at": payload["saved_at"],
        "count": len(products),
        "label": payload["label"],
        "source_url": payload["source_url"],
        "file": str(history_path),
    })
    kept_records = records[:RAW_PRODUCTS_HISTORY_LIMIT]
    kept_files = {record.get("file") for record in kept_records}
    for old_record in records[RAW_PRODUCTS_HISTORY_LIMIT:]:
        old_file = old_record.get("file")
        if old_file and old_file not in kept_files:
            try:
                Path(old_file).unlink(missing_ok=True)
            except OSError:
                pass
    write_raw_history_index(kept_records)


def load_raw_products() -> tuple[list[Product], str]:
    if not RAW_PRODUCTS_CACHE.exists():
        return [], "还没有上次采集结果。"
    products, payload, error = load_raw_products_payload(RAW_PRODUCTS_CACHE)
    if error:
        return [], error.replace("采集结果", "上次采集结果")
    saved_at = payload.get("saved_at", "-")
    return products, f"已载入上次原始采集池：{len(products)} 条，保存时间：{saved_at}。"


def raw_history_options() -> list[tuple[str, Path]]:
    options = []
    for record in load_raw_history_index():
        file_path = Path(record.get("file", ""))
        if not file_path.exists():
            continue
        saved_at = str(record.get("saved_at", "-")).replace("T", " ")
        label = record.get("label") or "未命名采集"
        count = record.get("count", 0)
        options.append((f"{saved_at}｜{label}｜{count} 条", file_path))
    return options


def load_sellersprite_image_cache() -> dict[str, str]:
    if not SELLERSPRITE_IMAGE_CACHE.exists():
        return {}
    try:
        data = json.loads(SELLERSPRITE_IMAGE_CACHE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}
    return {asin: url for asin, url in data.items() if asin and url}


def load_sellersprite_cache_meta() -> dict:
    if not SELLERSPRITE_META_CACHE.exists():
        return {}
    try:
        return json.loads(SELLERSPRITE_META_CACHE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def learned_category_links_mtime() -> float:
    try:
        return LEARNED_CATEGORY_LINKS.stat().st_mtime
    except OSError:
        return 0.0


@lru_cache(maxsize=4)
def load_learned_category_links_cached(mtime: float) -> dict:
    if not LEARNED_CATEGORY_LINKS.exists():
        return {}
    try:
        return json.loads(LEARNED_CATEGORY_LINKS.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}


def load_learned_category_links() -> dict:
    return deepcopy(load_learned_category_links_cached(learned_category_links_mtime()))


@lru_cache(maxsize=4)
def display_categories_cached(mtime: float) -> dict:
    categories = deepcopy(CATEGORIES)
    learned = load_learned_category_links_cached(mtime).get("categories", {})
    allowed_roots = set(categories)
    for path in learned:
        parts = [part.strip() for part in path.split(" > ") if part.strip()]
        if not parts or parts[0] not in allowed_roots:
            continue
        current = categories
        for part in parts:
            current.setdefault(part, {"zh": "", "count": 0, "children": {}})
            current[part].setdefault("children", {})
            node = current[part]
            current = node["children"]
    return categories


def display_categories() -> dict:
    return deepcopy(display_categories_cached(learned_category_links_mtime()))


def save_learned_category_links(seed_label: str, links) -> None:
    learned = load_learned_category_links()
    categories = learned.setdefault("categories", {})
    for link in links:
        title = str(getattr(link, "title", "") or "").strip()
        url = str(getattr(link, "url", "") or "").strip()
        if not title or not url:
            continue
        link_path = str(getattr(link, "path", "") or "").strip()
        storage_title = link_path or title
        if seed_label and not seed_label.startswith("http"):
            seed_parts = [part.strip() for part in seed_label.split(" > ") if part.strip()]
            if title in seed_parts:
                continue
            root = seed_parts[0] if seed_parts else ""
            root_children = set(CATEGORIES.get(root, {}).get("children", {}))
            if len(seed_parts) > 1 and title in root_children:
                continue
            if not storage_title.startswith(seed_label):
                storage_title = f"{seed_label} > {storage_title}"
        categories[storage_title] = {
            "title": title,
            "url": url,
            "source": seed_label,
            "is_leaf": bool(getattr(link, "is_leaf", True)),
            "node": str(getattr(link, "node", "") or ""),
            "depth": int(getattr(link, "depth", 0) or 0),
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
    LEARNED_CATEGORY_LINKS.parent.mkdir(parents=True, exist_ok=True)
    LEARNED_CATEGORY_LINKS.write_text(json.dumps(learned, ensure_ascii=False, indent=2), encoding="utf-8")
    load_learned_category_links_cached.cache_clear()
    learned_category_lookup.cache_clear()
    display_categories_cached.cache_clear()


def normalize_category_text(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.lower())


@lru_cache(maxsize=4)
def learned_category_lookup(mtime: float) -> tuple[dict[str, str], tuple[tuple[str, str], ...]]:
    learned = load_learned_category_links_cached(mtime).get("categories", {})
    exact = {}
    normalized = []
    for title, payload in learned.items():
        url = str(payload.get("url") or "")
        if not title or not url:
            continue
        exact[title] = url
        normalized.append((normalize_category_text(title), url))
    return exact, tuple(normalized)


def find_learned_category_url(path: str) -> str:
    exact, normalized_learned = learned_category_lookup(learned_category_links_mtime())
    if path in exact:
        return exact[path]
    leaf = path.split(" > ")[-1]
    candidates = [leaf]
    candidates.extend(CATEGORY_TITLE_ALIASES.get(path, []))
    normalized_candidates = [normalize_category_text(candidate) for candidate in candidates if candidate]
    if path.count(" > ") >= 1:
        return ""
    for normalized_title, url in normalized_learned:
        if any(candidate and (candidate in normalized_title or normalized_title in candidate) for candidate in normalized_candidates):
            return url
    return ""


def find_exact_category_url(path: str) -> str:
    return CATEGORY_BESTSELLER_URLS.get(path) or find_learned_category_url(path)


def find_category_seed_url(path: str) -> tuple[str, str]:
    direct_url = find_exact_category_url(path)
    if direct_url:
        return path, direct_url
    parts = path.split(" > ")
    for index in range(len(parts) - 1, 0, -1):
        ancestor = " > ".join(parts[:index])
        ancestor_url = find_exact_category_url(ancestor)
        if ancestor_url:
            return ancestor, ancestor_url
    return "", ""


def amazon_url_for_list_type(url: str, list_type: str) -> str:
    if not url:
        return url
    target_kind = "new-releases" if list_type == "New Releases" else "bestsellers"
    return re.sub(r"/gp/(?:bestsellers|new-releases)/", f"/gp/{target_kind}/", url, count=1)


def refresh_category_link_mapping(selected_paths: list[str], custom_url: str = "", progress_bar=None) -> tuple[int, int]:
    seed_candidates: list[tuple[str, str]] = []
    seen: set[str] = set()

    if custom_url:
        seed_candidates.append(("自定义链接", custom_url))
        seen.add(custom_url)

    paths = selected_paths or list(CATEGORIES.keys())
    for path in paths:
        _, seed_url = find_category_seed_url(path)
        if not seed_url:
            top = path.split(" > ")[0]
            seed_url = CATEGORY_BESTSELLER_URLS.get(top, "")
        if seed_url and seed_url not in seen:
            seed_candidates.append((path, seed_url))
            seen.add(seed_url)

    if not seed_candidates:
        raise ValueError("没有可用于刷新类目映射的 Amazon 榜单入口。")

    learned_count = 0
    failed_count = 0
    queue = list(seed_candidates)
    processed_urls: set[str] = set()
    index = 0
    while index < len(queue):
        label, seed_url = queue[index]
        index += 1
        if seed_url in processed_urls:
            continue
        processed_urls.add(seed_url)
        if progress_bar:
            progress_bar.progress(
                min(95, int(((index - 1) / max(len(queue), 1)) * 100)),
                text=f"正在准备类目链接 {index}/{len(queue)}：{label}",
            )
        try:
            links = discover_bestseller_category_links(seed_url, max_links=120)
            save_learned_category_links(label, links)
            learned_count += len(links)
            log(f"Category mapping refreshed: {label}, learned {len(links)} links.")

            for path in selected_paths:
                if find_exact_category_url(path):
                    continue
                parts = path.split(" > ")
                for depth in range(len(parts) - 1, 0, -1):
                    ancestor = " > ".join(parts[:depth])
                    ancestor_url = find_exact_category_url(ancestor)
                    if ancestor_url and ancestor_url not in processed_urls and all(ancestor_url != queued_url for _, queued_url in queue):
                        queue.append((ancestor, ancestor_url))
                    if ancestor_url:
                        break
        except Exception as exc:
            failed_count += 1
            log(f"Category mapping refresh failed: {label}. {exc}")

    if progress_bar:
        progress_bar.progress(100, text=f"类目链接准备完成：发现/更新 {learned_count} 个类目入口，失败 {failed_count} 个入口。")
        time.sleep(0.4)
        progress_bar.empty()
    return learned_count, failed_count


def sellersprite_cache_summary() -> str:
    meta = load_sellersprite_cache_meta()
    captured_at = meta.get("captured_at", "未知时间")
    source_url = meta.get("source_url", "未知链接")
    product_count = meta.get("product_count", 0)
    if not SELLERSPRITE_DOM_CACHE.exists():
        return "还没有卖家精灵插件数据，请先采集。"
    updated_at = datetime.fromtimestamp(SELLERSPRITE_DOM_CACHE.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
    return f"卖家精灵插件数据：{product_count} 条｜抓取时间：{captured_at or updated_at}｜来源：{source_url}"


def sellersprite_cache_hydration() -> tuple[int, int]:
    if not SELLERSPRITE_DOM_CACHE.exists():
        return 0, 0
    try:
        products = load_cached_sellersprite_products(limit=50)
    except Exception:
        return 0, 0
    hydrated = sum(1 for product in products if product.parent_monthly_sales or product.sales_amount or product.fba_fee)
    return len(products), hydrated


def sellersprite_cache_warning() -> str:
    if not SELLERSPRITE_DOM_CACHE.exists():
        return ""
    try:
        products = load_cached_sellersprite_products(limit=50)
    except Exception:
        return "卖家精灵插件数据存在，但解析失败。"
    if not products:
        return "卖家精灵插件数据里没有解析到产品。"
    hydrated = sum(1 for product in products if product.parent_monthly_sales or product.sales_amount or product.fba_fee)
    if hydrated < len(products):
        return f"卖家精灵插件字段未完全加载：共 {len(products)} 条产品，{hydrated} 条有销量/FBA/销售额字段。建议重新采集后再导出。"
    return ""


def sellersprite_cache_status_html(total: int, hydrated: int, chrome_ready: bool) -> str:
    meta = load_sellersprite_cache_meta()
    captured_at = str(meta.get("captured_at") or "未知时间").replace("T", " ")
    source_url = str(meta.get("source_url") or "")
    percent = int((hydrated / total) * 100) if total else 0
    status_text = "最近页面完整" if total and hydrated >= total else "等待插件加载"
    chrome_text = "Chrome 已连接" if chrome_ready else "Chrome 未连接"
    chrome_class = "ok" if chrome_ready else "warn"
    source_link = (
        f'<a href="{escape(source_url)}" target="_blank" rel="noopener noreferrer">来源链接</a>'
        if source_url
        else "<span>暂无来源</span>"
    )
    return f"""
    <div class="cache-card">
        <div class="cache-card-top">
            <div>
                <div class="cache-title">最近页面插件数据</div>
                <div class="cache-sub">抓取时间：{escape(captured_at)} · {source_link}</div>
            </div>
            <div class="cache-badges">
                <span class="cache-badge {chrome_class}">{chrome_text}</span>
                <span class="cache-badge">{status_text}</span>
            </div>
        </div>
        <div class="cache-progress"><span style="width:{percent}%"></span></div>
        <div class="cache-foot"><strong>{total}</strong> 条产品，<strong>{hydrated}</strong> 条插件字段完整</div>
    </div>
    """


def detect_risk(title: str, brand: str) -> str:
    text = f"{title} {brand}".lower()
    tags = [keyword for keyword in BLOCKED_KEYWORDS if keyword in text]
    return "; ".join(tags)


def score_product(rank, price, rating, reviews, bought, risks, min_price, max_price, max_reviews, min_bought):
    score = 50
    reasons = []

    if rank <= 20:
        score += 12
        reasons.append("top rank")
    if reviews <= max_reviews:
        score += 12
        reasons.append("low reviews")
    if bought >= min_bought:
        score += 14
        reasons.append("monthly bought signal")
    if min_price <= price <= max_price:
        score += 8
        reasons.append("good price band")
    if rating >= 4.2:
        score += 6
        reasons.append("healthy rating")
    if reviews > 1000:
        score -= 12
        reasons.append("high competition")
    if price < min_price:
        score -= 10
        reasons.append("low price")
    if rating < 4.0:
        score -= 10
        reasons.append("rating risk")
    if risks:
        score -= 30
        reasons.append("blocked/risk keyword")

    score = max(0, min(100, score))
    if risks:
        level = "Risk"
    elif score >= 82:
        level = "A"
    elif score >= 68:
        level = "B"
    elif score >= 52:
        level = "C"
    else:
        level = "D"
    return score, level, ", ".join(reasons) or "no strong signal"


def parse_filter_number(value: str, default=None, as_int: bool = False):
    text = str(value or "").replace(",", "").replace("$", "").strip()
    if not text:
        return default
    try:
        number = float(text)
    except ValueError:
        return default
    return int(number) if as_int else number


def in_optional_range(value, min_value=None, max_value=None) -> bool:
    if min_value is not None and value < min_value:
        return False
    if max_value is not None and value > max_value:
        return False
    return True


def has_active_range(min_value=None, max_value=None) -> bool:
    return min_value is not None or max_value is not None


def launched_at_matches(value: str, filter_value: str) -> bool:
    if filter_value in ("不限", "Any"):
        return True
    try:
        launched = datetime.strptime(str(value), "%Y-%m-%d").date()
    except ValueError:
        return False
    today = datetime.now().date()
    days = (today - launched).days
    ranges = {
        "近30天": (0, 30),
        "Last 30 days": (0, 30),
        "近60天": (0, 60),
        "Last 60 days": (0, 60),
        "近3个月": (0, 90),
        "Last 3 months": (0, 90),
        "近半年": (0, 183),
        "Last 6 months": (0, 183),
        "近1年": (0, 365),
        "Last year": (0, 365),
        "近2年": (0, 730),
        "Last 2 years": (0, 730),
        "近1~2年": (366, 730),
        "1-2 years": (366, 730),
    }
    min_days, max_days = ranges.get(filter_value, (0, 999999))
    return min_days <= days <= max_days


def product_matches_filters(product: Product, filters: dict) -> bool:
    if has_active_range(filters["min_child_sales"], filters["max_child_sales"]) and product.child_monthly_sales <= 0:
        return False
    if has_active_range(filters["min_bsr"], filters["max_bsr"]) and product.bsr_rank <= 0:
        return False
    return (
        in_optional_range(product.price, filters["min_price"], filters["max_price"])
        and in_optional_range(product.review_count, filters["min_reviews"], filters["max_reviews"])
        and in_optional_range(product.monthly_bought, filters["min_bought"], filters["max_bought"])
        and in_optional_range(product.child_monthly_sales, filters["min_child_sales"], filters["max_child_sales"])
        and in_optional_range(product.bsr_rank, filters["min_bsr"], filters["max_bsr"])
        and launched_at_matches(product.launched_at, filters["launch_window"])
    )


def apply_product_filters(products: list[Product], filters: dict) -> list[Product]:
    return [product for product in products if product_matches_filters(product, filters)]


def build_collection_summary(raw_products: list[Product], filtered_products: list[Product], filters: dict) -> str:
    if not raw_products:
        return "当前没有原始采集结果。请先点击“开始采集”，或载入上次采集结果。"
    removed_count = len(raw_products) - len(filtered_products)
    summary = (
        f"原始采集池 {len(raw_products)} 条，当前筛选保留 {len(filtered_products)} 条，"
        f"筛掉 {removed_count} 条。"
    )
    rejection_lines = filter_rejection_summary(raw_products, filters)[:5]
    if rejection_lines:
        summary += " 主要筛选原因：" + "；".join(rejection_lines)
    return summary


def apply_filters_to_raw_pool(filters: dict) -> None:
    selected_by_asin = {product.asin: product.selected for product in st.session_state.products}
    filtered_products = apply_product_filters(st.session_state.raw_products, filters)
    for product in filtered_products:
        product.selected = bool(selected_by_asin.get(product.asin, product.selected))
    st.session_state.products = filtered_products
    st.session_state.last_collection_summary = build_collection_summary(
        st.session_state.raw_products,
        st.session_state.products,
        filters,
    )


def filter_rejection_summary(products: list[Product], filters: dict) -> list[str]:
    reasons = {
        "价格低于最低值": 0,
        "价格高于最高值": 0,
        "评分数低于最低值": 0,
        "评分数高于最高值": 0,
        "月销量低于最低值": 0,
        "月销量高于最高值": 0,
        "子体销量不在范围": 0,
        "BSR 不在范围": 0,
        "上架时间不符合": 0,
    }
    for product in products:
        if filters["min_price"] is not None and product.price < filters["min_price"]:
            reasons["价格低于最低值"] += 1
        if filters["max_price"] is not None and product.price > filters["max_price"]:
            reasons["价格高于最高值"] += 1
        if filters["min_reviews"] is not None and product.review_count < filters["min_reviews"]:
            reasons["评分数低于最低值"] += 1
        if filters["max_reviews"] is not None and product.review_count > filters["max_reviews"]:
            reasons["评分数高于最高值"] += 1
        if filters["min_bought"] is not None and product.monthly_bought < filters["min_bought"]:
            reasons["月销量低于最低值"] += 1
        if filters["max_bought"] is not None and product.monthly_bought > filters["max_bought"]:
            reasons["月销量高于最高值"] += 1
        if not in_optional_range(product.child_monthly_sales, filters["min_child_sales"], filters["max_child_sales"]):
            reasons["子体销量不在范围"] += 1
        if not in_optional_range(product.bsr_rank, filters["min_bsr"], filters["max_bsr"]):
            reasons["BSR 不在范围"] += 1
        if not launched_at_matches(product.launched_at, filters["launch_window"]):
            reasons["上架时间不符合"] += 1
    return [f"{name}：{count} 条" for name, count in sorted(reasons.items(), key=lambda item: item[1], reverse=True) if count]


def render_range_filter(title: str, key_prefix: str, min_default: str = "", max_default: str = "", money: bool = False):
    st.markdown(f"<div class='filter-label'>{escape(title)} <span>?</span></div>", unsafe_allow_html=True)
    st.markdown("<div class='range-filter-anchor'></div>", unsafe_allow_html=True)
    cols = st.columns([1, 0.18, 1], vertical_alignment="center")
    with cols[0]:
        min_value = st.text_input(
            f"{title} 最小值",
            value=min_default,
            placeholder="最小值",
            key=f"{key_prefix}_min",
            label_visibility="collapsed",
        )
    with cols[1]:
        st.markdown("<div class='filter-range-sep'>~</div>", unsafe_allow_html=True)
    with cols[2]:
        max_value = st.text_input(
            f"{title} 最大值",
            value=max_default,
            placeholder="最大值",
            key=f"{key_prefix}_max",
            label_visibility="collapsed",
        )
    if money:
        st.markdown("<div class='filter-money-hint'>$</div>", unsafe_allow_html=True)
    return min_value, max_value


def reset_filter_widgets() -> None:
    defaults = {
        "filter_price_min": "",
        "filter_price_max": "",
        "filter_reviews_min": "",
        "filter_reviews_max": "",
        "filter_monthly_sales_min": "",
        "filter_monthly_sales_max": "",
        "filter_child_sales_min": "",
        "filter_child_sales_max": "",
        "filter_bsr_min": "",
        "filter_bsr_max": "",
        "filter_launch_window": "不限" if st.session_state.get("ui_lang", "中文") == "中文" else "Any",
    }
    for key, value in defaults.items():
        st.session_state[key] = value


def fake_image(seed: str) -> str:
    colors = ["f4a261", "2a9d8f", "e76f51", "8ab17d", "577590", "e9c46a"]
    color = random.choice(colors)
    return f"https://placehold.co/320x320/{color}/ffffff?text={seed}"


def make_fake_product(category_path, list_type, rank, filters) -> Product:
    asin = "B0" + "".join(random.choices("ABCDEFGHJKLMNPQRSTUVWXYZ23456789", k=8))
    leaf = category_path.split(" > ")[-1]
    title = random.choice(
        [
            f"Upgraded {leaf} with Easy Clean Design",
            f"Portable {leaf} Kit for Small Spaces",
            f"Compact {leaf} Set for Families",
            f"Smart {leaf} Accessory for Daily Use",
            f"Foldable {leaf} Organizer with Travel Case",
        ]
    )
    brand = random.choice(["Novaly", "Petuno", "Auralis", "Mavento", "NorthPeak", "Amazon Basics"])
    price = round(random.uniform(12.99, 229.99), 2)
    rating = round(random.uniform(3.6, 4.9), 1)
    reviews = random.choice([12, 28, 43, 76, 128, 216, 304, 640, 1210])
    bought = random.choice([50, 100, 200, 300, 500, 1000, 2000])
    risks = detect_risk(title, brand)
    score, level, reason = score_product(
        rank,
        price,
        rating,
        reviews,
        bought,
        risks,
        filters["min_price"],
        filters["max_price"],
        filters["max_reviews"],
        filters["min_bought"],
    )
    return Product(
        selected=False,
        review_status="Pending",
        note="",
        list_type=list_type,
        category_path=category_path,
        rank=rank,
        title=title,
        asin=asin,
        amazon_url=f"https://www.amazon.com/dp/{asin}",
        image_url=fake_image(asin),
        price=price,
        rating=rating,
        review_count=reviews,
        monthly_bought=bought,
        brand=brand,
        prime_fba=random.choice(["Prime/FBA", "Prime", "Unknown"]),
        delivery=random.choice(["Free delivery", "Ships from Amazon", "Unknown"]),
        variant_count=random.randint(1, 8),
        launched_at=random.choice(["2026-04-18", "2026-02-09", "2025-11-22", "Unknown"]),
        package_dimensions=random.choice(["8.3 x 5.4 x 2.1 in", "12.0 x 9.2 x 4.8 in", "15.7 x 11.0 x 6.5 in"]),
        package_weight_lb=round(random.uniform(0.4, 8.5), 2),
        fba_fee=round(random.uniform(3.28, 18.75), 2),
        potential_score=score,
        potential_level=level,
        reason=reason,
        risk_tags=risks,
        scraped_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        status="OK",
        error="",
    )


def generate_products(list_type, selected_paths, pages, filters):
    products = []
    for path in selected_paths:
        log(f"Generate demo data: {list_type} / {path}")
        for page in range(1, pages + 1):
            for rank in range(1, 13):
                products.append(make_fake_product(path, list_type, (page - 1) * 12 + rank, filters))
    return products


def make_product_from_scraped(scraped, list_type, category_path, filters) -> Product:
    risks = detect_risk(scraped.title, scraped.brand)
    price = scraped.price or 0.0
    rating = scraped.rating or 0.0
    reviews = scraped.review_count or 0
    bought = scraped.monthly_bought or 0
    score, level, reason = score_product(
        scraped.rank,
        price,
        rating,
        reviews,
        bought,
        risks,
        filters["min_price"],
        filters["max_price"],
        filters["max_reviews"],
        filters["min_bought"],
    )
    return Product(
        selected=False,
        review_status="Pending",
        note="",
        list_type=list_type,
        category_path=category_path,
        rank=scraped.rank,
        title=scraped.title,
        asin=scraped.asin,
        amazon_url=scraped.amazon_url,
        image_url=scraped.image_url or fake_image(scraped.asin),
        price=price,
        rating=rating,
        review_count=reviews,
        monthly_bought=bought,
        brand=scraped.brand,
        prime_fba="Unknown",
        delivery="Unknown",
        variant_count=0,
        launched_at="Unknown",
        package_dimensions="Unknown",
        package_weight_lb=0.0,
        fba_fee=0.0,
        potential_score=score,
        potential_level=level,
        reason=reason,
        risk_tags=risks,
        scraped_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        status="OK",
        error="",
    )


def collect_real_products(list_type, url, filters) -> list[Product]:
    scraped_products = fetch_best_sellers(url or DEFAULT_TEST_URL, limit=30)
    return [
        make_product_from_scraped(
            scraped,
            list_type,
            "Pet Supplies > Dogs > Dog Feeding & Watering Supplies",
            filters,
        )
        for scraped in scraped_products
    ]


def csv_bytes(products):
    if not products:
        return b""
    output = StringIO()
    headers = [header for _, header in SELLERSPRITE_EXPORT_COLUMNS]
    writer = csv.DictWriter(output, fieldnames=headers)
    writer.writeheader()
    for product in products:
        row = asdict(product)
        row["image_preview_formula"] = image_formula(row.get("image_url", ""))
        writer.writerow({header: row.get(field, "") for field, header in SELLERSPRITE_EXPORT_COLUMNS})
    return output.getvalue().encode("utf-8-sig")


def image_formula(image_url: str) -> str:
    if not image_url:
        return ""
    escaped_url = str(image_url).replace('"', '""')
    return f'=IMAGE("{escaped_url}","",3,50,50)'


def excel_bytes(products):
    if not products:
        return b""
    fields = [field for field, _ in SELLERSPRITE_EXPORT_COLUMNS]
    headers = [header for _, header in SELLERSPRITE_EXPORT_COLUMNS]
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "卖家精灵采集"
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True

    sheet.append(headers)
    header_fill = PatternFill("solid", fgColor="F6F7F9")
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="5F6673")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    for row_number, product in enumerate(products, start=2):
        row = asdict(product)
        row["image_preview_formula"] = image_formula(row.get("image_url", ""))
        sheet.append([row.get(field, "") for field in fields])
        sheet.row_dimensions[row_number].height = 42

    width_by_field = {
        "image_preview_formula": 10,
        "rank": 6,
        "title": 34,
        "asin": 16,
        "brand": 14,
        "seller_name": 18,
        "bsr_category": 18,
        "sub_category": 22,
        "category_path": 42,
        "amazon_url": 36,
        "image_url": 36,
        "reason": 42,
        "risk_tags": 20,
        "package_dimensions": 28,
        "scraped_at": 20,
        "error": 24,
    }
    money_fields = {"sales_amount", "price", "fba_fee"}
    integer_fields = {"rank", "seller_count", "bsr_rank", "sub_rank", "monthly_bought", "variant_count", "review_count", "potential_score"}
    for column_index, (field, header) in enumerate(SELLERSPRITE_EXPORT_COLUMNS, start=1):
        width = width_by_field.get(field, min(max(len(header) + 4, 12), 20))
        letter = get_column_letter(column_index)
        sheet.column_dimensions[letter].width = width
        for cell in sheet[letter][1:]:
            cell.alignment = Alignment(vertical="center", wrap_text=field in {"title", "reason", "category_path", "package_dimensions"})
            if field in money_fields:
                cell.number_format = '$#,##0.00'
            elif field in integer_fields:
                cell.number_format = '#,##0'
    for column_index, (field, _) in enumerate(SELLERSPRITE_EXPORT_COLUMNS, start=1):
        if field in {"amazon_url", "image_url"}:
            for row_number in range(2, len(products) + 2):
                cell = sheet.cell(row=row_number, column=column_index)
                if cell.value:
                    cell.hyperlink = cell.value
                    cell.style = "Hyperlink"

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def level_badge(level: str):
    colors = {
        "A": "#0f766e",
        "B": "#2563eb",
        "C": "#a16207",
        "D": "#6b7280",
        "Risk": "#b91c1c",
    }
    return f"<span style='background:{colors.get(level, '#6b7280')}; color:white; padding:3px 8px; border-radius:6px; font-size:12px;'>{level}</span>"


def render_product_identity(product: Product):
    return None


def category_label(name: str, node: dict, indent: int = 0) -> str:
    zh = f" ({node.get('zh', '')})" if node.get("zh") else ""
    return f"{name}{zh}"


def leaf_label(name: str, node: dict, indent: int = 0) -> str:
    prefix = "- " if indent else ""
    return prefix + category_label(name, node, indent)


def category_matches_filter(name: str, node: dict, query: str) -> bool:
    if not query:
        return True
    haystack = f"{name} {node.get('zh', '')}".lower()
    if query in haystack:
        return True
    return any(category_matches_filter(child_name, child_node, query) for child_name, child_node in node.get("children", {}).items())


def category_widget_key(prefix: str, path: str) -> str:
    return f"{prefix}_{re.sub(r'[^A-Za-z0-9]+', '_', path).strip('_')}"


def iter_category_paths(tree: dict | None = None, prefix: str = "", seen: set[int] | None = None):
    source_tree = display_categories() if tree is None else tree
    if not isinstance(source_tree, dict) or not source_tree:
        return
    seen = seen or set()
    tree_id = id(source_tree)
    if tree_id in seen:
        return
    seen.add(tree_id)
    for name, node in source_tree.items():
        path = f"{prefix} > {name}" if prefix else name
        yield path
        yield from iter_category_paths(node.get("children", {}), path, seen)


def handle_category_select_all_change():
    selected = bool(st.session_state.get("category_select_all", False))
    for path in iter_category_paths():
        st.session_state[category_widget_key("cat_sel", path)] = selected


def category_path_parts(path: str) -> list[str]:
    return [part.strip() for part in path.split(">") if part.strip()]


def get_category_node(path: str, tree: dict | None = None) -> dict:
    node = None
    current = display_categories() if tree is None else tree
    for part in category_path_parts(path):
        node = current.get(part, {})
        current = node.get("children", {}) if isinstance(node, dict) else {}
    return node if isinstance(node, dict) else {}


def iter_child_paths(path: str, node: dict):
    for child_name, child_node in node.get("children", {}).items():
        child_path = f"{path} > {child_name}"
        yield child_path, child_node


def iter_category_branch_paths(path: str, node: dict):
    yield path
    for child_path, child_node in iter_child_paths(path, node):
        yield from iter_category_branch_paths(child_path, child_node)


def iter_category_leaf_paths(path: str, node: dict):
    children = list(iter_child_paths(path, node))
    if not children:
        yield path
        return
    for child_path, child_node in children:
        yield from iter_category_leaf_paths(child_path, child_node)


def set_category_branch_selected(path: str, node: dict, selected: bool):
    for branch_path in iter_category_branch_paths(path, node):
        st.session_state[category_widget_key("cat_sel", branch_path)] = selected


def set_category_descendants_selected(path: str, node: dict, selected: bool):
    set_category_branch_selected(path, node, selected)


def category_branch_fully_selected(path: str, node: dict) -> bool:
    if not st.session_state.get(category_widget_key("cat_sel", path), False):
        return False
    return all(category_branch_fully_selected(child_path, child_node) for child_path, child_node in iter_child_paths(path, node))


def sync_category_ancestors(path: str):
    parts = category_path_parts(path)
    for depth in range(len(parts) - 1, 0, -1):
        parent_path = " > ".join(parts[:depth])
        parent_node = get_category_node(parent_path)
        children = list(iter_child_paths(parent_path, parent_node))
        if not children:
            continue
        parent_selected = all(category_branch_fully_selected(child_path, child_node) for child_path, child_node in children)
        st.session_state[category_widget_key("cat_sel", parent_path)] = parent_selected


def sync_category_select_all_state():
    all_paths = list(iter_category_paths())
    st.session_state["category_select_all"] = bool(all_paths) and all(
        st.session_state.get(category_widget_key("cat_sel", path), False) for path in all_paths
    )


def handle_category_row_select_change(path: str, node: dict):
    selected = bool(st.session_state.get(category_widget_key("cat_sel", path), False))
    set_category_branch_selected(path, node, selected)
    sync_category_ancestors(path)
    sync_category_select_all_state()


def selected_category_paths_from_state(tree: dict | None = None, prefix: str = "", seen: set[int] | None = None) -> list[str]:
    source_tree = display_categories() if tree is None else tree
    if not isinstance(source_tree, dict) or not source_tree:
        return []
    seen = seen or set()
    tree_id = id(source_tree)
    if tree_id in seen:
        return []
    seen.add(tree_id)
    selected_paths = []
    selected_seen = set()
    for name, node in source_tree.items():
        path = f"{prefix} > {name}" if prefix else name
        if st.session_state.get(category_widget_key("cat_sel", path), False):
            for branch_path in iter_category_branch_paths(path, node):
                if branch_path not in selected_seen:
                    selected_paths.append(branch_path)
                    selected_seen.add(branch_path)
            continue
        for branch_path in selected_category_paths_from_state(node.get("children", {}), path, seen):
            if branch_path not in selected_seen:
                selected_paths.append(branch_path)
                selected_seen.add(branch_path)
    return selected_paths


def render_category_row(path: str, name: str, node: dict, depth: int, selected_paths: list[str], all_categories: bool, confirmed: set[str], query: str):
    children = node.get("children", {})
    has_children = bool(children)
    expand_key = category_widget_key("cat_exp", path)
    select_key = category_widget_key("cat_sel", path)
    if expand_key not in st.session_state:
        st.session_state[expand_key] = False
    if select_key not in st.session_state:
        st.session_state[select_key] = all_categories or path in confirmed

    if all_categories:
        st.session_state[select_key] = True

    indent = max(0.001, min(depth, 4) * 0.055)
    label_width = max(0.52, 0.72 - indent)
    cols = st.columns([indent, 0.04, 0.045, label_width, 0.19], gap=None, vertical_alignment="center")
    with cols[0]:
        if depth:
            st.markdown(f"<span class='category-tree-indent depth-{depth}'></span>", unsafe_allow_html=True)
        else:
            st.markdown("<span class='category-tree-root-indent'></span>", unsafe_allow_html=True)
    with cols[1]:
        if has_children:
            arrow = "▾" if st.session_state[expand_key] else "▸"
            if st.button(arrow, key=f"toggle_{expand_key}", use_container_width=True):
                st.session_state[expand_key] = not st.session_state[expand_key]
        else:
            st.markdown("<span class='category-tree-spacer'></span>", unsafe_allow_html=True)
    with cols[2]:
        selected = st.checkbox(
            f"选择 {path}",
            key=select_key,
            on_change=handle_category_row_select_change,
            args=(path, node),
            label_visibility="collapsed",
        )
    with cols[3]:
        label = category_label(name, node, depth)
        st.markdown(f"<div class='category-tree-label'>{escape(label)}</div>", unsafe_allow_html=True)
    with cols[4]:
        count = int(node.get("count") or 0)
        if count:
            st.markdown(f"<div class='category-count-badge'>{count:,}</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div></div>", unsafe_allow_html=True)

    if selected:
        selected_paths.append(path)

    if has_children and (st.session_state[expand_key] or query):
        for child_name, child_node in children.items():
            if query and not category_matches_filter(child_name, child_node, query):
                continue
            render_category_row(
                f"{path} > {child_name}",
                child_name,
                child_node,
                depth + 1,
                selected_paths,
                all_categories,
                confirmed,
                query,
            )


def render_category_tree(query: str = ""):
    selected_paths = []
    all_categories = st.checkbox(
        T["all_categories"],
        key="category_select_all",
        on_change=handle_category_select_all_change,
    )
    confirmed = set(st.session_state.confirmed_category_paths)
    for main, main_node in display_categories().items():
        if not category_matches_filter(main, main_node, query):
            continue
        render_category_row(main, main, main_node, 0, selected_paths, all_categories, confirmed, query)
    return selected_category_paths_from_state()


@st.dialog("选择类目", width="large")
def render_category_dialog():
    if st.session_state.get("category_clear_requested"):
        st.session_state.confirmed_category_paths = []
        for key in list(st.session_state.keys()):
            if key.startswith(("main_", "mid_", "leaf_", "cat_sel_", "cat_exp_")) or key == "category_select_all":
                del st.session_state[key]
        st.session_state.category_clear_requested = False

    left_panel, right_panel = st.columns([0.50, 0.50], gap="large", vertical_alignment="top")
    with left_panel:
        query = st.text_input(
            "搜索类目",
            key="category_search",
            placeholder="请输入 Node ID/类目关键词，如281407/Electronics",
            label_visibility="collapsed",
        ).strip().lower()
        with st.container(height=365, border=False):
            selected_paths = render_category_tree(query)
    with right_panel:
        title_col, clear_col = st.columns([0.72, 0.28], vertical_alignment="center")
        title_col.markdown(f"<div class='category-selected-title'>已选（{len(selected_paths)}）</div>", unsafe_allow_html=True)
        clear_clicked = clear_col.button("清空", use_container_width=True)
        if selected_paths:
            selected_html = "".join(f"<span class='selected-pill'>{escape(path)}</span>" for path in selected_paths[:24])
            if len(selected_paths) > 24:
                selected_html += f"<span class='selected-pill'>+{len(selected_paths) - 24}</span>"
            st.markdown(f"<div class='category-selected-panel'>{selected_html}</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='category-selected-empty'>暂未选择类目</div>", unsafe_allow_html=True)
        if clear_clicked:
            st.session_state.category_clear_requested = True
            st.rerun()
    st.markdown("<span class='category-footer-anchor'></span>", unsafe_allow_html=True)
    spacer_left, cancel_col, spacer_mid, confirm_col = st.columns([0.76, 0.10, 0.02, 0.12])
    if cancel_col.button("取消", use_container_width=True):
        st.session_state.show_category_dialog = False
        st.rerun()
    if confirm_col.button("确认选择", type="primary", use_container_width=True):
        st.session_state.confirmed_category_paths = selected_paths
        st.session_state.show_category_dialog = False
        st.rerun()


def _display_dash(value, suffix=""):
    if value in (None, "", "Unknown"):
        return "-"
    return f"{value}{suffix}"


def _display_money(value):
    return "$0.00" if not value else f"${value:,.2f}"


def _display_int(value, suffix=""):
    return "0" if not value else f"{value:,}{suffix}"


def _copy_js(value: str) -> str:
    payload = json.dumps(value or "")
    return (
        f"navigator.clipboard.writeText({payload}).then(()=>{{"
        "this.textContent='✓';this.classList.add('copied');"
        "setTimeout(()=>{this.textContent='⧉';this.classList.remove('copied');},1200);"
        "});"
    )


def render_clipboard_bridge():
    components.html(
        """
        <script>
        window.parent.eval(`
          (() => {
            document.querySelectorAll('.copy-icon[data-copy-value]').forEach((button) => {
              if (button.dataset.copyBound === '1') return;
              button.dataset.copyBound = '1';
              button.addEventListener('click', async (event) => {
                event.preventDefault();
                event.stopPropagation();
                const value = button.dataset.copyValue || '';
                if (!value) return;
                const oldText = button.textContent;
                const oldTitle = button.title;
                button.textContent = '✓';
                button.classList.add('copied');
                button.title = '已复制';
                setTimeout(() => {
                  button.textContent = oldText || '⧉';
                  button.classList.remove('copied');
                  button.title = oldTitle || '复制';
                }, 1200);
                try {
                  await navigator.clipboard.writeText(value);
                } catch (error) {
                  const input = document.createElement('textarea');
                  input.value = value;
                  input.style.position = 'fixed';
                  input.style.left = '-9999px';
                  document.body.appendChild(input);
                  input.select();
                  document.execCommand('copy');
                  input.remove();
                }
              });
            });
          })();
        `);
        </script>
        """,
        height=0,
    )


def seller_product_html(product: Product) -> str:
    title = escape(product.title)
    asin = escape(product.asin)
    brand = escape(product.brand or "-")
    image_url = escape(product.image_url or fake_image(product.asin))
    amazon_url = escape(product.amazon_url or f"https://www.amazon.com/dp/{product.asin}")
    category_path = escape(product.category_path)
    leaf_category = escape(product.category_path.split(" > ")[-1])
    monthly_bought = _display_int(product.monthly_bought, "+")
    sales_amount_value = getattr(product, "sales_amount", 0)
    sales_amount = "-" if not sales_amount_value else f"${sales_amount_value:,.0f}"
    child_sales_label = getattr(product, "child_monthly_sales_label", "")
    child_sales = child_sales_label or _display_int(getattr(product, "child_monthly_sales", 0), "+")
    bsr_rank = getattr(product, "bsr_rank", 0) or product.rank
    sub_rank = getattr(product, "sub_rank", 0)
    bsr_category = getattr(product, "bsr_category", "")
    sub_category = getattr(product, "sub_category", "")
    qa_count = "-" if not product.review_count else f"{max(12, product.review_count // 3):,}"
    review_new = "-" if not product.review_count else f"{max(8, product.review_count // 4):,}"
    margin = product.margin_rate or ("0" if not product.fba_fee else f"{min(45, max(12, int(product.potential_score / 3)))}%")
    delivery = product.delivery or product.prime_fba or "0"
    package_weight = _display_dash(product.package_weight_lb, " pounds")
    package_dimensions = _display_dash(product.package_dimensions)
    return f"""
    <div class="seller-row">
        <div class="seller-main">
            <div class="seller-rank">{product.rank}</div>
            <div class="seller-product">
                <div class="seller-image-wrap">
                    <span class="level-corner">{escape(product.potential_level)}</span>
                    <img src="{image_url}" alt="{title}" />
                    <div class="signal-tags"><span class="tag tag-bs">BS</span><span class="tag tag-ac">AC</span><span class="tag tag-nr">NR</span></div>
                </div>
                <div class="seller-info">
                    <div class="seller-title">{title}</div>
                    <div class="meta-line">ASIN: <strong>{asin}</strong><button type="button" class="copy-icon" data-copy-value="{asin}" title="复制 ASIN">⧉</button><a class="mini-link" href="{amazon_url}" target="_blank" rel="noopener noreferrer" title="打开 Amazon DP 链接">↗</a></div>
                    <div class="meta-line muted">父ASIN：- <button type="button" class="copy-icon disabled-icon" title="暂无父ASIN">⧉</button><a class="mini-link" href="{amazon_url}" target="_blank" rel="noopener noreferrer" title="打开 Amazon DP 链接">↗</a></div>
                    <div class="meta-line">品牌:<strong>{brand}</strong><button type="button" class="copy-icon" data-copy-value="{brand}" title="复制品牌">⧉</button><a class="mini-link" href="{amazon_url}" target="_blank" rel="noopener noreferrer" title="打开 Amazon DP 链接">↗</a></div>
                </div>
            </div>
            <div class="cell bsr-cell"><strong>{_display_int(bsr_rank)}</strong><span class="green">↑ -</span><span class="green">↑ -</span></div>
            <div class="trend-cell"><svg viewBox="0 0 150 64" preserveAspectRatio="none"><path d="M0 42 L18 42 L32 36 L44 22 L58 34 L72 24 L88 39 L104 32 L122 20 L150 24" fill="none" stroke="#ff8a1f" stroke-width="2"/><path d="M0 42 L18 42 L32 36 L44 22 L58 34 L72 24 L88 39 L104 32 L122 20 L150 24 L150 64 L0 64 Z" fill="#fff1df"/></svg></div>
            <div class="cell"><strong>{monthly_bought}</strong><span class="muted">0</span></div>
            <div class="cell"><strong>{sales_amount}</strong></div>
            <div class="cell"><strong>{child_sales}</strong><span class="muted">0</span></div>
            <div class="cell"><strong>{_display_dash(product.variant_count)}</strong></div>
            <div class="cell"><strong>{_display_money(product.price)}</strong><span class="muted">{qa_count}</span></div>
            <div class="cell"><strong>{_display_int(product.review_count)}</strong><span class="muted">{review_new}</span></div>
            <div class="cell"><strong>{_display_dash(product.rating)}</strong><span class="muted">-</span></div>
            <div class="cell"><strong>{_display_money(product.fba_fee)}</strong><span class="muted">{margin}</span></div>
            <div class="cell"><strong>{escape(str(_display_dash(product.launched_at)))}</strong><span class="muted">-</span></div>
            <div class="cell"><strong>{delivery}</strong><span class="muted">-</span></div>
            <div class="ops-cell"><span>▥</span><span>⊙</span><span>⊕</span><span>▦</span><span>⋯</span></div>
        </div>
        <div class="seller-detail">
            <div>浏览同类目: <span class="orange">{escape(bsr_category or category_path)}</span> <span class="pill orange-pill">BS榜单</span> <span class="pill orange-pill">新品榜</span> <span class="pill orange-pill">市场分析</span> <span class="pill orange-pill">找相似</span></div>
            <div>中文类目名: - <span class="rank-pill">#{_display_int(sub_rank) if sub_rank else 1}</span> in {escape(sub_category or leaf_category)}</div>
            <div>LQS: <strong>0</strong>　卖家: <strong>{escape(product.seller_name or '0')}</strong>　BuyBox卖家: <strong>{escape(product.seller_name or '0')}</strong>　商品重量: <strong>{package_weight}</strong>　商品尺寸: <strong>{escape(str(package_dimensions))}</strong>　包装重量: <strong>{package_weight}</strong>　包装尺寸: <strong>{escape(str(package_dimensions))}</strong></div>
        </div>
    </div>
    """


def collect_sellersprite_products(list_type, filters) -> list[Product]:
    scraped_products = load_cached_sellersprite_products(limit=200)
    image_cache = load_sellersprite_image_cache()
    products = []
    for scraped in scraped_products:
        risks = detect_risk(scraped.title, scraped.brand)
        score, level, reason = score_product(
            scraped.rank,
            scraped.price,
            scraped.rating,
            scraped.review_count,
            scraped.parent_monthly_sales,
            risks,
            filters["min_price"],
            filters["max_price"],
            filters["max_reviews"],
            filters["min_bought"],
        )
        product = Product(
            selected=False,
            review_status="待审核",
            note="",
            list_type=list_type,
            category_path=scraped.sub_category or scraped.bsr_category or "Pet Supplies > Dogs > Dog Wireless Fences",
            rank=scraped.rank,
            title=scraped.title,
            asin=scraped.asin,
            amazon_url=f"https://www.amazon.com/dp/{scraped.asin}",
            image_url=image_cache.get(scraped.asin, fake_image(scraped.asin)),
            price=scraped.price,
            rating=scraped.rating,
            review_count=scraped.review_count,
            monthly_bought=scraped.parent_monthly_sales,
            brand=scraped.brand,
            prime_fba=scraped.fulfillment or "Unknown",
            delivery=scraped.fulfillment or "Unknown",
            variant_count=scraped.variant_count,
            launched_at=scraped.launched_at or "Unknown",
            package_dimensions=scraped.package_dimensions or "Unknown",
            package_weight_lb=scraped.package_weight_lb,
            fba_fee=scraped.fba_fee,
            potential_score=score,
            potential_level=level,
            reason=reason,
            risk_tags=risks,
            scraped_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            status="OK",
            error="",
            seller_name=scraped.seller,
            seller_count=scraped.seller_count,
            sales_amount=scraped.sales_amount,
            child_monthly_sales=scraped.child_monthly_sales,
            child_monthly_sales_label=scraped.child_monthly_sales_label,
            bsr_rank=scraped.bsr_rank,
            bsr_category=scraped.bsr_category,
            sub_rank=scraped.sub_rank,
            sub_category=scraped.sub_category,
            margin_rate=scraped.margin_rate,
        )
        products.append(product)
    return products


def collect_sellersprite_entry(
    target_url: str,
    list_type: str,
    filters: dict,
    progress=None,
    page_count: int = 2,
    progress_label: str = "",
) -> tuple[list[Product], list]:
    products_by_asin: dict[str, Product] = {}
    duplicate_pages = 0
    refresh_results = []

    def update_progress(percent: int, message: str):
        if progress:
            progress(percent, f"{message}｜累计 {len(products_by_asin)} 条")

    def collect_page_products(page: int, refresh_result):
        nonlocal duplicate_pages
        refresh_results.append(refresh_result)
        if not refresh_result.ok and not refresh_result.product_count:
            raise RuntimeError(refresh_result.message)
        page_name = "第一页" if page == 1 else "第二页" if page == 2 else f"第 {page} 页"
        parsed_products = collect_sellersprite_products(list_type, filters)
        before_count = len(products_by_asin)
        for product in parsed_products:
            products_by_asin.setdefault(product.asin, product)
        added_count = len(products_by_asin) - before_count
        if progress:
            progress(
                99,
                f"{page_name}｜本页解析 {len(parsed_products)} 条，新增 {added_count} 条，累计 {len(products_by_asin)} 条",
            )
        if page > 1 and parsed_products and added_count == 0:
            duplicate_pages += 1
            log(f"{progress_label or target_url}: {page_name} parsed {len(parsed_products)} rows but added 0 unique ASINs. Pagination may be duplicated.")

    refresh_sellersprite_cache_pages(
        target_url,
        SELLERSPRITE_DOM_CACHE,
        SELLERSPRITE_IMAGE_CACHE,
        SELLERSPRITE_META_CACHE,
        expected_products=SELLERSPRITE_EXPECTED_PRODUCTS_PER_PAGE,
        page_count=page_count,
        progress=update_progress,
        page_callback=collect_page_products,
    )
    if duplicate_pages:
        log(f"{progress_label or target_url}: {duplicate_pages} duplicated page(s) detected during collection.")
    return list(products_by_asin.values()), refresh_results


def sellersprite_collection_quality(products: list[Product], refresh_results: list) -> tuple[bool, str]:
    if not refresh_results:
        return False, "没有读取到页面"
    page_counts = [max(int(getattr(result, "product_count", 0) or 0), int(getattr(result, "hydrated_count", 0) or 0)) for result in refresh_results]
    page_detail = "，".join(f"第{i + 1}页 {count} 条" for i, count in enumerate(page_counts))
    if len(refresh_results) < 2:
        return False, f"只读取到 {len(refresh_results)} 页（{page_detail}）"
    weak_pages = [
        f"第{i + 1}页 {count} 条"
        for i, count in enumerate(page_counts[:2])
        if count < SELLERSPRITE_MIN_PRODUCTS_PER_PAGE
    ]
    if weak_pages:
        return False, f"页面产品数不足：{'，'.join(weak_pages)}；预期每页接近 {SELLERSPRITE_EXPECTED_PRODUCTS_PER_PAGE} 条"
    if len(products) < SELLERSPRITE_MIN_PRODUCTS_TWO_PAGES:
        return False, f"两页去重后 {len(products)} 条，低于预期接近 100 条（{page_detail}）"
    return True, f"两页采集正常：{len(products)} 条（{page_detail}）"


def collect_sellersprite_entry_with_quality_retry(
    target_url: str,
    list_type: str,
    filters: dict,
    progress=None,
    page_count: int = 2,
    progress_label: str = "",
    retry_limit: int = SELLERSPRITE_CATEGORY_RETRY_LIMIT,
) -> tuple[list[Product], list, bool, str]:
    best_products, best_results = collect_sellersprite_entry(
        target_url,
        list_type,
        filters,
        progress=progress,
        page_count=page_count,
        progress_label=progress_label,
    )
    best_ok, best_message = sellersprite_collection_quality(best_products, best_results)
    label = progress_label or target_url
    retry_index = 0
    while not best_ok and retry_index < retry_limit:
        retry_index += 1
        if progress:
            progress(99, f"{best_message}，自动补采 {retry_index}/{retry_limit}")
        log(f"{label}: {best_message}. Retrying collection {retry_index}/{retry_limit}.")
        retry_products, retry_results = collect_sellersprite_entry(
            target_url,
            list_type,
            filters,
            progress=progress,
            page_count=page_count,
            progress_label=f"{label} 重试{retry_index}",
        )
        retry_ok, retry_message = sellersprite_collection_quality(retry_products, retry_results)
        if len(retry_products) > len(best_products):
            best_products, best_results = retry_products, retry_results
            best_ok, best_message = retry_ok, retry_message
        elif retry_ok and not best_ok:
            best_products, best_results = retry_products, retry_results
            best_ok, best_message = retry_ok, retry_message
        else:
            best_ok, best_message = sellersprite_collection_quality(best_products, best_results)
        log(f"{label}: retry {retry_index}/{retry_limit} result: {retry_message}. best: {best_message}.")
    if progress and not best_ok:
        progress(99, f"{best_message}，已保留当前可解析产品")
    return best_products, best_results, best_ok, best_message


def collect_sellersprite_batch(
    seed_url: str,
    list_type: str,
    filters: dict,
    progress_bar,
    progress_start: int = 0,
    progress_end: int = 100,
    progress_prefix: str = "",
) -> list[Product]:
    progress_span = max(1, progress_end - progress_start)

    def set_batch_progress(local_percent: int, text: str):
        global_percent = min(100, progress_start + int((max(0, min(100, local_percent)) / 100) * progress_span))
        progress_bar.progress(global_percent, text=text)

    if not is_rank_category_url(seed_url):
        raise ValueError("当前链接不是具体榜单类目页。为避免自动跳类目，本次不会打开父类页继续发现链接。")

    set_batch_progress(0, "正在采集当前已映射类目，不再自动打开其它类目链接...")

    def update_entry_progress(percent: int, message: str):
        set_batch_progress(percent, message)

    products, refresh_results, quality_ok, quality_message = collect_sellersprite_entry_with_quality_retry(
        seed_url,
        list_type,
        filters,
        progress=update_entry_progress,
        page_count=2,
        progress_label="当前类目",
    )
    if not quality_ok:
        set_batch_progress(99, f"{quality_message}，疑似漏采；已保留当前可解析产品。")
    set_batch_progress(100, f"当前类目采集完成：原始采集 {len(products)} 条，读取 {len(refresh_results)} 页。")
    return products


def resolve_category_seed_urls(selected_paths: list[str], custom_url: str = "") -> list[tuple[str, str]]:
    seeds: list[tuple[str, str]] = []
    seen: set[str] = set()
    for path in collection_seed_paths(selected_paths):
        url = find_exact_category_url(path)
        seed_label = path
        if not url or url in seen:
            continue
        label = path if seed_label == path else f"{path}（通过 {seed_label} 自动发现）"
        seeds.append((label, url))
        seen.add(url)
    if not seeds and custom_url:
        seeds.append(("自定义链接", custom_url))
    return seeds


def collection_seed_paths(selected_paths: list[str]) -> list[str]:
    selected_set = set(selected_paths)
    leaf_paths: list[str] = []
    seen: set[str] = set()
    for path in selected_paths:
        if any(other != path and other.startswith(f"{path} > ") for other in selected_set):
            continue
        node = get_category_node(path)
        expanded_paths = list(iter_category_leaf_paths(path, node)) if node else [path]
        for leaf_path in expanded_paths:
            if leaf_path not in seen:
                leaf_paths.append(leaf_path)
                seen.add(leaf_path)
    return leaf_paths or selected_paths


def selection_contains_parent_category(selected_paths: list[str]) -> bool:
    selected_set = set(selected_paths)
    has_selected_descendants = any(
        any(other != path and other.startswith(f"{path} > ") for other in selected_set)
        for path in selected_paths
    )
    likely_parent_depth = any(path.count(" > ") <= 1 for path in selected_paths)
    return has_selected_descendants or likely_parent_depth


def resolve_primary_collection_url(selected_paths: list[str], custom_url: str = "") -> tuple[str, str]:
    if custom_url:
        return "自定义链接", custom_url
    seeds = resolve_category_seed_urls(selected_paths, "")
    if seeds:
        return seeds[0]
    return "默认测试链接", DEFAULT_TEST_URL


def collect_sellersprite_batch_from_seeds(
    seed_urls: list[tuple[str, str]],
    list_type: str,
    filters: dict,
    progress_bar,
) -> list[Product]:
    if not seed_urls:
        raise ValueError("没有可采集的类目链接。请先选择已映射的类目，或填写自定义 Amazon Best Sellers 链接。")
    raw_by_asin: dict[str, Product] = {}
    for seed_index, (seed_label, seed_url) in enumerate(seed_urls, start=1):
        if stop_collection_requested():
            log(f"Batch seed collection stopped before seed {seed_index}/{len(seed_urls)}.")
            break
        seed_start = int(((seed_index - 1) / len(seed_urls)) * 100)
        seed_end = int((seed_index / len(seed_urls)) * 100)
        progress_bar.progress(
            seed_start,
            text=f"入口 {seed_index}/{len(seed_urls)}｜总原始采集 {len(raw_by_asin)} 条",
        )
        seed_products = collect_sellersprite_batch(
            amazon_url_for_list_type(seed_url, list_type),
            list_type,
            filters,
            progress_bar,
            progress_start=seed_start,
            progress_end=seed_end,
            progress_prefix="",
        )
        added = 0
        for product in seed_products:
            if product.asin in raw_by_asin:
                continue
            raw_by_asin[product.asin] = product
            added += 1
        progress_bar.progress(
            seed_end,
            text=f"入口 {seed_index}/{len(seed_urls)} 完成｜新增 {added} 条｜总原始采集 {len(raw_by_asin)} 条",
        )
        log(f"Seed {seed_index}/{len(seed_urls)} finished: {seed_label}. added raw {added}, total raw {len(raw_by_asin)}.")
    progress_bar.progress(100, text=f"全部入口采集完成：总原始采集 {len(raw_by_asin)} 条。")
    return list(raw_by_asin.values())


def render_cards(products):
    st.markdown("<div class='seller-list-frame'>", unsafe_allow_html=True)
    header_left, header_body = st.columns([0.035, 0.965], gap=None, vertical_alignment="top")
    with header_left:
        st.markdown("<div class='seller-select-header'></div>", unsafe_allow_html=True)
    with header_body:
        st.markdown(
            """
            <div class="seller-header">
                <div>#</div><div>产品信息 <span>数据解释</span></div><div>大类BSR</div><div>销量趋势(父)</div>
                <div>销量(父)<br>增长率</div><div>销售额</div><div>子体销量<br>子体销售额</div><div>变体数</div>
                <div>价格<br>Q&A</div><div>评分数<br>月新增</div><div>评分<br>留评率</div><div>FBA<br>毛利率</div>
                <div>上架时间</div><div>配送<br>买家运费</div><div>操作</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    for product in products:
        row_select, row_body = st.columns([0.035, 0.965], gap=None, vertical_alignment="top")
        with row_select:
            st.markdown("<div class='seller-select-cell'>", unsafe_allow_html=True)
            product.selected = st.checkbox(
                "Include in export",
                value=product.selected,
                key=f"row_include_{product.asin}",
                label_visibility="collapsed",
            )
            st.markdown("</div>", unsafe_allow_html=True)
        with row_body:
            st.markdown(seller_product_html(product), unsafe_allow_html=True)
        st.markdown("<div class='seller-row-space'></div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

ensure_state()

header_left, header_right = st.columns([4.5, 1.15], vertical_alignment="center")
with header_right:
    UI_LANG = st.radio("Language / 语言", ["中文", "English"], horizontal=True, label_visibility="collapsed")
T = TEXT[UI_LANG]
with header_left:
    st.title(T["title"])
    st.caption(T["caption"])

st.markdown(
    """
    <style>
    html, body, [data-testid="stAppViewContainer"] {
        background: #f4f5f7;
        overflow-x: auto;
    }
    [data-testid="stHeader"] {
        background: transparent;
    }
    .block-container {
        padding-top: 1rem;
        padding-left: 1.25rem;
        padding-right: 1.25rem;
        max-width: none;
        background: transparent;
    }
    h1 {
        color: #222733;
        font-size: clamp(30px, 3.4vw, 46px) !important;
        letter-spacing: 0 !important;
        line-height: 1.08 !important;
        margin-bottom: 8px !important;
    }
    div[data-testid="stCaptionContainer"] {
        color: #8b94a3;
        font-size: 14px;
    }
    html:has(div[data-testid="stDialog"]),
    body:has(div[data-testid="stDialog"]) {
        overflow: hidden !important;
    }
    body:has(div[data-testid="stDialog"]) [data-testid="stAppViewContainer"] {
        height: 100vh;
        overflow: hidden !important;
    }
    div[data-testid="stAppViewContainer"]:has(div[data-testid="stDialog"])::before {
        content: "";
        position: fixed;
        inset: 0;
        background: rgba(17, 24, 39, 0.58);
        z-index: 999;
        pointer-events: none;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background: #ffffff;
        border: 1px solid #e7ebf1;
        border-radius: 6px;
        box-shadow: 0 1px 2px rgba(15, 23, 42, .03);
        padding: 8px 10px 10px;
    }
    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #eef0f4;
        border-radius: 6px;
        box-shadow: 0 1px 2px rgba(15, 23, 42, .025);
        padding: 10px 12px;
    }
    div[data-testid="stMetric"] label {
        color: #7b8491 !important;
        font-size: 13px !important;
    }
    div[data-testid="stMetric"] [data-testid="stMetricValue"] {
        color: #2b303b;
        font-size: 28px;
    }
    div[data-testid="stTabs"] {
        background: transparent;
    }
    div[data-testid="stTabs"] [role="tablist"] {
        background: transparent;
        border-bottom: 1px solid #e2e6ee;
    }
    div[data-testid="stTabs"] [role="tabpanel"] {
        padding-top: 12px;
    }
    div[data-testid="stPopover"] button {
        border: 1px solid #d9dee8;
        border-radius: 6px;
        min-height: 44px;
        justify-content: flex-start;
    }
    div[data-testid="stTextInput"] input,
    div[data-testid="stNumberInput"] input,
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
    div[data-testid="stButton"] button {
        min-height: 42px;
    }
    div[data-testid="stButton"] button {
        align-items: center;
        display: inline-flex;
        justify-content: center;
    }
    div[data-testid="stElementContainer"]:has(.collection-action-toolbar) + div[data-testid="stHorizontalBlock"] div[data-testid="stButton"] button {
        min-height: 40px;
        height: 40px;
        border-radius: 8px;
        font-size: 14px;
        padding: 0 14px;
        white-space: nowrap;
    }
    div[data-testid="stElementContainer"]:has(.collection-action-toolbar) + div[data-testid="stHorizontalBlock"] div[data-testid="stButton"] button p {
        white-space: nowrap;
    }
    div[data-testid="stElementContainer"]:has(.collection-action-toolbar) + div[data-testid="stHorizontalBlock"] > div:nth-child(2) button {
        background: #ffffff !important;
        border: 1px solid #ff4b4b !important;
        color: #ff4b4b !important;
        font-weight: 700 !important;
    }
    div[data-testid="stElementContainer"]:has(.collection-action-toolbar) + div[data-testid="stHorizontalBlock"] > div:nth-child(2) button:hover {
        background: #fff1f1 !important;
        border-color: #ff4b4b !important;
        color: #ff4b4b !important;
    }
    div[data-testid="stElementContainer"]:has(.collection-action-toolbar) + div[data-testid="stHorizontalBlock"] {
        flex-wrap: nowrap !important;
        gap: 8px !important;
        overflow-x: auto;
        padding-bottom: 2px;
    }
    div[data-testid="stElementContainer"]:has(.collection-action-toolbar) + div[data-testid="stHorizontalBlock"] > div {
        flex: 0 0 108px !important;
        min-width: 108px !important;
    }
    div[data-testid="stElementContainer"]:has(.collection-action-toolbar) + div[data-testid="stHorizontalBlock"] > div:last-child {
        flex-basis: 132px !important;
        min-width: 132px !important;
    }
    .filter-label {
        color: #5f6673;
        font-size: 15px;
        font-weight: 700;
        margin: 6px 0 8px;
    }
    .filter-label span {
        align-items: center;
        border: 1px solid #ff7a1a;
        border-radius: 50%;
        color: #ff7a1a;
        display: inline-flex;
        font-size: 12px;
        font-weight: 700;
        height: 18px;
        justify-content: center;
        margin-left: 8px;
        width: 18px;
    }
    .filter-range-sep {
        color: #b8bfca;
        font-size: 18px;
        font-weight: 700;
        line-height: 42px;
        text-align: center;
    }
    div[data-testid="stElementContainer"]:has(.range-filter-anchor) + div[data-testid="stHorizontalBlock"] {
        flex-wrap: nowrap !important;
        gap: 8px !important;
        overflow-x: auto;
        padding-bottom: 2px;
    }
    div[data-testid="stColumn"] div[data-testid="stVerticalBlock"]:has(.range-filter-anchor) div[data-testid="stHorizontalBlock"] {
        flex-wrap: nowrap !important;
        gap: 8px !important;
        overflow-x: auto;
        padding-bottom: 2px;
    }
    div[data-testid="stElementContainer"]:has(.range-filter-anchor) + div[data-testid="stHorizontalBlock"] > div:nth-child(1),
    div[data-testid="stElementContainer"]:has(.range-filter-anchor) + div[data-testid="stHorizontalBlock"] > div:nth-child(3),
    div[data-testid="stColumn"] div[data-testid="stVerticalBlock"]:has(.range-filter-anchor) div[data-testid="stHorizontalBlock"] > div:nth-child(1),
    div[data-testid="stColumn"] div[data-testid="stVerticalBlock"]:has(.range-filter-anchor) div[data-testid="stHorizontalBlock"] > div:nth-child(3) {
        flex: 1 0 132px !important;
        min-width: 132px !important;
    }
    div[data-testid="stElementContainer"]:has(.range-filter-anchor) + div[data-testid="stHorizontalBlock"] > div:nth-child(2) {
        flex: 0 0 24px !important;
        min-width: 24px !important;
    }
    div[data-testid="stColumn"] div[data-testid="stVerticalBlock"]:has(.range-filter-anchor) div[data-testid="stHorizontalBlock"] > div:nth-child(2) {
        flex: 0 0 24px !important;
        min-width: 24px !important;
    }
    .filter-money-hint {
        display: none;
    }
    .source-static {
        align-items: center;
        background: #f1f4f8;
        border: 1px solid #e7ebf1;
        border-radius: 6px;
        color: #1f2937;
        display: flex;
        font-size: 15px;
        font-weight: 600;
        min-height: 42px;
        padding: 0 14px;
    }
    div[data-testid="stRadio"] > label,
    div[data-testid="stSelectbox"] > label,
    div[data-testid="stTextInput"] > label,
    div[data-testid="stNumberInput"] > label {
        min-height: 24px;
    }
    div[data-testid="stDialog"] {
        max-height: 100vh;
        overflow: hidden !important;
        z-index: 1000;
    }
    div[data-testid="stDialog"] [role="dialog"] {
        border-radius: 4px !important;
        height: min(72vh, 680px) !important;
        max-height: calc(100vh - 56px) !important;
        display: flex !important;
        flex-direction: column !important;
        overflow: hidden !important;
        width: min(78vw, 1240px) !important;
    }
    div[data-testid="stDialog"] section {
        height: 100% !important;
        padding-bottom: 0 !important;
        max-height: calc(100vh - 72px) !important;
        overflow: hidden !important;
    }
    div[data-testid="stDialog"] section > div {
        height: 100% !important;
        overflow: hidden !important;
    }
    div[data-testid="stDialog"] div[data-testid="stVerticalBlock"] {
        gap: 0.55rem;
    }
    div[data-testid="stDialog"] input {
        background: #ffffff !important;
        box-shadow: none !important;
        font-size: 16px !important;
    }
    div[data-testid="stDialog"] div[data-baseweb="input"] {
        border-color: #d8dee8 !important;
        box-shadow: none !important;
    }
    div[data-testid="stDialog"] div[data-baseweb="input"]:focus-within {
        border-color: #d8dee8 !important;
        box-shadow: none !important;
        outline: none !important;
    }
    div[data-testid="stDialog"] input:focus {
        box-shadow: none !important;
        outline: none !important;
    }
    div[data-testid="stDialog"] div[data-testid="stCheckbox"] {
        min-height: 32px;
    }
    div[data-testid="stDialog"] div[data-testid="stCheckbox"] label {
        align-items: center;
    }
    div[data-testid="stDialog"] div[data-testid="stHorizontalBlock"] > div:first-child button[kind="secondary"] {
        background: transparent !important;
        border: 0 !important;
        box-shadow: none !important;
        color: #b8c0cc !important;
        font-size: 17px !important;
        height: 28px !important;
        min-height: 28px !important;
        padding: 0 !important;
    }
    div[data-testid="stDialog"] div[data-testid="stHorizontalBlock"] > div:first-child button[kind="secondary"]:hover {
        background: transparent !important;
        color: #ff7a1a !important;
    }
    .category-tree-label {
        color: #4b5563;
        font-size: 15px;
        line-height: 32px;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }
    .category-tree-indent {
        display: block;
        height: 32px;
        position: relative;
    }
    .category-tree-indent::after {
        background: #edf0f5;
        content: "";
        height: 32px;
        position: absolute;
        right: 8px;
        top: 0;
        width: 1px;
    }
    .category-tree-root-indent {
        display: block;
        height: 32px;
        width: 1px;
    }
    .category-count-badge {
        background: #fff4e5;
        border-radius: 5px;
        color: #ff991f;
        display: inline-block;
        float: right;
        font-size: 14px;
        line-height: 26px;
        min-width: 54px;
        padding: 0 7px;
        text-align: center;
    }
    .category-tree-spacer {
        display: inline-block;
        height: 32px;
        width: 100%;
    }
    .category-dialog-body {
        max-height: min(62vh, 620px);
        overflow-y: auto;
        padding: 2px 4px 10px 0;
        scrollbar-gutter: stable;
    }
    .category-selected-title {
        color: #2f3642;
        font-size: 20px;
        font-weight: 700;
        margin: 0;
        line-height: 48px;
    }
    .category-selected-panel {
        background: #ffffff;
        border: 1px solid #e7ebf1;
        border-radius: 6px;
        max-height: min(44vh, 420px);
        min-height: 92px;
        overflow-y: auto;
        padding: 12px;
    }
    .category-selected-empty {
        color: #8b94a3;
        font-size: 15px;
        line-height: 1.5;
        min-height: 92px;
        padding: 8px 0 12px;
    }
    .category-dialog-footer {
        background: #ffffff;
        border-top: 1px solid #e7ebf1;
        bottom: 0;
        flex: 0 0 auto;
        margin: 8px -4px 0;
        padding: 12px 4px 0;
        position: sticky;
        z-index: 10;
    }
    div[data-testid="stDialog"] div[data-testid="stElementContainer"]:has(.category-footer-anchor) {
        margin-top: auto !important;
        padding-top: 10px !important;
        border-top: 1px solid #e7ebf1 !important;
    }
    div[data-testid="stDialog"] div[data-testid="stElementContainer"]:has(.category-footer-anchor) + div[data-testid="stHorizontalBlock"] {
        background: #ffffff !important;
        padding: 0 0 10px !important;
        position: sticky !important;
        bottom: 0 !important;
        z-index: 20 !important;
    }
    div[data-testid="stDialog"] div[data-testid="stElementContainer"]:has(.category-footer-anchor) + div[data-testid="stHorizontalBlock"] button {
        min-height: 40px !important;
        height: 40px !important;
        border-radius: 4px !important;
        font-size: 14px !important;
    }
    .cache-card {
        background: #ffffff;
        border: 1px solid #e7ebf1;
        border-radius: 6px;
        box-shadow: 0 1px 2px rgba(15, 23, 42, .03);
        margin: 12px 0 4px;
        padding: 12px 14px;
    }
    .cache-card-top {
        align-items: center;
        display: flex;
        gap: 12px;
        justify-content: space-between;
    }
    .cache-title {
        color: #1f2937;
        font-size: 14px;
        font-weight: 700;
        line-height: 1.35;
    }
    .cache-sub {
        color: #8a94a3;
        font-size: 12px;
        line-height: 1.45;
        margin-top: 2px;
    }
    .cache-sub a {
        color: #ff7a1a;
        text-decoration: none;
    }
    .cache-badges {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        justify-content: flex-end;
    }
    .cache-badge {
        background: #f3f5f8;
        border-radius: 999px;
        color: #657080;
        display: inline-flex;
        font-size: 12px;
        padding: 3px 9px;
        white-space: nowrap;
    }
    .cache-badge.ok {
        background: #ecfdf5;
        color: #059669;
    }
    .cache-badge.warn {
        background: #fff7ed;
        color: #ea580c;
    }
    .cache-progress {
        background: #eef1f5;
        border-radius: 999px;
        height: 6px;
        margin-top: 10px;
        overflow: hidden;
    }
    .cache-progress span {
        background: #ff4d4f;
        border-radius: inherit;
        display: block;
        height: 100%;
    }
    .cache-foot {
        color: #8a94a3;
        font-size: 12px;
        margin-top: 6px;
    }
    .cache-foot strong {
        color: #ff7a1a;
    }
    .cache-source-note {
        color: #6b7280;
        font-size: 12px;
        line-height: 1.5;
        margin-top: 6px;
    }
    .selected-pill {
        display: inline-block;
        background: #fff4e5;
        color: #f28c18;
        border-radius: 6px;
        padding: 4px 8px;
        margin: 2px 4px 2px 0;
        font-size: 12px;
        max-width: 100%;
        overflow: hidden;
        text-overflow: ellipsis;
        vertical-align: middle;
        white-space: nowrap;
    }
    .product-row-gap {
        height: 14px;
    }
    .cards-scroll-title {
        color: #5f6b7a;
        font-size: 13px;
        margin: 8px 0 4px;
    }
    .toolbar-meta {
        color: #7a8491;
        font-size: 14px;
        line-height: 40px;
        white-space: nowrap;
    }
    .toolbar-meta strong {
        color: #ff7a1a;
    }
    .toolbar-spacer {
        height: 12px;
    }
    .seller-list-frame {
        background: #f1f3f6;
        border-radius: 4px;
        overflow-x: auto;
        padding: 0 0 6px;
        position: relative;
        scrollbar-gutter: auto;
    }
    .seller-list-frame::-webkit-scrollbar {
        height: 10px;
    }
    .seller-list-frame::-webkit-scrollbar-thumb {
        background: #d9dee7;
        border-radius: 999px;
    }
    .seller-row-space {
        height: 10px;
    }
    .seller-select-header {
        background: transparent;
        border-top: 0;
        border-bottom: 0;
        border-left: 0;
        border-right: 0;
        box-shadow: none;
        min-height: 58px;
        position: sticky;
        top: 0;
        z-index: 20;
    }
    .seller-select-cell {
        align-items: flex-start;
        background: transparent;
        border: 0;
        border-radius: 0;
        display: flex;
        min-height: 178px;
        padding: 18px 0 0 0;
        justify-content: center;
    }
    .seller-header {
        display: grid;
        grid-template-columns: 30px 254px 56px 92px 68px 78px 82px 46px 62px 64px 52px 58px 70px 56px 32px;
        align-items: center;
        gap: 7px;
        min-width: 1190px;
        width: max(100%, 1190px);
        background: #f6f7f9;
        border: 1px solid #eef0f4;
        border-left: 0;
        border-right: 0;
        color: #6f7782;
        font-size: 12px;
        font-weight: 700;
        min-height: 58px;
        padding: 8px 10px;
        position: sticky;
        top: 0;
        z-index: 20;
        box-shadow: none;
    }
    .seller-header span {
        color: #ff7a1a;
        font-weight: 500;
        margin-left: 8px;
    }
    .seller-header > div:nth-child(2) {
        padding-left: 96px;
    }
    .seller-row {
        min-width: 1190px;
        width: max(100%, 1190px);
        background: #ffffff;
        border: 1px solid #e8edf3;
        border-radius: 3px;
        box-shadow: 0 1px 1px rgba(15, 23, 42, .02);
        padding: 12px 10px 9px;
        transition: border-color .15s ease, box-shadow .15s ease, background .15s ease;
    }
    .seller-row:hover {
        border-color: #e2e7ef;
        box-shadow: 0 4px 12px rgba(15, 23, 42, .04);
    }
    .seller-main {
        display: grid;
        grid-template-columns: 30px 254px 56px 92px 68px 78px 82px 46px 62px 64px 52px 58px 70px 56px 32px;
        align-items: center;
        gap: 7px;
    }
    .seller-rank {
        color: #a6aeb9;
        font-size: 16px;
        text-align: center;
    }
    .seller-product {
        display: grid;
        grid-template-columns: 88px minmax(0, 1fr);
        gap: 12px;
        align-items: center;
    }
    .seller-image-wrap {
        position: relative;
        width: 88px;
        min-height: 106px;
    }
    .seller-image-wrap img {
        width: 84px;
        height: 84px;
        object-fit: contain;
        margin-top: 22px;
    }
    .seller-info {
        min-width: 0;
    }
    .seller-title {
        color: #111827;
        font-size: 14px;
        font-weight: 500;
        line-height: 1.3;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
        margin-bottom: 6px;
    }
    .cell {
        color: #151a24;
        font-size: 12px;
        line-height: 1.55;
        text-align: center;
    }
    .cell strong {
        display: block;
        font-size: 14px;
        font-weight: 500;
    }
    .muted {
        color: #9aa2ae;
    }
    .green {
        color: #43af52;
        display: block;
    }
    .trend-cell svg {
        width: 100px;
        height: 52px;
        display: block;
    }
    .ops-cell {
        color: #ff7a1a;
        display: grid;
        gap: 7px;
        justify-items: center;
        font-size: 15px;
    }
    .seller-detail {
        border-top: 1px solid #eef0f4;
        color: #7a828e;
        font-size: 13px;
        line-height: 1.75;
        margin: 9px 0 0 30px;
        padding: 8px 0 0 0;
    }
    .orange {
        color: #ff7a1a;
    }
    .pill {
        border-radius: 4px;
        color: #ffffff;
        display: inline-block;
        font-size: 13px;
        font-weight: 600;
        line-height: 1;
        margin-left: 6px;
        padding: 4px 8px;
    }
    .orange-pill { background: #ff8617; }
    .blue-pill { background: #4b6ff3; }
    .green-pill { background: #35c28d; }
    .rank-pill {
        background: #ff8617;
        border-radius: 10px;
        color: #ffffff;
        display: inline-block;
        font-size: 12px;
        margin: 0 6px;
        padding: 1px 7px;
    }
    .product-identity {
        display: grid;
        grid-template-columns: 42px 150px minmax(0, 1fr);
        gap: 22px;
        align-items: start;
        min-height: 170px;
        padding-top: 8px;
    }
    div[data-testid="stCheckbox"]:has(input[id*="image_include_"]) {
        margin-bottom: -34px;
        position: relative;
        z-index: 3;
        width: 28px;
    }
    .product-rank {
        color: #a3abb7;
        font-size: 22px;
        text-align: center;
        padding-top: 82px;
    }
    .product-media {
        position: relative;
        width: 150px;
        height: 165px;
        display: flex;
        flex-direction: column;
        align-items: flex-start;
        justify-content: flex-start;
    }
    .product-media img {
        width: 128px;
        height: 128px;
        object-fit: contain;
        margin-left: 4px;
        margin-top: 22px;
    }
    .level-corner {
        position: absolute;
        top: 0;
        left: 0;
        background: #ef2b13;
        color: #ffffff;
        border-radius: 5px;
        padding: 2px 6px;
        font-weight: 700;
        font-size: 15px;
        line-height: 1.1;
    }
    .signal-tags {
        display: flex;
        gap: 3px;
        margin-top: 7px;
    }
    .tag {
        color: #ffffff;
        border-radius: 4px;
        padding: 2px 5px;
        font-weight: 700;
        font-size: 13px;
        line-height: 1.2;
    }
    .tag-bs { background: #ff8617; }
    .tag-ac { background: #101827; }
    .tag-nr { background: #f02816; }
    .product-info {
        min-width: 0;
        padding-top: 44px;
    }
    .product-title {
        color: #111827;
        font-size: 21px;
        font-weight: 500;
        line-height: 1.3;
        max-width: 760px;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
        margin-bottom: 6px;
    }
    .meta-line {
        color: #7b8491;
        font-size: 15px;
        line-height: 1.5;
        white-space: nowrap;
    }
    .meta-line strong {
        color: #111827;
        font-weight: 500;
    }
    .meta-line.muted {
        color: #a6adb8;
    }
    .mini-icon {
        color: #c7ced8;
        display: inline-block;
        font-size: 14px;
        margin-left: 5px;
        vertical-align: 1px;
    }
    .copy-icon {
        background: transparent;
        border: 1px solid transparent;
        border-radius: 4px;
        color: #aeb7c3;
        cursor: pointer;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 16px;
        height: 22px;
        margin-left: 7px;
        min-width: 22px;
        padding: 0 4px;
        position: relative;
        transition: all .15s ease;
        vertical-align: 0;
    }
    .copy-icon:hover,
    .copy-icon.copied {
        background: #fff4e8;
        border-color: #ffb26f;
        color: #ff7a1a;
    }
    .copy-icon:active {
        background: #ff7a1a;
        border-color: #ff7a1a;
        color: #ffffff;
    }
    .copy-icon.disabled-icon {
        cursor: default;
        opacity: .45;
    }
    .copy-icon.disabled-icon:hover {
        background: transparent;
        border-color: transparent;
        color: #aeb7c3;
    }
    .mini-link {
        border: 1px solid transparent;
        border-radius: 4px;
        color: #aeb7c3 !important;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 17px;
        height: 22px;
        margin-left: 5px;
        min-width: 22px;
        padding: 0 4px;
        position: relative;
        text-decoration: none !important;
        transition: all .15s ease;
        vertical-align: 0;
    }
    .mini-link:hover {
        background: #fff4e8;
        border-color: #ffb26f;
        color: #ff7a1a !important;
    }
    div[data-testid="stExpander"] label p,
    div[data-testid="stCheckbox"] label p {
        font-family: inherit;
        white-space: normal;
        line-height: 1.35;
    }
    @media (max-width: 760px) {
        .block-container {
            padding-left: 0.85rem;
            padding-right: 0.85rem;
        }
        h1 {
            font-size: 32px !important;
            max-width: 12em;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] {
            padding: 6px;
        }
        .filter-label {
            font-size: 14px;
        }
        .source-static,
        div[data-testid="stTextInput"] input,
        div[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
        div[data-testid="stButton"] button {
            min-height: 40px;
        }
        div[data-testid="stColumn"] div[data-testid="stVerticalBlock"]:has(.range-filter-anchor) div[data-testid="stHorizontalBlock"] {
            gap: 6px !important;
            overflow-x: visible;
        }
        div[data-testid="stColumn"] div[data-testid="stVerticalBlock"]:has(.range-filter-anchor) div[data-testid="stHorizontalBlock"] > div:nth-child(1),
        div[data-testid="stColumn"] div[data-testid="stVerticalBlock"]:has(.range-filter-anchor) div[data-testid="stHorizontalBlock"] > div:nth-child(3) {
            flex: 1 1 76px !important;
            min-width: 76px !important;
        }
        div[data-testid="stColumn"] div[data-testid="stVerticalBlock"]:has(.range-filter-anchor) div[data-testid="stHorizontalBlock"] > div:nth-child(2) {
            flex: 0 0 18px !important;
            min-width: 18px !important;
        }
        div[data-testid="stColumn"] div[data-testid="stVerticalBlock"]:has(.range-filter-anchor) input {
            font-size: 13px;
            padding-left: 10px;
            padding-right: 10px;
        }
        .toolbar-meta {
            line-height: 32px;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if st.session_state.show_category_dialog:
    render_category_dialog()

with st.container(border=True):
    top_cols = st.columns([1.1, 1.25, 2.45, 2.0], vertical_alignment="bottom")
    with top_cols[0]:
        list_type = st.radio(T["list_type"], ["New Releases", "Best Sellers"], horizontal=True)
    with top_cols[1]:
        data_source = "卖家精灵插件"
        st.write("**数据源**")
        st.markdown("<div class='source-static'>卖家精灵插件</div>", unsafe_allow_html=True)
    with top_cols[2]:
        st.write(f"**{T['categories']}**")
        if st.button(
            f"选择类目（已选 {len(st.session_state.confirmed_category_paths)}）",
            use_container_width=True,
        ):
            st.session_state.show_category_dialog = True
            st.rerun()
    with top_cols[3]:
        custom_url = st.text_input(T["custom_url"], placeholder="https://www.amazon.com/...")
        batch_category_collect = st.checkbox("大类批量采集", value=False, help="打开当前大类页，自动发现小类链接，并逐个采集合格产品。")

    selected_paths = st.session_state.confirmed_category_paths
    if selected_paths:
        preview = selected_paths[:6]
        pills = "".join(f"<span class='selected-pill'>{escape(path)}</span>" for path in preview)
        more = f" +{len(selected_paths) - len(preview)}" if len(selected_paths) > len(preview) else ""
        st.markdown(pills + more, unsafe_allow_html=True)
    else:
        st.caption("暂未选择类目。")
    mapped_seed_urls = resolve_category_seed_urls(selected_paths, "")
    unmapped_paths = [path for path in selected_paths if not find_category_seed_url(path)[1]]
    if data_source == "卖家精灵插件" and batch_category_collect and selected_paths:
        if mapped_seed_urls:
            st.caption("批量采集入口：" + "；".join(label for label, _ in mapped_seed_urls))
        if unmapped_paths:
            st.warning("这些类目还没有绑定具体 Amazon 榜单链接。为避免自动乱跳链接，本次不会自动打开父类页去发现链接：" + "；".join(unmapped_paths[:6]))

    seller_cache_total = 0
    seller_cache_hydrated = 0
    chrome_ready = False
    if data_source == "卖家精灵插件":
        seller_cache_total, seller_cache_hydrated = sellersprite_cache_hydration()
        chrome_ready = chrome_debugger_available()
        cache_warning = sellersprite_cache_warning()
        if cache_warning:
            st.warning(cache_warning)
        if not chrome_ready:
            st.warning("实时采集需要连接采集 Chrome。未连接时不会使用旧产品缓存替代，请先双击“一键启动工具.bat”。")

    if st.session_state.last_cache_refresh_message:
        st.info(st.session_state.last_cache_refresh_message)
    if st.session_state.last_category_mapping_message:
        st.info(st.session_state.last_category_mapping_message)

    st.divider()
    history_options = raw_history_options()
    if history_options:
        history_cols = st.columns([3, 1], vertical_alignment="bottom")
        history_labels = [label for label, _ in history_options]
        selected_history_label = history_cols[0].selectbox("最近采集记录", history_labels, label_visibility="collapsed")
        load_history = history_cols[1].button("载入该记录", use_container_width=True)
    else:
        selected_history_label = ""
        load_history = False

    st.write(f"**{T['filters']}**")
    filter_top = st.columns(3)
    with filter_top[0]:
        min_price_raw, max_price_raw = render_range_filter(T["price"], "filter_price", "24.99", "200.00", money=True)
    with filter_top[1]:
        min_reviews_raw, max_reviews_raw = render_range_filter(T["reviews"], "filter_reviews", "", "300")
    with filter_top[2]:
        min_bought_raw, max_bought_raw = render_range_filter(T["monthly_sales"], "filter_monthly_sales", "100", "")

    filter_bottom = st.columns([1, 1, 1], vertical_alignment="bottom")
    with filter_bottom[0]:
        min_child_sales_raw, max_child_sales_raw = render_range_filter(T["child_sales"], "filter_child_sales")
    with filter_bottom[1]:
        min_bsr_raw, max_bsr_raw = render_range_filter(T["bsr"], "filter_bsr")
    with filter_bottom[2]:
        st.markdown(f"<div class='filter-label'>{escape(T['launched_at'])} <span>?</span></div>", unsafe_allow_html=True)
        launch_options = ["不限", "近30天", "近60天", "近3个月", "近半年", "近1年", "近2年", "近1~2年"] if UI_LANG == "中文" else ["Any", "Last 30 days", "Last 60 days", "Last 3 months", "Last 6 months", "Last year", "Last 2 years", "1-2 years"]
        launch_window = st.selectbox("上架时间", launch_options, key="filter_launch_window", label_visibility="collapsed")
    current_filters = {
        "min_price": parse_filter_number(min_price_raw, 0.0),
        "max_price": parse_filter_number(max_price_raw, 999999.0),
        "min_reviews": parse_filter_number(min_reviews_raw, None, as_int=True),
        "max_reviews": parse_filter_number(max_reviews_raw, 999999999, as_int=True),
        "min_bought": parse_filter_number(min_bought_raw, 0, as_int=True),
        "max_bought": parse_filter_number(max_bought_raw, None, as_int=True),
        "min_child_sales": parse_filter_number(min_child_sales_raw, None, as_int=True),
        "max_child_sales": parse_filter_number(max_child_sales_raw, None, as_int=True),
        "min_bsr": parse_filter_number(min_bsr_raw, None, as_int=True),
        "max_bsr": parse_filter_number(max_bsr_raw, None, as_int=True),
        "launch_window": launch_window,
    }

    st.markdown("<div class='collection-action-toolbar'></div>", unsafe_allow_html=True)
    action_cols = st.columns([1.05, 1.05, 1.05, 1.05, 1.2], vertical_alignment="center")
    seller_cache_can_run = data_source != "卖家精灵插件" or chrome_ready
    run = action_cols[0].button(T["run"], type="primary", use_container_width=True, disabled=not seller_cache_can_run)
    action_cols[1].button("停止采集", use_container_width=True, on_click=request_stop_collection)
    apply_filter = action_cols[2].button("应用筛选", use_container_width=True, disabled=not st.session_state.raw_products)
    clear_filters = action_cols[3].button("清空筛选", use_container_width=True, on_click=reset_filter_widgets)
    load_last_raw = action_cols[4].button("载入上次采集", use_container_width=True)

    raw_count = len(st.session_state.raw_products)
    filtered_count = len(st.session_state.products)
    if raw_count:
        st.caption(f"原始采集池：{raw_count} 条｜当前筛选：{filtered_count} 条｜筛掉：{raw_count - filtered_count} 条")
    if st.session_state.last_raw_products_message:
        st.info(st.session_state.last_raw_products_message)

if clear_filters and st.session_state.raw_products:
    apply_filters_to_raw_pool(current_filters)
    log("Filter widgets reset and filters re-applied to raw product pool.")

if load_last_raw:
    loaded_products, message = load_raw_products()
    st.session_state.raw_products = loaded_products
    st.session_state.last_raw_products_message = message
    if loaded_products:
        apply_filters_to_raw_pool(current_filters)
        log(f"Loaded raw product pool from disk: {len(loaded_products)} products.")
    else:
        st.session_state.products = []
        st.session_state.last_collection_summary = ""

if load_history:
    selected_history_path = dict(history_options).get(selected_history_label)
    if selected_history_path:
        loaded_products, payload, error = load_raw_products_payload(selected_history_path)
        if error:
            st.session_state.raw_products = []
            st.session_state.products = []
            st.session_state.last_collection_summary = ""
            st.session_state.last_raw_products_message = error
        else:
            st.session_state.raw_products = loaded_products
            apply_filters_to_raw_pool(current_filters)
            saved_at = payload.get("saved_at", "-")
            label = payload.get("label") or "未命名采集"
            st.session_state.last_raw_products_message = (
                f"已载入历史采集：{label}，{len(loaded_products)} 条，保存时间：{saved_at}。"
            )
            log(f"Loaded historical raw product pool: {len(loaded_products)} products from {selected_history_path}.")

if apply_filter:
    apply_filters_to_raw_pool(current_filters)
    log(f"Re-applied filters to raw product pool: kept {len(st.session_state.products)}/{len(st.session_state.raw_products)}.")

if run:
    clear_stop_collection_flag()
    filters = current_filters
    try:
        collected_products = []
        if data_source == "卖家精灵插件":
            target_label, target_url = resolve_primary_collection_url(selected_paths, custom_url)
            target_url = amazon_url_for_list_type(target_url, list_type)
            collection_label = "；".join(selected_paths[:3]) if selected_paths else target_label
            if len(selected_paths) > 3:
                collection_label += f" +{len(selected_paths) - 3}"
            should_batch_category_collect = batch_category_collect or selection_contains_parent_category(selected_paths)
            if should_batch_category_collect:
                seed_urls = resolve_category_seed_urls(selected_paths, custom_url)
                log("Start batch category collection from category selection.")
                batch_bar = st.progress(0, text="正在准备大类批量采集...")
                collected_products = collect_sellersprite_batch_from_seeds(seed_urls, list_type, filters, batch_bar)
            else:
                if selected_paths and (len(selected_paths) > 1 or any(find_exact_category_url(path) for path in selected_paths)):
                    st.session_state.last_category_mapping_message = (
                        (st.session_state.last_category_mapping_message + " " if st.session_state.last_category_mapping_message else "")
                        + "当前未勾选“大类批量采集”，本次只采集一个榜单入口页。"
                    )
                log(f"Open Amazon page and wait for SellerSprite plugin data: {target_label}.")
                refresh_bar = st.progress(0, text="正在打开 Amazon，并等待卖家精灵插件加载...")

                def update_run_refresh_progress(percent: int, message: str):
                    refresh_bar.progress(percent, text=message)

                collected_products, refresh_results, quality_ok, quality_message = collect_sellersprite_entry_with_quality_retry(
                    target_url,
                    list_type,
                    filters,
                    progress=update_run_refresh_progress,
                    page_count=2,
                    progress_label=target_label,
                )
                total_product_count = sum(result.product_count for result in refresh_results)
                total_hydrated_count = sum(result.hydrated_count for result in refresh_results)
                total_image_count = sum(result.image_count for result in refresh_results)
                log(
                    "SellerSprite plugin data refresh: "
                    f"{len(collected_products)} unique products, "
                    f"{total_product_count} page products, {total_hydrated_count} hydrated, "
                    f"{total_image_count} images."
                )
                st.session_state.last_cache_refresh_message = (
                    f"本入口采集完成：读取 {len(refresh_results)} 页，识别 {total_product_count} 条页面产品，"
                    f"去重后 {len(collected_products)} 条，{total_hydrated_count} 条插件字段完整。{quality_message}"
                )
                if not quality_ok:
                    st.session_state.last_cache_refresh_message += " 这次疑似漏采，建议保持采集 Chrome 前台可见后重试。"
                log(f"SellerSprite plugin collection finished. Parsed {len(collected_products)} unique products from {len(refresh_results)} pages.")
        for product in collected_products:
            product.selected = False
        st.session_state.raw_products = collected_products
        apply_filters_to_raw_pool(filters)
        if collected_products:
            save_raw_products(collected_products, collection_label, target_url)
            st.session_state.last_raw_products_message = (
                f"已保存原始采集池：{len(collected_products)} 条。最近 5 次采集会保留在本地。"
            )
        else:
            st.session_state.last_collection_summary = "本次没有解析到产品。请检查 Amazon 页面是否正常打开、卖家精灵插件是否已加载。"
            st.session_state.last_raw_products_message = ""
        log(
            "Filters applied: "
            f"price {filters['min_price']}-{filters['max_price']}, "
            f"reviews {filters['min_reviews'] or 0}-{filters['max_reviews']}, "
            f"monthly sales {filters['min_bought']}-{filters['max_bought'] or 'any'}, "
            f"child sales {filters['min_child_sales'] or 0}-{filters['max_child_sales'] or 'any'}, "
            f"BSR {filters['min_bsr'] or 0}-{filters['max_bsr'] or 'any'}, "
            f"launch {filters['launch_window']}. "
            f"Kept {len(st.session_state.products)}/{len(collected_products)}, "
            f"removed {len(collected_products) - len(st.session_state.products)}."
        )
    except CollectionStopped as exc:
        if collected_products:
            for product in collected_products:
                product.selected = False
            st.session_state.raw_products = collected_products
            apply_filters_to_raw_pool(filters)
            save_raw_products(collected_products, collection_label if "collection_label" in locals() else "手动停止采集", target_url if "target_url" in locals() else "")
            st.session_state.last_raw_products_message = f"采集已停止，已保存当前原始采集池：{len(collected_products)} 条。"
        st.session_state.last_collection_summary = str(exc)
        log(f"Collection stopped by user. Kept {len(collected_products)} products.")
        st.warning(str(exc))
    except Exception as exc:
        st.session_state.raw_products = []
        st.session_state.products = []
        st.session_state.last_collection_summary = f"采集失败：{exc}"
        log(f"{data_source} collection failed: {exc}.")
        st.warning(f"{data_source} 实时采集失败：{exc}。请检查采集 Chrome、Amazon 登录、卖家精灵插件或类目链接。")

products = st.session_state.products
sync_product_selection_from_widgets(products)

if st.session_state.last_collection_summary:
    if products:
        st.success(st.session_state.last_collection_summary)
    else:
        st.warning(st.session_state.last_collection_summary)

summary_cols = st.columns(6)
summary_cols[0].metric("产品数" if UI_LANG == "中文" else "Products", len(products))
summary_cols[1].metric("已勾选" if UI_LANG == "中文" else "Selected", sum(1 for p in products if p.selected))
summary_cols[2].metric("A 级", sum(1 for p in products if p.potential_level == "A"))
summary_cols[3].metric("B 级", sum(1 for p in products if p.potential_level == "B"))
summary_cols[4].metric("风险" if UI_LANG == "中文" else "Risk", sum(1 for p in products if p.potential_level == "Risk"))
summary_cols[5].metric("平均分" if UI_LANG == "中文" else "Avg Score", round(sum(p.potential_score for p in products) / len(products), 1) if products else 0)

tab_cards, tab_table, tab_log = st.tabs([T["cards"], T["table"], T["log"]])

with tab_cards:
    if not products:
        if st.session_state.raw_products:
            st.info("原始采集池已有产品，但没有符合当前筛选条件的结果。可以放宽筛选条件后点击“应用筛选”，不需要重新采集。")
        elif st.session_state.last_collection_summary:
            st.info("本次没有解析到产品，或当前类目确实没有可读取产品。请看上方采集提示和日志。")
        else:
            st.info("还没有产品数据。选择类目后点击“开始采集”。")
    else:
        selected_products = [p for p in products if p.selected]
        toolbar = st.columns([0.45, 1.05, 1.05, 0.85, 1.0, 1.35, 0.9, 1.15, 0.95, 0.75], vertical_alignment="center")
        all_selected = bool(products) and len(selected_products) == len(products)
        select_summary = (
            f"已全选 <strong>{len(selected_products)}</strong> 条"
            if all_selected
            else f"已勾选 <strong>{len(selected_products)}</strong> / {len(products)} 条"
        )
        st.session_state["select_all_products"] = all_selected
        master_selected = toolbar[0].checkbox(
            "全选",
            key="select_all_products",
            label_visibility="collapsed",
            on_change=handle_select_all_products_change,
        )
        selected_products = [p for p in products if p.selected]
        toolbar[1].markdown(f"<div class='toolbar-meta'>{select_summary}</div>", unsafe_allow_html=True)
        if toolbar[2].button("复制ASIN", use_container_width=True, disabled=not selected_products):
            log(f"Copied {len(selected_products)} ASIN values.")
        toolbar[3].download_button(
            "导出",
            data=excel_bytes(selected_products),
            file_name="amazon_selection_selected.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=not selected_products,
            use_container_width=True,
        )
        toolbar[4].download_button(
            "导出明细",
            data=excel_bytes(products),
            file_name="amazon_selection_all.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        toolbar[5].markdown(f"<div class='toolbar-meta'>搜索结果数：<strong>{len(products):,}</strong></div>", unsafe_allow_html=True)
        toolbar[6].radio("View", ["列表", "大图"], horizontal=True, label_visibility="collapsed")
        toolbar[7].selectbox("排序字段", ["月销量", "评分", "价格", "上架时间"], label_visibility="collapsed")
        toolbar[8].selectbox("排序", ["降序", "升序"], label_visibility="collapsed")
        toolbar[9].button("确定", type="primary", use_container_width=True)
        st.markdown("<div class='toolbar-spacer'></div>", unsafe_allow_html=True)
        render_cards(products)
        render_clipboard_bridge()

with tab_table:
    if products:
        st.dataframe(
            [asdict(p) for p in products],
            use_container_width=True,
            hide_index=True,
        )
        selected_products = [p for p in products if p.selected]
        st.download_button(
            "导出已选 Excel" if UI_LANG == "中文" else "Download selected Excel",
            data=excel_bytes(selected_products),
            file_name="amazon_selection_selected.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=not selected_products,
        )
        st.download_button(
            "导出全部 Excel" if UI_LANG == "中文" else "Download all Excel",
            data=excel_bytes(products),
            file_name="amazon_selection_all.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            "导出已选 CSV" if UI_LANG == "中文" else "Download selected CSV",
            data=csv_bytes(selected_products),
            file_name="amazon_selection_selected.csv",
            mime="text/csv",
            disabled=not selected_products,
        )
        st.download_button(
            "导出全部 CSV" if UI_LANG == "中文" else "Download all CSV",
            data=csv_bytes(products),
            file_name="amazon_selection_all.csv",
            mime="text/csv",
        )
    else:
        st.info("暂无产品数据。" if UI_LANG == "中文" else "No products yet.")

with tab_log:
    st.code("\n".join(st.session_state.run_log[-80:]))





