from __future__ import annotations

import base64
import json
import os
import re
import socket
import struct
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.error import URLError
from urllib.parse import parse_qs, quote, urljoin, urlparse
from urllib.request import Request, urlopen


DEFAULT_CDP_PORT = 9222
EMPTY_NEW_RELEASES_TEXT = "there are no hot new releases available in this category"
EMPTY_NEW_RELEASES_MESSAGE = "Amazon 明确显示该类目暂无热门新品。"


def _check_stop(stop_check=None) -> None:
    if stop_check:
        stop_check()


def _interruptible_sleep(seconds: float, stop_check=None, interval: float = 0.2) -> None:
    deadline = time.monotonic() + max(0.0, seconds)
    while True:
        _check_stop(stop_check)
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            return
        time.sleep(min(interval, remaining))
SELLERSPRITE_MARKERS = ("近30天销量(父体)", "销售额", "FBA费用")


INVALID_RANK_CATEGORY_PREFIX = "榜单入口校验失败"


@dataclass
class ChromeRefreshResult:
    ok: bool
    product_count: int
    hydrated_count: int
    image_count: int
    source_url: str
    message: str
    next_page_url: str = ""


@dataclass
class CategoryLink:
    title: str
    url: str
    node: str = ""
    depth: int = 0
    is_leaf: bool = True
    path: str = ""


def _rank_category_parts(url: str) -> tuple[str, str, str]:
    parsed = urlparse(url or "")
    gp_match = re.search(
        r"/gp/(bestsellers|new-releases)/([^/?#]+)(?:/(\d{5,}))?",
        parsed.path,
    )
    if gp_match:
        kind = gp_match.group(1)
        department = gp_match.group(2).strip("/")
        node = gp_match.group(3) or parse_qs(parsed.query).get("node", [""])[0]
        return kind, department, node

    zgbs_match = re.search(r"/zgbs/([^/?#]+)/(\d{5,})(?:/|$)", parsed.path)
    if zgbs_match:
        return "bestsellers", zgbs_match.group(1).strip("/"), zgbs_match.group(2)

    return "", "", ""


def is_rank_category_url(url: str) -> bool:
    kind, department, node = _rank_category_parts(url)
    return bool(kind and department and node)


def rank_category_identity(url: str) -> tuple[str, str]:
    _, department, node = _rank_category_parts(url)
    return department, node


def validate_rank_category_page(expected_url: str, page_state: dict | None) -> tuple[bool, str]:
    state = page_state if isinstance(page_state, dict) else {}
    actual_url = str(state.get("url") or "")
    selected_text = re.sub(r"\s+", " ", str(state.get("selectedText") or "")).strip()
    unavailable_text = re.sub(r"\s+", " ", str(state.get("unavailableText") or "")).strip()
    expected_kind, expected_department, expected_node = _rank_category_parts(expected_url)
    actual_kind, actual_department, actual_node = _rank_category_parts(actual_url)

    if not expected_department or not expected_node:
        return False, f"{INVALID_RANK_CATEGORY_PREFIX}：请求链接不是具体榜单类目页。"
    if actual_kind != expected_kind or actual_node != expected_node:
        return (
            False,
            f"{INVALID_RANK_CATEGORY_PREFIX}：页面跳离目标类目 "
            f"({expected_kind}/{expected_department}/{expected_node} -> "
            f"{actual_kind or '-'}/{actual_department or '-'}"
            f"/{actual_node or '-'}）。",
        )
    if selected_text.lower().startswith("any department"):
        return False, f"{INVALID_RANK_CATEGORY_PREFIX}：Amazon 将入口回退到了 Any Department。"
    if unavailable_text:
        return False, f"{INVALID_RANK_CATEGORY_PREFIX}：{unavailable_text}"
    return True, ""


class CDPError(RuntimeError):
    pass


class CDPClient:
    def __init__(self, websocket_url: str, timeout: int = 15):
        self.websocket_url = websocket_url
        self.timeout = timeout
        self._id = 0
        self._socket = self._connect(websocket_url, timeout)

    def close(self):
        try:
            self._socket.close()
        except OSError:
            pass

    def command(self, method: str, params: dict | None = None) -> dict:
        self._id += 1
        message = {"id": self._id, "method": method}
        if params is not None:
            message["params"] = params
        self._send_json(message)
        deadline = time.time() + self.timeout
        while time.time() < deadline:
            payload = self._recv_json()
            if payload.get("id") != self._id:
                continue
            if "error" in payload:
                raise CDPError(json.dumps(payload["error"], ensure_ascii=False))
            return payload.get("result", {})
        raise CDPError(f"Timed out waiting for CDP command: {method}")

    def evaluate(self, expression: str, timeout: int | None = None):
        previous_timeout = self.timeout
        if timeout is not None:
            self.timeout = timeout
        try:
            result = self.command(
                "Runtime.evaluate",
                {
                    "expression": expression,
                    "awaitPromise": True,
                    "returnByValue": True,
                    "timeout": (timeout or previous_timeout) * 1000,
                },
            )
            remote = result.get("result", {})
            if "value" in remote:
                return remote["value"]
            return remote.get("description", "")
        finally:
            self.timeout = previous_timeout

    def _connect(self, websocket_url: str, timeout: int) -> socket.socket:
        parsed = urlparse(websocket_url)
        if parsed.scheme != "ws":
            raise CDPError("Only local ws:// CDP endpoints are supported.")
        host = parsed.hostname or "127.0.0.1"
        port = parsed.port or 80
        path = parsed.path or "/"
        if parsed.query:
            path += "?" + parsed.query
        sock = socket.create_connection((host, port), timeout=timeout)
        sock.settimeout(timeout)
        key = base64.b64encode(os.urandom(16)).decode("ascii")
        request = (
            f"GET {path} HTTP/1.1\r\n"
            f"Host: {host}:{port}\r\n"
            "Upgrade: websocket\r\n"
            "Connection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {key}\r\n"
            "Sec-WebSocket-Version: 13\r\n\r\n"
        )
        sock.sendall(request.encode("ascii"))
        response = b""
        while b"\r\n\r\n" not in response:
            chunk = sock.recv(4096)
            if not chunk:
                break
            response += chunk
        if b" 101 " not in response.split(b"\r\n", 1)[0]:
            raise CDPError("Chrome did not accept the WebSocket upgrade.")
        return sock

    def _send_json(self, payload: dict):
        data = json.dumps(payload, separators=(",", ":")).encode("utf-8")
        self._socket.sendall(_encode_client_frame(data))

    def _recv_json(self) -> dict:
        while True:
            opcode, payload = _read_frame(self._socket)
            if opcode == 0x1:
                return json.loads(payload.decode("utf-8", errors="replace"))
            if opcode == 0x8:
                raise CDPError("CDP WebSocket closed.")
            if opcode == 0x9:
                self._socket.sendall(_encode_client_frame(payload, opcode=0xA))


def _encode_client_frame(payload: bytes, opcode: int = 0x1) -> bytes:
    first = 0x80 | opcode
    length = len(payload)
    mask_bit = 0x80
    if length < 126:
        header = struct.pack("!BB", first, mask_bit | length)
    elif length < 65536:
        header = struct.pack("!BBH", first, mask_bit | 126, length)
    else:
        header = struct.pack("!BBQ", first, mask_bit | 127, length)
    mask = os.urandom(4)
    masked = bytes(byte ^ mask[index % 4] for index, byte in enumerate(payload))
    return header + mask + masked


def _read_frame(sock: socket.socket) -> tuple[int, bytes]:
    first_two = _recv_exact(sock, 2)
    first, second = first_two
    opcode = first & 0x0F
    length = second & 0x7F
    if length == 126:
        length = struct.unpack("!H", _recv_exact(sock, 2))[0]
    elif length == 127:
        length = struct.unpack("!Q", _recv_exact(sock, 8))[0]
    masked = bool(second & 0x80)
    mask = _recv_exact(sock, 4) if masked else b""
    payload = _recv_exact(sock, length) if length else b""
    if masked:
        payload = bytes(byte ^ mask[index % 4] for index, byte in enumerate(payload))
    return opcode, payload


def _recv_exact(sock: socket.socket, size: int) -> bytes:
    chunks = []
    remaining = size
    while remaining:
        chunk = sock.recv(remaining)
        if not chunk:
            raise CDPError("Unexpected end of WebSocket stream.")
        chunks.append(chunk)
        remaining -= len(chunk)
    return b"".join(chunks)


def chrome_debugger_available(port: int = DEFAULT_CDP_PORT) -> bool:
    try:
        chrome_json("/json/version", port=port, timeout=1)
        return True
    except Exception:
        return False


def chrome_json(path: str, port: int = DEFAULT_CDP_PORT, timeout: int = 5, method: str = "GET") -> dict | list:
    request = Request(f"http://127.0.0.1:{port}{path}", method=method)
    with urlopen(request, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8", errors="replace"))


def open_tab(url: str, port: int = DEFAULT_CDP_PORT) -> dict:
    encoded = quote(url, safe=":/?&=%#")
    try:
        return chrome_json(f"/json/new?{encoded}", port=port, timeout=8, method="PUT")
    except URLError:
        return chrome_json(f"/json/new?{encoded}", port=port, timeout=8, method="GET")


def close_tab(target_id: str | None, port: int = DEFAULT_CDP_PORT) -> None:
    if not target_id:
        return
    try:
        chrome_json(f"/json/close/{target_id}", port=port, timeout=3)
    except Exception:
        pass


def discover_bestseller_category_links(
    url: str,
    port: int = DEFAULT_CDP_PORT,
    max_links: int = 500,
    wait_seconds: float = 4.0,
) -> list[CategoryLink]:
    if not chrome_debugger_available(port):
        raise RuntimeError(f"未连接到 Chrome 调试端口 {port}。")

    target = open_tab(url, port=port)
    websocket_url = target.get("webSocketDebuggerUrl")
    if not websocket_url:
        raise RuntimeError("Chrome 已打开页面，但没有返回 CDP WebSocket 地址。")

    parsed_seed = urlparse(url)
    seed_without_ref = url.split("/ref=", 1)[0].rstrip("/")
    seed_match = re.search(r"/gp/(?:bestsellers|new-releases)/([^/?#]+)", parsed_seed.path)
    list_kind = "new-releases" if "/gp/new-releases/" in parsed_seed.path else "bestsellers"
    department_slug = seed_match.group(1) if seed_match else ""
    current_node_match = re.search(r"/(\d{5,})(?:/)?$", parsed_seed.path.rstrip("/"))
    current_node = current_node_match.group(1) if current_node_match else ""
    client = CDPClient(websocket_url, timeout=20)
    try:
        client.command("Runtime.enable")
        client.command("Page.enable")
        time.sleep(wait_seconds)
        raw_links = client.evaluate(
            r"""
            (() => {
              const browseGroups = Array.from(
                document.querySelectorAll('ul[class*="zg-browse-group"]')
              );
              const browseRoot =
                document.querySelector('#zg_browseRoot') ||
                document.querySelector('[data-csa-c-slot-id*="browse"]') ||
                document.querySelector('[class*="browseRoot"]') ||
                (browseGroups[0] && browseGroups[0].parentElement);
              if (!browseRoot) return [];

              const nodeFromHref = (href) => {
                try {
                  const url = new URL(href, location.href);
                  const pathMatch = url.pathname.match(/\/(\d{5,})(?:\/|$)/);
                  return pathMatch ? pathMatch[1] : (url.searchParams.get('node') || '');
                } catch {
                  return '';
                }
              };
              const directAnchor = (item) => {
                if (!item) return null;
                for (const child of Array.from(item.children || [])) {
                  if (child.tagName === 'A' && child.href) return child;
                  const anchor = child.querySelector && child.querySelector(':scope > a[href]');
                  if (anchor) return anchor;
                }
                return null;
              };

              const selected =
                browseRoot.querySelector('.zg_selected') ||
                browseRoot.querySelector('[aria-current="page"]') ||
                null;
              const childGroup = browseGroups[browseGroups.length - 1] || null;
              if (!childGroup || (selected && childGroup.contains(selected))) return [];
              const childItems = Array.from(childGroup.children || []).filter(
                (child) =>
                  child.tagName === 'LI' ||
                  child.getAttribute('role') === 'treeitem'
              );

              return childItems.map((item, index) => {
                const anchor = directAnchor(item);
                return {
                  index,
                  title: (anchor.innerText || anchor.textContent || anchor.getAttribute('aria-label') || '').trim(),
                  href: anchor.href,
                  node: nodeFromHref(anchor.href),
                  hasChildren: true
                };
              });
            })()
            """,
            timeout=20,
        ) or []
    finally:
        client.close()
        close_tab(target.get("id"), port=port)

    links: list[CategoryLink] = []
    seen: set[str] = set()
    ignored_titles = {
        "any department",
        "amazon best sellers",
        "best sellers",
        "main content",
        "sign in",
        "start here.",
        "deals & coupons",
        "subscribe & save",
        "pet care tips",
        "accessibility",
        "amazon devices",
        "become an amazon hub partner",
        "see more ways to make money",
        "credit card marketplace",
        "sell on amazon",
        "sell apps on amazon",
        "shop with points",
        "amazon business",
        "amazon global",
        "advertise your products",
        "self-publish with us",
        "host an amazon hub",
    }
    for item in raw_links:
        href = str(item.get("href") or "")
        title = re.sub(r"\s+", " ", str(item.get("title") or "")).strip()
        if not href or not title:
            continue
        lowered_title = title.lower()
        absolute = urljoin(f"{parsed_seed.scheme}://{parsed_seed.netloc}", href)
        normalized = absolute.split("/ref=", 1)[0].rstrip("/")
        parsed_link = urlparse(absolute)
        node = ""
        node_match = re.search(r"/(\d{5,})(?:/|$)", parsed_link.path)
        if node_match:
            node = node_match.group(1)
        if f"/gp/{list_kind}/" in absolute:
            if not node:
                node = parse_qs(parsed_link.query).get("node", [""])[0]
            if not node or not department_slug:
                continue
            category_url = absolute
        elif "/zgbs/" in parsed_link.path and department_slug and node:
            category_url = f"{parsed_seed.scheme}://{parsed_seed.netloc}/gp/{list_kind}/{department_slug}/{node}"
            normalized = category_url.rstrip("/")
        else:
            node = parse_qs(parsed_link.query).get("node", [""])[0]
            if not node or not department_slug:
                continue
            category_url = f"{parsed_seed.scheme}://{parsed_seed.netloc}/gp/{list_kind}/{department_slug}/{node}"
            normalized = category_url.rstrip("/")
        if node and node == current_node:
            continue
        if normalized == seed_without_ref or normalized in seen:
            continue
        if lowered_title in ignored_titles:
            continue
        seen.add(normalized)
        links.append(
            CategoryLink(
                title=title,
                url=category_url,
                node=node,
                depth=1,
                is_leaf=not bool(item.get("hasChildren")),
                path=title,
            )
        )
        if len(links) >= max_links:
            break
    return links


def refresh_sellersprite_cache(
    url: str,
    dom_cache_path: Path,
    image_cache_path: Path,
    meta_cache_path: Path,
    port: int = DEFAULT_CDP_PORT,
    expected_products: int = 50,
    max_rounds: int = 24,
    wait_seconds: float = 2.5,
    min_capture_seconds: float = 25.0,
    progress=None,
) -> ChromeRefreshResult:
    if not chrome_debugger_available(port):
        return ChromeRefreshResult(False, 0, 0, 0, url, f"\u672a\u8fde\u63a5\u5230 Chrome \u8c03\u8bd5\u7aef\u53e3 {port}\u3002\u8bf7\u5148\u542f\u52a8\u91c7\u96c6 Chrome\u3002")

    target = open_tab(url, port=port)
    websocket_url = target.get("webSocketDebuggerUrl")
    if not websocket_url:
        return ChromeRefreshResult(False, 0, 0, 0, url, "\u0043\u0068\u0072\u006f\u006d\u0065 \u5df2\u6253\u5f00\u9875\u9762\uff0c\u4f46\u6ca1\u6709\u8fd4\u56de CDP WebSocket \u5730\u5740\u3002")

    client = CDPClient(websocket_url, timeout=20)
    best_text = ""
    best_images: dict[str, str] = {}
    best_product_count = 0
    best_hydrated_count = 0
    best_score = (-1, -1, -1)
    last_signature = None
    stable_rounds = 0
    target_reached_rounds = 0
    try:
        client.command("Runtime.enable")
        client.command("Page.enable")
        opened_at = time.monotonic()
        _report(progress, 5, "\u0041\u006d\u0061\u007a\u006f\u006e \u9875\u9762\u5df2\u6253\u5f00\uff0c\u6b63\u5728\u7b49\u5f85\u9875\u9762\u548c\u5356\u5bb6\u7cbe\u7075\u52a0\u8f7d\u3002")
        time.sleep(8)
        _report(progress, 8, "\u9875\u9762\u57fa\u7840\u52a0\u8f7d\u5b8c\u6210\uff0c\u76f4\u63a5\u6eda\u52a8\u5230\u5e95\u90e8\u89e6\u53d1\u5356\u5bb6\u7cbe\u7075\u3002")
        client.evaluate(_SCROLL_BOTTOM_SCRIPT, timeout=10)
        time.sleep(2)
        for round_index in range(max_rounds):
            text = client.evaluate("document.body ? document.body.innerText : ''", timeout=20) or ""
            images = client.evaluate(_IMAGE_MAP_SCRIPT, timeout=20) or {}
            product_count = count_products(text)
            hydrated_count = count_hydrated_products(text)
            image_count = len(images) if isinstance(images, dict) else 0
            score = (hydrated_count, product_count, image_count)
            signature = (product_count, hydrated_count, image_count)
            if signature == last_signature:
                stable_rounds += 1
            else:
                stable_rounds = 0
                last_signature = signature
            if score >= best_score:
                best_text = text
                best_images = {asin: src for asin, src in images.items() if asin and src} if isinstance(images, dict) else {}
                best_product_count = product_count
                best_hydrated_count = hydrated_count
                best_score = score
            elapsed = time.monotonic() - opened_at
            wait_remaining = max(0, int(min_capture_seconds - elapsed))
            percent = min(95, int((best_hydrated_count / max(expected_products, 1)) * 90) + 5)
            wait_note = f"\uff0c\u7a33\u5b9a\u7b49\u5f85 {wait_remaining} \u79d2" if wait_remaining else ""
            _report(
                progress,
                percent,
                f"\u5356\u5bb6\u7cbe\u7075\u52a0\u8f7d\u4e2d\uff1a\u5df2\u8bc6\u522b {max(product_count, best_product_count)} \u6761\u4ea7\u54c1\uff0c"
                f"{best_hydrated_count} \u6761\u63d2\u4ef6\u5b57\u6bb5\u5b8c\u6574{wait_note}\u3002",
            )
            if best_hydrated_count >= expected_products:
                target_reached_rounds += 1
            if (
                best_hydrated_count >= expected_products
                and target_reached_rounds >= 2
                and stable_rounds >= 1
                and elapsed >= min_capture_seconds
            ):
                break
            client.evaluate(_SCROLL_BOTTOM_SCRIPT, timeout=10)
            time.sleep(wait_seconds)
        if not best_text:
            return ChromeRefreshResult(False, 0, 0, 0, url, "\u9875\u9762\u6587\u672c\u4e3a\u7a7a\uff0c\u53ef\u80fd\u9875\u9762\u672a\u52a0\u8f7d\u5b8c\u6210\u6216\u88ab\u9a8c\u8bc1\u7801\u62e6\u622a\u3002")
        _report(progress, 96, "\u6b63\u5728\u6eda\u52a8\u5230\u5e95\u90e8\u68c0\u67e5\u4e0b\u4e00\u9875\u3002")
        client.evaluate(_SCROLL_BOTTOM_SCRIPT, timeout=10)
        time.sleep(1.2)
        final_text = client.evaluate("document.body ? document.body.innerText : ''", timeout=20) or ""
        final_images = client.evaluate(_IMAGE_MAP_SCRIPT, timeout=20) or {}
        final_product_count = count_products(final_text)
        final_hydrated_count = count_hydrated_products(final_text)
        final_image_count = len(final_images) if isinstance(final_images, dict) else 0
        final_score = (final_hydrated_count, final_product_count, final_image_count)
        if final_score >= best_score:
            best_text = final_text
            best_images = {asin: src for asin, src in final_images.items() if asin and src} if isinstance(final_images, dict) else {}
            best_product_count = final_product_count
            best_hydrated_count = final_hydrated_count
            best_score = final_score
        next_page_url = client.evaluate(_NEXT_PAGE_SCRIPT, timeout=10) or ""
        dom_cache_path.parent.mkdir(parents=True, exist_ok=True)
        dom_cache_path.write_text(best_text, encoding="utf-8")
        image_cache_path.write_text(json.dumps(best_images, ensure_ascii=False, indent=2), encoding="utf-8")
        meta = {
            "source_url": url,
            "captured_at": datetime.now().isoformat(timespec="seconds"),
            "product_count": best_product_count,
            "image_count": len(best_images),
            "hydrated_count": best_hydrated_count,
            "next_page_url": next_page_url,
            "loaded_markers": list(SELLERSPRITE_MARKERS),
            "driver": f"local-chrome-cdp:{port}",
            "capture_seconds": round(time.monotonic() - opened_at, 1),
            "min_capture_seconds": min_capture_seconds,
        }
        meta_cache_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
        _report(
            progress,
            100,
            f"\u5237\u65b0\u5b8c\u6210\uff1a{best_product_count} \u6761\u4ea7\u54c1\uff0c{best_hydrated_count} \u6761\u63d2\u4ef6\u5b57\u6bb5\u5b8c\u6574\u3002",
        )
        required_count = best_product_count or expected_products
        ok = best_hydrated_count >= required_count
        message = "\u5237\u65b0\u5b8c\u6210\u3002" if ok else "\u5df2\u4fdd\u5b58\u5f53\u524d\u9875\u9762\u6570\u636e\uff0c\u4f46\u5356\u5bb6\u7cbe\u7075\u8865\u5145\u5b57\u6bb5\u4ecd\u672a\u5b8c\u5168\u52a0\u8f7d\u3002"
        return ChromeRefreshResult(ok, best_product_count, best_hydrated_count, len(best_images), url, message, next_page_url)
    finally:
        client.close()
        close_tab(target.get("id"), port=port)


def refresh_sellersprite_cache_pages(
    url: str,
    dom_cache_path: Path,
    image_cache_path: Path,
    meta_cache_path: Path,
    port: int = DEFAULT_CDP_PORT,
    expected_products: int = 50,
    page_count: int = 2,
    progress=None,
    page_callback=None,
    stop_check=None,
) -> list[ChromeRefreshResult]:
    _check_stop(stop_check)
    if not is_rank_category_url(url):
        return [ChromeRefreshResult(False, 0, 0, 0, url, "\u8df3\u8fc7\u975e\u5177\u4f53\u699c\u5355\u7c7b\u76ee\u9875\uff1a\u8fd9\u7c7b\u603b\u5165\u53e3\u9875\u5356\u5bb6\u7cbe\u7075\u4e0d\u4f1a\u52a0\u8f7d\u4ea7\u54c1\u6570\u636e\u3002")]
    if not chrome_debugger_available(port):
        return [ChromeRefreshResult(False, 0, 0, 0, url, f"\u672a\u8fde\u63a5\u5230 Chrome \u8c03\u8bd5\u7aef\u53e3 {port}\u3002\u8bf7\u5148\u542f\u52a8\u91c7\u96c6 Chrome\u3002")]

    target = open_tab(url, port=port)
    websocket_url = target.get("webSocketDebuggerUrl")
    if not websocket_url:
        return [ChromeRefreshResult(False, 0, 0, 0, url, "\u0043\u0068\u0072\u006f\u006d\u0065 \u5df2\u6253\u5f00\u9875\u9762\uff0c\u4f46\u6ca1\u6709\u8fd4\u56de CDP WebSocket \u5730\u5740\u3002")]

    client = CDPClient(websocket_url, timeout=20)
    results: list[ChromeRefreshResult] = []
    try:
        client.command("Runtime.enable")
        client.command("Page.enable")
        for page in range(1, page_count + 1):
            _check_stop(stop_check)
            page_name = "\u7b2c\u4e00\u9875" if page == 1 else "\u7b2c\u4e8c\u9875" if page == 2 else f"\u7b2c {page} \u9875"

            def page_progress(percent: int, message: str, page_name=page_name):
                _report(progress, percent, f"{page_name}\uff5c{message}")

            result = _capture_current_sellersprite_page(
                client,
                dom_cache_path,
                image_cache_path,
                meta_cache_path,
                expected_category_url=url,
                expected_products=expected_products,
                progress=page_progress,
                stop_check=stop_check,
            )
            _check_stop(stop_check)
            results.append(result)
            if page_callback:
                page_callback(page, result)
            _check_stop(stop_check)
            if result.message == EMPTY_NEW_RELEASES_MESSAGE:
                break
            if page >= page_count:
                break
            clicked = client.evaluate(_CLICK_NEXT_PAGE_SCRIPT, timeout=10)
            _check_stop(stop_check)
            if not isinstance(clicked, dict) or not clicked.get("clicked"):
                results.append(ChromeRefreshResult(False, 0, 0, 0, result.source_url, "\u672a\u627e\u5230\u53ef\u70b9\u51fb\u7684\u4e0b\u4e00\u9875\u6309\u94ae\u3002"))
                break
            _report(progress, 99, f"{page_name}\uff5c\u5207\u5230\u4e0b\u4e00\u9875")
            _interruptible_sleep(8, stop_check)
            current_url = client.evaluate("location.href", timeout=10) or ""
            _check_stop(stop_check)
            if not is_rank_category_url(current_url):
                results.append(
                    ChromeRefreshResult(
                        False,
                        0,
                        0,
                        0,
                        current_url,
                        "\u7ffb\u9875\u540e\u79bb\u5f00\u4e86 Amazon \u699c\u5355\u7c7b\u76ee\u9875\uff0c\u5df2\u4e2d\u6b62\u5f53\u524d\u5165\u53e3\uff0c\u907f\u514d\u8bef\u91c7\u5546\u54c1\u8be6\u60c5\u9875\u3002",
                    )
                )
                break
        return results
    finally:
        client.close()
        close_tab(target.get("id"), port=port)


def _capture_current_sellersprite_page(
    client: CDPClient,
    dom_cache_path: Path,
    image_cache_path: Path,
    meta_cache_path: Path,
    expected_category_url: str = "",
    expected_products: int = 50,
    max_rounds: int = 24,
    wait_seconds: float = 2.5,
    min_capture_seconds: float = 25.0,
    progress=None,
    stop_check=None,
) -> ChromeRefreshResult:
    _check_stop(stop_check)
    source_url = client.evaluate("location.href", timeout=10) or ""
    best_text = ""
    best_images: dict[str, str] = {}
    seen_texts: list[str] = []
    seen_images: dict[str, str] = {}
    best_product_count = 0
    best_hydrated_count = 0
    best_score = (-1, -1, -1)
    last_signature = None
    stable_rounds = 0
    target_reached_rounds = 0
    opened_at = time.monotonic()

    def observe_best():
        nonlocal best_text, best_images, best_product_count, best_hydrated_count, best_score
        _check_stop(stop_check)
        best_text, best_images, best_product_count, best_hydrated_count, best_score = _observe_sellersprite_page(
            client, best_text, best_images, best_product_count, best_hydrated_count, best_score, seen_texts, seen_images
        )
        _check_stop(stop_check)

    def page_has_enough_products() -> bool:
        return best_product_count >= expected_products and best_hydrated_count >= expected_products

    _report(progress, 5, "顶部加载：等待 Amazon 页面和卖家精灵插件出现（约 5 秒）")
    client.evaluate(_SCROLL_TOP_SCRIPT, timeout=10)
    _interruptible_sleep(5, stop_check)
    observe_best()
    page_state = client.evaluate(_RANK_PAGE_STATE_SCRIPT, timeout=10) or {}
    page_valid, page_error = validate_rank_category_page(
        expected_category_url or source_url,
        page_state,
    )
    if not page_valid:
        return ChromeRefreshResult(
            False,
            best_product_count,
            best_hydrated_count,
            len(best_images),
            str(page_state.get("url") or source_url),
            page_error,
        )
    _report(progress, 18, f"顶部检测：页面产品 {best_product_count} 条，卖家精灵字段完整 {best_hydrated_count} 条")

    _report(progress, 35, "中部加载：滚到页面中部，等待懒加载商品和插件字段（约 5 秒）")
    client.evaluate(_SCROLL_MIDDLE_SCRIPT, timeout=10)
    _interruptible_sleep(5, stop_check)
    observe_best()
    _report(progress, 48, f"中部检测：页面产品 {best_product_count} 条，卖家精灵字段完整 {best_hydrated_count} 条")

    _report(progress, 65, "底部加载：滚到页码区域，触发本页剩余商品和插件字段（约 5 秒）")
    client.evaluate(_SCROLL_TO_PAGINATION_SCRIPT, timeout=10)
    _interruptible_sleep(5, stop_check)
    observe_best()
    _report(progress, 78, f"底部检测：页面产品 {best_product_count} 条，卖家精灵字段完整 {best_hydrated_count} 条")

    _report(progress, 86, "最后补触发：轻微补滚一次，避免底部商品漏加载（约 5 秒）")
    client.evaluate(_BOTTOM_NUDGE_SCRIPT, timeout=10)
    _interruptible_sleep(5, stop_check)
    observe_best()
    _report(progress, 90, f"最后检测：页面产品 {best_product_count} 条，卖家精灵字段完整 {best_hydrated_count} 条")
    monitor_rounds = 2 if page_has_enough_products() else min(max_rounds, 8)
    for _round_index in range(monitor_rounds):
        _check_stop(stop_check)
        text = client.evaluate("document.body ? document.body.innerText : ''", timeout=20) or ""
        images = client.evaluate(_IMAGE_MAP_SCRIPT, timeout=20) or {}
        _remember_sellersprite_snapshot(text, images, seen_texts, seen_images)
        product_count = count_products(text)
        hydrated_count = count_hydrated_products(text)
        image_count = len(images) if isinstance(images, dict) else 0
        score = (hydrated_count, product_count, image_count)
        signature = (product_count, hydrated_count, image_count)
        if signature == last_signature:
            stable_rounds += 1
        else:
            stable_rounds = 0
            last_signature = signature
        if score >= best_score:
            best_text = text
            best_images = {asin: src for asin, src in images.items() if asin and src} if isinstance(images, dict) else {}
            best_product_count = product_count
            best_hydrated_count = hydrated_count
            best_score = score
        percent = min(95, int((best_hydrated_count / max(expected_products, 1)) * 90) + 5)
        wait_note = "复查稳定性" if page_has_enough_products() else "继续等待插件字段"
        _report(
            progress,
            percent,
            f"{wait_note}：页面产品 {max(product_count, best_product_count)} 条，卖家精灵字段完整 {best_hydrated_count} 条",
        )
        if page_has_enough_products():
            target_reached_rounds += 1
        if (
            page_has_enough_products()
            and target_reached_rounds >= 2
            and stable_rounds >= 1
        ):
            break
        _interruptible_sleep(wait_seconds, stop_check)
    _check_stop(stop_check)
    if not best_text:
        return ChromeRefreshResult(False, 0, 0, 0, source_url, "\u9875\u9762\u6587\u672c\u4e3a\u7a7a\uff0c\u53ef\u80fd\u9875\u9762\u672a\u52a0\u8f7d\u5b8c\u6210\u6216\u88ab\u9a8c\u8bc1\u7801\u62e6\u622a\u3002")
    _report(progress, 96, "复查数据：读取最终页面文本、图片和下一页按钮")
    final_text = client.evaluate("document.body ? document.body.innerText : ''", timeout=20) or ""
    final_images = client.evaluate(_IMAGE_MAP_SCRIPT, timeout=20) or {}
    _remember_sellersprite_snapshot(final_text, final_images, seen_texts, seen_images)
    final_product_count = count_products(final_text)
    final_hydrated_count = count_hydrated_products(final_text)
    final_image_count = len(final_images) if isinstance(final_images, dict) else 0
    final_score = (final_hydrated_count, final_product_count, final_image_count)
    if final_score >= best_score:
        best_text = final_text
        best_images = {asin: src for asin, src in final_images.items() if asin and src} if isinstance(final_images, dict) else {}
        best_product_count = final_product_count
        best_hydrated_count = final_hydrated_count
        best_score = final_score
    next_page_url = client.evaluate(_NEXT_PAGE_SCRIPT, timeout=10) or ""
    source_url = client.evaluate("location.href", timeout=10) or source_url
    combined_text = "\n\n".join(seen_texts) if seen_texts else best_text
    combined_product_count = count_products(combined_text)
    combined_hydrated_count = count_hydrated_products(combined_text)
    if combined_product_count >= best_product_count:
        best_text = combined_text
        best_images = seen_images or best_images
        best_product_count = combined_product_count
        best_hydrated_count = max(best_hydrated_count, combined_hydrated_count)
    dom_cache_path.parent.mkdir(parents=True, exist_ok=True)
    dom_cache_path.write_text(best_text, encoding="utf-8")
    image_cache_path.write_text(json.dumps(best_images, ensure_ascii=False, indent=2), encoding="utf-8")
    meta = {
        "source_url": source_url,
        "captured_at": datetime.now().isoformat(timespec="seconds"),
        "product_count": best_product_count,
        "image_count": len(best_images),
        "hydrated_count": best_hydrated_count,
        "next_page_url": next_page_url,
        "loaded_markers": list(SELLERSPRITE_MARKERS),
        "driver": f"local-chrome-cdp:{DEFAULT_CDP_PORT}",
        "capture_seconds": round(time.monotonic() - opened_at, 1),
        "min_capture_seconds": min_capture_seconds,
    }
    meta_cache_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    _report(progress, 100, f"本页完成：页面产品 {best_product_count} 条，卖家精灵字段完整 {best_hydrated_count} 条")
    empty_new_releases = EMPTY_NEW_RELEASES_TEXT in best_text.lower()
    required_count = best_product_count or expected_products
    ok = empty_new_releases or best_hydrated_count >= required_count
    if empty_new_releases:
        message = EMPTY_NEW_RELEASES_MESSAGE
    else:
        message = "\u5237\u65b0\u5b8c\u6210\u3002" if ok else "\u5df2\u4fdd\u5b58\u5f53\u524d\u9875\u9762\u6570\u636e\uff0c\u4f46\u5356\u5bb6\u7cbe\u7075\u8865\u5145\u5b57\u6bb5\u4ecd\u672a\u5b8c\u5168\u52a0\u8f7d\u3002"
    return ChromeRefreshResult(ok, best_product_count, best_hydrated_count, len(best_images), source_url, message, next_page_url)


def _observe_sellersprite_page(
    client: CDPClient,
    best_text: str,
    best_images: dict[str, str],
    best_product_count: int,
    best_hydrated_count: int,
    best_score: tuple[int, int, int],
    seen_texts: list[str] | None = None,
    seen_images: dict[str, str] | None = None,
) -> tuple[str, dict[str, str], int, int, tuple[int, int, int]]:
    text = client.evaluate("document.body ? document.body.innerText : ''", timeout=20) or ""
    images = client.evaluate(_IMAGE_MAP_SCRIPT, timeout=20) or {}
    _remember_sellersprite_snapshot(text, images, seen_texts, seen_images)
    product_count = count_products(text)
    hydrated_count = count_hydrated_products(text)
    image_count = len(images) if isinstance(images, dict) else 0
    score = (hydrated_count, product_count, image_count)
    if score >= best_score:
        return (
            text,
            {asin: src for asin, src in images.items() if asin and src} if isinstance(images, dict) else {},
            product_count,
            hydrated_count,
            score,
        )
    return best_text, best_images, best_product_count, best_hydrated_count, best_score


def _remember_sellersprite_snapshot(
    text: str,
    images: dict | None,
    seen_texts: list[str] | None,
    seen_images: dict[str, str] | None,
) -> None:
    if seen_texts is not None and text and text not in seen_texts:
        seen_texts.append(text)
    if seen_images is not None and isinstance(images, dict):
        for asin, src in images.items():
            if asin and src and asin not in seen_images:
                seen_images[asin] = src


_RANK_PAGE_STATE_SCRIPT = r"""
(() => {
  const selected =
    document.querySelector('#zg_browseRoot .zg_selected') ||
    document.querySelector('#zg_browseRoot [aria-current="page"]') ||
    document.querySelector('ul[class*="zg-browse-group"] .zg_selected') ||
    document.querySelector('ul[class*="zg-browse-group"] [aria-current="page"]');
  const bodyText = document.body ? document.body.innerText : "";
  const unavailablePhrases = [
    "sorry, we couldn't find that page",
    "page not found",
    "the web address you entered is not a functioning page on our site"
  ];
  const lowered = bodyText.toLowerCase();
  return {
    url: location.href,
    selectedText: selected ? (selected.innerText || selected.textContent || "").trim() : "",
    unavailableText:
      unavailablePhrases.find((phrase) => lowered.includes(phrase)) || ""
  };
})()
"""


_SCROLL_BOTTOM_SCRIPT = r"""
(() => {
  const scroller = document.scrollingElement || document.documentElement || document.body;
  scroller.scrollTop = scroller.scrollHeight;
  window.dispatchEvent(new Event('scroll'));
  return scroller.scrollTop;
})()
"""


_LEGACY_NEXT_PAGE_SCRIPT = r"""
(() => {
  const current = new URL(location.href);
  const currentPage = parseInt(current.searchParams.get("pg") || "1", 10);
  const nextPage = currentPage + 1;
  const direct = document.querySelector("li.a-last:not(.a-disabled) a[href], .a-pagination .a-last:not(.a-disabled) a[href]");
  if (direct) {
    try { return new URL(direct.href, location.href).href; } catch {}
  }
  const links = Array.from(document.querySelectorAll("a[href]"));
  for (const link of links) {
    const text = `${link.innerText || ""} ${link.getAttribute("aria-label") || ""}`.trim().toLowerCase();
    if (link.getAttribute("aria-disabled") === "true" || link.classList.contains("a-disabled")) continue;
    let href;
    try { href = new URL(link.href, location.href); } catch { continue; }
    const linkPage = parseInt(href.searchParams.get("pg") || "0", 10);
    if (text === "next" || text.includes("next") || text.includes("下一页") || linkPage === nextPage) {
      return href.href;
    }
  }
  if (currentPage < 2) {
    current.searchParams.set("pg", String(nextPage));
    current.pathname = current.pathname.replace(/\/ref=[^/?#]+/, `/ref=zg_bs_pg_${nextPage}`);
    return current.href;
  }
  return "";
})()
"""


_LEGACY_CLICK_NEXT_PAGE_SCRIPT = r"""
(() => {
  const selectors = [
    "li.a-last:not(.a-disabled) a[href]",
    ".a-pagination .a-last:not(.a-disabled) a[href]"
  ];
  for (const selector of selectors) {
    const link = document.querySelector(selector);
    if (link) {
      link.scrollIntoView({block: "center", inline: "center"});
      link.click();
      return {clicked: true, href: link.href || "", method: selector};
    }
  }
  const links = Array.from(document.querySelectorAll("a[href]"));
  for (const link of links) {
    const text = `${link.innerText || ""} ${link.getAttribute("aria-label") || ""}`.trim().toLowerCase();
    const disabled = link.getAttribute("aria-disabled") === "true" || link.classList.contains("a-disabled") || link.closest(".a-disabled");
    if (disabled) continue;
    if (text === "next" || text.includes("next") || text.includes("下一页")) {
      link.scrollIntoView({block: "center", inline: "center"});
      link.click();
      return {clicked: true, href: link.href || "", method: "text"};
    }
  }
  return {clicked: false, href: "", method: ""};
})()
"""


_SCROLL_BOTTOM_SCRIPT = r"""
(() => {
  const scroller = document.scrollingElement || document.documentElement || document.body;
  scroller.scrollTop = scroller.scrollHeight;
  window.dispatchEvent(new Event("scroll"));
  return scroller.scrollTop;
})()
"""


_SCROLL_TOP_SCRIPT = r"""
(() => {
  const scroller = document.scrollingElement || document.documentElement || document.body;
  scroller.scrollTop = 0;
  window.dispatchEvent(new Event("scroll"));
  return scroller.scrollTop;
})()
"""


_SCROLL_MIDDLE_SCRIPT = r"""
(() => {
  const scroller = document.scrollingElement || document.documentElement || document.body;
  scroller.scrollTop = Math.floor(scroller.scrollHeight * 0.5);
  window.dispatchEvent(new Event("scroll"));
  return scroller.scrollTop;
})()
"""


_SCROLL_TO_PAGINATION_SCRIPT = r"""
(() => {
  const scroller = document.scrollingElement || document.documentElement || document.body;
  const pagination =
    document.querySelector("ul.a-pagination") ||
    document.querySelector("li.a-last") ||
    Array.from(document.querySelectorAll("a, span")).find((node) => {
      const text = (node.innerText || node.textContent || "").trim().toLowerCase();
      return text === "next" || text.includes("next page") || text === "previous";
    });
  if (pagination) {
    pagination.scrollIntoView({block: "center", inline: "nearest"});
  } else {
    scroller.scrollTop = scroller.scrollHeight;
  }
  window.dispatchEvent(new Event("scroll"));
  return {
    top: scroller.scrollTop,
    hasPagination: Boolean(pagination),
    productCount: (document.body.innerText.match(/ASIN:\s*[A-Z0-9]{10}/g) || []).length
  };
})()
"""


_BOTTOM_NUDGE_SCRIPT = r"""
(() => {
  const scroller = document.scrollingElement || document.documentElement || document.body;
  const pagination = document.querySelector("ul.a-pagination") || document.querySelector("li.a-last");
  if (pagination) {
    pagination.scrollIntoView({block: "end", inline: "nearest"});
    scroller.scrollTop = Math.min(
      scroller.scrollHeight - scroller.clientHeight,
      scroller.scrollTop + Math.floor(window.innerHeight * 0.25)
    );
  } else {
    scroller.scrollTop = Math.min(scroller.scrollHeight, scroller.scrollTop + Math.floor(window.innerHeight * 0.35));
  }
  window.dispatchEvent(new Event("scroll"));
  return {
    top: scroller.scrollTop,
    productCount: (document.body.innerText.match(/ASIN:\s*[A-Z0-9]{10}/g) || []).length
  };
})()
"""


_SCROLL_THROUGH_PRODUCTS_SCRIPT = r"""
(async () => {
  const scroller = document.scrollingElement || document.documentElement || document.body;
  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  scroller.scrollTop = 0;
  window.dispatchEvent(new Event("scroll"));
  await sleep(450);
  const step = Math.max(420, Math.floor(window.innerHeight * 0.82));
  let sawPagination = false;
  for (let i = 0; i < 18; i += 1) {
    const pagination = document.querySelector("ul.a-pagination") || document.querySelector("li.a-last");
    if (pagination) {
      const rect = pagination.getBoundingClientRect();
      if (rect.top < window.innerHeight * 0.92) {
        sawPagination = true;
        pagination.scrollIntoView({block: "center", inline: "nearest"});
        window.dispatchEvent(new Event("scroll"));
        await sleep(900);
        break;
      }
    }
    const before = scroller.scrollTop;
    scroller.scrollTop = Math.min(scroller.scrollHeight, scroller.scrollTop + step);
    window.dispatchEvent(new Event("scroll"));
    await sleep(650);
    if (Math.abs(scroller.scrollTop - before) < 3) break;
  }
  if (!sawPagination) {
    scroller.scrollTop = scroller.scrollHeight;
    window.dispatchEvent(new Event("scroll"));
    await sleep(900);
  }
  return {
    top: scroller.scrollTop,
    sawPagination,
    productCount: (document.body.innerText.match(/ASIN:\s*[A-Z0-9]{10}/g) || []).length
  };
})()
"""


_SWEEP_PAGE_SCRIPT = r"""
(async () => {
  const scroller = document.scrollingElement || document.documentElement || document.body;
  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const stops = [0, Math.floor(scroller.scrollHeight * 0.35), Math.floor(scroller.scrollHeight * 0.7), scroller.scrollHeight];
  for (const stop of stops) {
    scroller.scrollTop = stop;
    window.dispatchEvent(new Event("scroll"));
    await sleep(350);
  }
  return scroller.scrollTop;
})()
"""


_NEXT_PAGE_SCRIPT = r"""
(() => {
  const rankIdentity = (value) => {
    try {
      const url = new URL(value, location.href);
      const gp = url.pathname.match(/\/gp\/(bestsellers|new-releases)\/([^/?#]+)(?:\/(\d{5,}))?/);
      if (gp) {
        return {
          kind: gp[1],
          department: gp[2],
          node: gp[3] || url.searchParams.get("node") || ""
        };
      }
      const zgbs = url.pathname.match(/\/zgbs\/([^/?#]+)\/(\d{5,})(?:\/|$)/);
      if (zgbs) {
        return {kind: "bestsellers", department: zgbs[1], node: zgbs[2]};
      }
    } catch {}
    return {kind: "", department: "", node: ""};
  };
  const current = rankIdentity(location.href);
  const selectors = [
    ".a-pagination li.a-last:not(.a-disabled) a[href]",
    "li.a-last:not(.a-disabled) a[href]",
    "a.s-pagination-next:not(.s-pagination-disabled)[href]",
    "a[aria-label='Go to next page'][href]"
  ];
  for (const selector of selectors) {
    const link = document.querySelector(selector);
    if (!link) continue;
    try {
      const url = new URL(link.href, location.href);
      const target = rankIdentity(url.href);
      const isSameCategory = Boolean(
        current.kind &&
        current.node &&
        target.kind === current.kind &&
        target.node === current.node
      );
      const isNextPage = url.searchParams.get("pg") === "2" || /(?:^|_)pg_2(?:_|$)/.test(url.pathname + url.search);
      if (isSameCategory && isNextPage) return url.href;
    } catch {}
  }
  return "";
})()
"""


_CLICK_NEXT_PAGE_SCRIPT = r"""
(() => {
  const rankIdentity = (value) => {
    try {
      const url = new URL(value, location.href);
      const gp = url.pathname.match(/\/gp\/(bestsellers|new-releases)\/([^/?#]+)(?:\/(\d{5,}))?/);
      if (gp) {
        return {
          kind: gp[1],
          department: gp[2],
          node: gp[3] || url.searchParams.get("node") || ""
        };
      }
      const zgbs = url.pathname.match(/\/zgbs\/([^/?#]+)\/(\d{5,})(?:\/|$)/);
      if (zgbs) {
        return {kind: "bestsellers", department: zgbs[1], node: zgbs[2]};
      }
    } catch {}
    return {kind: "", department: "", node: ""};
  };
  const current = rankIdentity(location.href);
  const selectors = [
    ".a-pagination li.a-last:not(.a-disabled) a[href]",
    "li.a-last:not(.a-disabled) a[href]",
    "a.s-pagination-next:not(.s-pagination-disabled)[href]",
    "a[aria-label='Go to next page'][href]"
  ];
  for (const selector of selectors) {
    const link = document.querySelector(selector);
    if (!link) continue;
    try {
      const url = new URL(link.href, location.href);
      const target = rankIdentity(url.href);
      const isSameCategory = Boolean(
        current.kind &&
        current.node &&
        target.kind === current.kind &&
        target.node === current.node
      );
      const isNextPage = url.searchParams.get("pg") === "2" || /(?:^|_)pg_2(?:_|$)/.test(url.pathname + url.search);
      if (!isSameCategory || !isNextPage) continue;
      link.scrollIntoView({block: "center", inline: "center"});
      link.click();
      return {clicked: true, href: url.href, method: selector};
    } catch {}
  }
  return {clicked: false, href: "", method: ""};
})()
"""


def count_products(text: str) -> int:
    return len(set(re.findall(r"ASIN:\s*([A-Z0-9]{10})", text)))


def count_hydrated_products(text: str) -> int:
    blocks = re.split(r"\n#\d+\n", "\n" + text)
    hydrated_asins: set[str] = set()
    for block in blocks:
        asin_match = re.search(r"ASIN:\s*([A-Z0-9]{10})", block)
        if not asin_match:
            continue
        has_parent_sales = re.search(r"近30天销量\(父体\):", block)
        if has_parent_sales:
            hydrated_asins.add(asin_match.group(1))
    return len(hydrated_asins)


def _report(progress, percent: int, message: str):
    if progress:
        progress(percent, message)


_IMAGE_MAP_SCRIPT = r"""
(() => {
  const result = {};
  const asinPattern = /(?:ASIN:?\s*|\/dp\/|\/gp\/product\/)([A-Z0-9]{10})/;
  const images = Array.from(document.images || []);
  for (const img of images) {
    const src = img.currentSrc || img.src || "";
    if (!src || src.startsWith("data:")) continue;
    if (!/m\.media-amazon\.com\/images\/I\/|images-na\.ssl-images-amazon\.com\/images\/I\/|ssl-images-amazon\.com\/images\/I\/|images-amazon\.com\/images\/I\//i.test(src)) continue;
    if (img.naturalWidth && img.naturalWidth < 80) continue;
    if (img.naturalHeight && img.naturalHeight < 80) continue;
    let asin = "";
    let node = img;
    for (let depth = 0; node && depth < 8; depth += 1, node = node.parentElement) {
      const dataAsin = node.getAttribute && node.getAttribute("data-asin");
      if (dataAsin && /^[A-Z0-9]{10}$/.test(dataAsin)) {
        asin = dataAsin;
        break;
      }
      const text = (node.innerText || node.textContent || "").slice(0, 2000);
      const textMatch = text.match(asinPattern);
      if (textMatch) {
        asin = textMatch[1];
        break;
      }
      const link = node.querySelector && node.querySelector('a[href*="/dp/"], a[href*="/gp/product/"]');
      const href = link ? link.href : "";
      const hrefMatch = href.match(asinPattern);
      if (hrefMatch) {
        asin = hrefMatch[1];
        break;
      }
    }
    if (asin && !result[asin]) result[asin] = src;
  }
  return result;
})()
"""
