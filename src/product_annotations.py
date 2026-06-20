import json
from datetime import datetime
from pathlib import Path
from typing import Any


def _load_mapping(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}
    return data if isinstance(data, dict) else {}


def _write_mapping(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_favorites(path: Path) -> dict[str, dict[str, Any]]:
    data = _load_mapping(path)
    return {
        str(asin): record
        for asin, record in data.items()
        if isinstance(record, dict) and str(asin).strip()
    }


def save_favorites(path: Path, favorites: dict[str, dict[str, Any]]) -> None:
    saveable = {
        str(asin): record
        for asin, record in favorites.items()
        if str(asin).strip() and isinstance(record, dict)
    }
    _write_mapping(path, saveable)


def load_notes(path: Path) -> dict[str, str]:
    data = _load_mapping(path)
    return {
        str(asin): str(note)
        for asin, note in data.items()
        if str(asin).strip() and str(note).strip()
    }


def save_notes(path: Path, notes: dict[str, str]) -> None:
    saveable = {
        str(asin): str(note).strip()
        for asin, note in notes.items()
        if str(asin).strip() and str(note).strip()
    }
    _write_mapping(path, saveable)


def set_product_note(notes: dict[str, str], asin: str, note: str) -> None:
    asin = str(asin or "").strip()
    if not asin:
        return
    clean_note = str(note or "").strip()
    if clean_note:
        notes[asin] = clean_note
    else:
        notes.pop(asin, None)


def upsert_favorite(
    favorites: dict[str, dict[str, Any]],
    product: dict[str, Any],
    source_label: str = "",
    source_url: str = "",
    now: str | None = None,
) -> dict[str, Any] | None:
    asin = str(product.get("asin") or "").strip()
    if not asin:
        return None
    timestamp = now or datetime.now().isoformat(timespec="seconds")
    source = {
        "label": str(source_label or product.get("category_path") or "").strip(),
        "url": str(source_url or product.get("amazon_url") or "").strip(),
        "saved_at": timestamp,
    }
    record = favorites.get(asin) or {
        "asin": asin,
        "favorited_at": timestamp,
        "product": {},
        "sources": [],
    }
    record["product"] = dict(product)
    record["product"]["asin"] = asin
    sources = [
        item
        for item in record.get("sources", [])
        if isinstance(item, dict)
    ]
    if source["label"] or source["url"]:
        source_key = (source["label"], source["url"])
        if source_key not in {(item.get("label", ""), item.get("url", "")) for item in sources}:
            sources.append(source)
    record["sources"] = sources
    favorites[asin] = record
    return record


def remove_favorite(favorites: dict[str, dict[str, Any]], asin: str) -> None:
    favorites.pop(str(asin or "").strip(), None)
