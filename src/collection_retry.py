from __future__ import annotations

from dataclasses import dataclass
from collections.abc import Iterator
from typing import Any


@dataclass
class DeferredCollectionSeed:
    label: str
    url: str
    pages: int
    products: list[Any]
    error: str


class CollectionRetryQueue:
    def __init__(self) -> None:
        self._items: list[DeferredCollectionSeed] = []

    def defer(
        self,
        *,
        label: str,
        url: str,
        pages: int,
        products: list[Any],
        error: str,
    ) -> None:
        self._items.append(
            DeferredCollectionSeed(
                label=label,
                url=url,
                pages=pages,
                products=list(products),
                error=error,
            )
        )

    def __len__(self) -> int:
        return len(self._items)

    def __iter__(self) -> Iterator[DeferredCollectionSeed]:
        return iter(tuple(self._items))
