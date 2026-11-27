"""Klient lokalnego API Zotero 10 implementujący kontrakt `ZoteroSource`."""

from __future__ import annotations

import re
from typing import Any, Sequence

import requests

from azp.contracts import (
    STYLE_APA,
    STYLE_IEEE_V4,
    FormattedRef,
    StyleInfo,
    ZoteroItem,
    ZoteroOfflineError,
    spans_to_plain_text,
)
from azp.zotero.html_parser import parse_csl_bib_html, parse_csl_citation_html

DEFAULT_BASE_URL = "http://127.0.0.1:23119/api/users/0"
DEFAULT_TIMEOUT = 5.0
OFFLINE_MESSAGE = (
    "Zotero nie odpowiada. Uruchom Zotero 10 i upewnij się, że lokalne API jest włączone."
)

_EXCLUDED_ITEM_TYPES = frozenset({"attachment", "note", "annotation"})
_YEAR_RE = re.compile(r"(?<!\d)(\d{4})(?!\d)")
_BATCH_SIZE = 50

_DEFAULT_STYLES: tuple[StyleInfo, ...] = (
    StyleInfo(
        style_id=STYLE_IEEE_V4,
        title="IEEE Praca Inzynierska Access Date v4",
        is_numeric=True,
    ),
    StyleInfo(
        style_id=STYLE_APA,
        title="APA Style 7th edition",
        is_numeric=False,
    ),
)


def _extract_year(date_str: str, parsed_date: str = "") -> str:
    """Wyciąga 4-cyfrowy rok z pola `date` lub `meta.parsedDate` Zotero."""
    for candidate in (date_str, parsed_date):
        if not candidate:
            continue
        match = _YEAR_RE.search(str(candidate))
        if match:
            return match.group(1)
    return ""


def _format_single_creator(creator: dict[str, Any]) -> str:
    name = str(creator.get("name") or "").strip()
    if name:
        return name
    last_name = str(creator.get("lastName") or "").strip()
    first_name = str(creator.get("firstName") or "").strip()
    if last_name and first_name:
        return f"{last_name}, {first_name}"
    return last_name or first_name


def _parse_creators(creators_raw: Any) -> tuple[str, ...]:
    """Parsuje tablicę `data.creators` z Zotero do krotki czytelnych nazw autorów."""
    if not isinstance(creators_raw, list):
        return ()

    valid_creators = [c for c in creators_raw if isinstance(c, dict)]
    primary = [c for c in valid_creators if c.get("creatorType") == "author"]
    selected = primary if primary else valid_creators

    authors: list[str] = []
    for creator in selected:
        formatted = _format_single_creator(creator)
        if formatted:
            authors.append(formatted)
    return tuple(authors)


def _parse_zotero_item(raw: dict[str, Any]) -> ZoteroItem | None:
    """Konwertuje pojedynczy obiekt JSON z Zotero na `ZoteroItem` lub `None` dla typów pomocniczych."""
    data = raw.get("data")
    if not isinstance(data, dict):
        data = raw
    meta = raw.get("meta")
    if not isinstance(meta, dict):
        meta = {}

    item_type = str(data.get("itemType") or "journalArticle").strip()
    if item_type in _EXCLUDED_ITEM_TYPES:
        return None

    key = str(raw.get("key") or data.get("key") or "").strip()
    if not key:
        return None

    title = str(data.get("title") or "").strip()
    authors = _parse_creators(data.get("creators"))
    year = _extract_year(
        str(data.get("date") or ""),
        str(meta.get("parsedDate") or ""),
    )

    return ZoteroItem(
        key=key,
        title=title,
        authors=authors,
        year=year,
        item_type=item_type,
    )


class ZoteroClient:
    """Klient lokalnego API REST Zotero 10 (`127.0.0.1:23119/api/users/0`)."""

    def __init__(
        self,
        base_url: str = DEFAULT_BASE_URL,
        timeout: float = DEFAULT_TIMEOUT,
        session: requests.Session | None = None,
    ) -> None:
        self._base_url = base_url.rstrip("/")
        self._timeout = timeout
        self._session = session if session is not None else requests.Session()
        self._numeric_styles = frozenset(
            style.style_id for style in _DEFAULT_STYLES if style.is_numeric
        )

    @property
    def base_url(self) -> str:
        return self._base_url

    @property
    def timeout(self) -> float:
        return self._timeout

    def _request_json(
        self,
        path: str,
        params: dict[str, Any] | None = None,
        *,
        allow_404: bool = False,
    ) -> Any:
        url = f"{self._base_url}{path}"
        try:
            response = self._session.get(url, params=params, timeout=self._timeout)
        except requests.RequestException as exc:
            raise ZoteroOfflineError(OFFLINE_MESSAGE) from exc

        if allow_404 and response.status_code == 404:
            return None

        if response.status_code >= 400:
            raise ZoteroOfflineError(
                f"{OFFLINE_MESSAGE} (HTTP {response.status_code})."
            )

        try:
            return response.json()
        except ValueError as exc:
            raise ZoteroOfflineError(
                f"{OFFLINE_MESSAGE} (niepoprawna odpowiedź JSON)."
            ) from exc

    def _build_formatted_ref(
        self,
        raw: dict[str, Any],
        key: str,
        style_id: str,
    ) -> FormattedRef:
        bib_html = str(raw.get("bib") or "")
        citation_html = str(raw.get("citation") or "")

        bib_body_spans = parse_csl_bib_html(bib_html)
        in_text = parse_csl_citation_html(citation_html)
        bib_plain = spans_to_plain_text(bib_body_spans)

        sort_key = bib_plain
        if style_id in self._numeric_styles:
            parsed_item = _parse_zotero_item(raw)
            if parsed_item is not None:
                authors_part = ", ".join(parsed_item.authors)
                year_part = f" ({parsed_item.year})" if parsed_item.year else ""
                if authors_part:
                    sort_key = f"{authors_part}{year_part}"
                elif parsed_item.title:
                    sort_key = parsed_item.title

        return FormattedRef(
            key=key,
            style_id=style_id,
            in_text=in_text,
            bib_body_spans=bib_body_spans,
            sort_key=sort_key,
        )

    def is_available(self) -> bool:
        """Zwraca `True`, jeśli lokalne API Zotero odpowiada."""
        try:
            response = self._session.get(
                f"{self._base_url}/items/top",
                params={"limit": 1},
                timeout=self._timeout,
            )
            return response.status_code == 200
        except requests.RequestException:
            return False

    def search_items(self, query: str = "", limit: int = 25) -> list[ZoteroItem]:
        """Wyszukuje pozycje w bibliotece po autorze, tytule lub roku."""
        if limit <= 0:
            return []

        clean_query = query.strip()
        fetch_limit = max(limit, 50) if clean_query else limit
        params: dict[str, Any] = {
            "format": "json",
            "include": "data",
            "limit": fetch_limit,
        }
        if clean_query:
            params["q"] = clean_query
            params["qmode"] = "titleCreatorYear"

        payload = self._request_json("/items/top", params=params)
        if not isinstance(payload, list):
            raise ZoteroOfflineError(
                f"{OFFLINE_MESSAGE} (oczekiwano listy pozycji)."
            )

        items: list[ZoteroItem] = []
        for raw_item in payload:
            if not isinstance(raw_item, dict):
                continue
            item = _parse_zotero_item(raw_item)
            if item is None:
                continue
            items.append(item)
            if len(items) >= limit:
                break
        return items

    def get_formatted_ref(self, key: str, style_id: str) -> FormattedRef:
        """Zwraca sformatowane cytowanie i opis bibliograficzny dla klucza i stylu."""
        clean_key = key.strip()
        if not clean_key:
            raise KeyError("Pusty klucz pozycji Zotero.")

        payload = self._request_json(
            f"/items/{clean_key}",
            params={
                "include": "data,bib,citation",
                "style": style_id,
            },
            allow_404=True,
        )
        if payload is None or not isinstance(payload, dict):
            raise KeyError(f"Brak pozycji o kluczu {key!r} w Zotero.")

        return self._build_formatted_ref(payload, clean_key, style_id)

    def get_formatted_refs(
        self,
        keys: Sequence[str],
        style_id: str,
    ) -> dict[str, FormattedRef]:
        """Zwraca słownik `{key: FormattedRef}` dla podanych kluczy i stylu."""
        if not keys:
            return {}

        cleaned_keys: list[str] = []
        for key in keys:
            clean_key = key.strip()
            if not clean_key:
                raise KeyError("Pusty klucz pozycji Zotero.")
            cleaned_keys.append(clean_key)

        unique_keys = list(dict.fromkeys(cleaned_keys))

        if len(unique_keys) == 1:
            single_key = unique_keys[0]
            return {single_key: self.get_formatted_ref(single_key, style_id)}

        results: dict[str, FormattedRef] = {}
        for start in range(0, len(unique_keys), _BATCH_SIZE):
            chunk = unique_keys[start : start + _BATCH_SIZE]
            wanted = set(chunk)
            payload = self._request_json(
                "/items",
                params={
                    "itemKey": ",".join(chunk),
                    "include": "data,bib,citation",
                    "style": style_id,
                    "limit": max(len(chunk) * 4, 25),
                },
            )
            if isinstance(payload, list):
                for entry in payload:
                    if not isinstance(entry, dict):
                        continue
                    entry_key = str(entry.get("key") or "").strip()
                    if entry_key in wanted and entry_key not in results:
                        results[entry_key] = self._build_formatted_ref(
                            entry, entry_key, style_id
                        )

            for key in chunk:
                if key not in results:
                    results[key] = self.get_formatted_ref(key, style_id)

        return {key: results[key] for key in unique_keys}

    def list_styles(self) -> list[StyleInfo]:
        """Zwraca listę dostępnych stylów cytowania (IEEE v4 oraz APA 7)."""
        return list(_DEFAULT_STYLES)
