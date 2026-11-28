"""Testy jednostkowe i integracyjne modułu `azp.zotero` (Etap B1.1)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import pytest
import requests

from azp.contracts import (
    STYLE_APA,
    STYLE_IEEE_V4,
    FormattedRef,
    TextSpan,
    ZoteroCollection,
    ZoteroItem,
    ZoteroOfflineError,
    ZoteroSource,
    spans_to_plain_text,
)
from azp.zotero import (
    DEFAULT_BASE_URL,
    ZoteroClient,
    parse_csl_bib_html,
    parse_csl_citation_html,
)


@dataclass
class _FakeResponse:
    status_code: int
    payload: Any = None
    json_error: bool = False

    def json(self) -> Any:
        if self.json_error:
            raise ValueError("Invalid JSON")
        return self.payload


class _FakeSession:
    """Prosta atrapa `requests.Session` rejestrująca wywołania GET."""

    def __init__(self, handler: Any) -> None:
        self._handler = handler
        self.calls: list[tuple[str, dict[str, Any] | None, float | None]] = []

    def get(
        self,
        url: str,
        params: dict[str, Any] | None = None,
        timeout: float | None = None,
    ) -> _FakeResponse:
        self.calls.append((url, params, timeout))
        result = self._handler(url, params, timeout)
        if isinstance(result, Exception):
            raise result
        return result


# ---------------------------------------------------------------------------
# Testy parsera HTML CSL
# ---------------------------------------------------------------------------


def test_parse_csl_bib_html_ieee_left_margin_and_italics() -> None:
    html = (
        '<div class="csl-bib-body" style="line-height: 1.35; ">\n'
        '  <div class="csl-entry" style="clear: left; ">\n'
        '    <div class="csl-left-margin" style="float: left;">[1]</div>'
        '<div class="csl-right-inline" style="margin: 0 .4em 0 1.5em;">'
        'E. S. Dove, &ldquo;An ethics safe harbor,&rdquo; <i>Genome Med</i>, '
        'vol. 5, no. 11, p. 99, 2013.'
        "</div>\n"
        "  </div>\n"
        '  <span class="Z3988" title="ignored_coins"></span>\n'
        "</div>"
    )

    spans = parse_csl_bib_html(html)
    assert spans == (
        TextSpan('E. S. Dove, “An ethics safe harbor,” '),
        TextSpan("Genome Med", italic=True),
        TextSpan(", vol. 5, no. 11, p. 99, 2013."),
    )
    assert not spans_to_plain_text(spans).startswith("[1]")


def test_parse_csl_bib_html_strips_leading_numeric_prefix_without_left_margin() -> None:
    html = '<div class="csl-entry">[12]&nbsp;A. Kowalski, <em>Przegląd</em>, 2023.</div>'
    spans = parse_csl_bib_html(html)
    assert spans == (
        TextSpan("A. Kowalski, "),
        TextSpan("Przegląd", italic=True),
        TextSpan(", 2023."),
    )


def test_parse_csl_bib_html_strips_bold_numeric_prefix() -> None:
    html = '<div class="csl-entry"><b>[3]</b> B. Nowak, <strong>Tytuł</strong>, 2024.</div>'
    spans = parse_csl_bib_html(html)
    assert spans == (
        TextSpan("B. Nowak, "),
        TextSpan("Tytuł", bold=True),
        TextSpan(", 2024."),
    )


def test_parse_csl_bib_html_strips_span_numeric_prefix_before_italic() -> None:
    html = '<div class="csl-entry"><span>[4]</span> <i>Tytuł pracy</i>, 2025.</div>'
    spans = parse_csl_bib_html(html)
    assert spans == (
        TextSpan("Tytuł pracy", italic=True),
        TextSpan(", 2025."),
    )


def test_parse_csl_bib_html_apa_keeps_brackets_and_decodes_entities() -> None:
    html = (
        '<div class="csl-bib-body">\n'
        '  <div class="csl-entry">Dove, E. S., &amp; Zawati, M. H. (2013). '
        'Title &#38; subtitle [Special issue]. <i>Genome Medicine</i>, '
        "<i>5</i>(11), 99&ndash;105.</div>\n"
        "</div>"
    )
    spans = parse_csl_bib_html(html)
    assert spans == (
        TextSpan("Dove, E. S., & Zawati, M. H. (2013). Title & subtitle [Special issue]. "),
        TextSpan("Genome Medicine", italic=True),
        TextSpan(", "),
        TextSpan("5", italic=True),
        TextSpan("(11), 99–105."),
    )


def test_parse_csl_bib_html_apa_non_numeric_leading_bracket_preserved() -> None:
    html = '<div class="csl-entry">[Anonymous]. (2020). <i>Untitled work</i>.</div>'
    spans = parse_csl_bib_html(html)
    assert spans_to_plain_text(spans).startswith("[Anonymous]")


def test_parse_csl_bib_html_nested_styles_and_merging() -> None:
    html = (
        "<i>A</i><i>B</i><b></b>"
        "<i><b>C</b></i>"
        '<span style="font-style: italic; font-weight: 700;">D</span>'
        "  "
    )
    spans = parse_csl_bib_html(html)
    assert spans == (
        TextSpan("AB", italic=True, bold=False),
        TextSpan("CD", italic=True, bold=True),
    )


def test_parse_csl_citation_html_variants() -> None:
    assert parse_csl_citation_html("[1]") == "[1]"
    assert parse_csl_citation_html("<span>[1]</span>") == "[1]"
    assert parse_csl_citation_html("<span>(Dove et al., 2013)</span>") == "(Dove et al., 2013)"
    assert (
        parse_csl_citation_html("(Krackhardt &#38; Hanson, 1993)")
        == "(Krackhardt & Hanson, 1993)"
    )
    assert (
        parse_csl_citation_html("(<i>Complete Guide</i>,&nbsp;2026)")
        == "(Complete Guide, 2026)"
    )
    assert parse_csl_citation_html("   ") == ""


def test_parse_csl_bib_html_void_tags_do_not_leak_style_or_skip() -> None:
    assert parse_csl_bib_html("<i>Title<br>Subtitle</i>, 2024.") == (
        TextSpan("Title Subtitle", italic=True),
        TextSpan(", 2024.", italic=False),
    )
    assert parse_csl_bib_html("<i>Title<br/> Subtitle</i>, 2024.") == (
        TextSpan("Title Subtitle", italic=True),
        TextSpan(", 2024.", italic=False),
    )
    assert (
        parse_csl_bib_html(
            '<div class="csl-left-margin">[1]<br></div>'
            '<div class="csl-right-inline">Autor, 2024.</div>'
        )
        == (TextSpan("Autor, 2024."),)
    )


def test_parse_csl_bib_html_strips_numeric_prefix_without_space_across_tags() -> None:
    html = '<div class="csl-entry"><b>[1]</b><span>A. Kowalski, 2023.</span></div>'
    assert parse_csl_bib_html(html) == (TextSpan("A. Kowalski, 2023."),)


# ---------------------------------------------------------------------------
# Testy jednostkowe ZoteroClient (mock HTTP)
# ---------------------------------------------------------------------------


def test_zotero_client_implements_protocol_and_defaults() -> None:
    session = _FakeSession(lambda *_: _FakeResponse(200, []))
    client = ZoteroClient(session=session)
    assert isinstance(client, ZoteroSource)
    assert client.base_url == DEFAULT_BASE_URL

    custom = ZoteroClient(base_url="http://127.0.0.1:23119/api/users/0/", timeout=2.5, session=session)
    assert custom.base_url == DEFAULT_BASE_URL
    assert custom.timeout == 2.5


def test_is_available_returns_true_and_false_without_raising() -> None:
    ok_client = ZoteroClient(session=_FakeSession(lambda *_: _FakeResponse(200, [])))
    assert ok_client.is_available() is True

    http_err_client = ZoteroClient(session=_FakeSession(lambda *_: _FakeResponse(503, [])))
    assert http_err_client.is_available() is False

    conn_err_client = ZoteroClient(
        session=_FakeSession(lambda *_: requests.ConnectionError("refused"))
    )
    assert conn_err_client.is_available() is False

    timeout_client = ZoteroClient(
        session=_FakeSession(lambda *_: requests.Timeout("timed out"))
    )
    assert timeout_client.is_available() is False


def test_search_items_parses_creators_dates_and_filters_excluded_types() -> None:
    raw_items = [
        {
            "key": "ATT1",
            "data": {"key": "ATT1", "itemType": "attachment", "title": "PDF"},
        },
        {
            "key": "NOTE1",
            "data": {"key": "NOTE1", "itemType": "note", "title": "Notatka"},
        },
        {
            "key": "ANN1",
            "data": {"key": "ANN1", "itemType": "annotation", "title": "Adnotacja"},
        },
        {
            "key": "K1",
            "meta": {"parsedDate": "2023"},
            "data": {
                "key": "K1",
                "itemType": "journalArticle",
                "title": "Klasyfikacja AI",
                "date": "2023-05-12",
                "creators": [
                    {"firstName": "Adam", "lastName": "Kowalski", "creatorType": "author"},
                    {"firstName": "Jan", "lastName": "Redaktor", "creatorType": "editor"},
                ],
            },
        },
        {
            "key": "K2",
            "data": {
                "key": "K2",
                "itemType": "webpage",
                "title": "Raport UNESCO",
                "date": "Nov 2013",
                "creators": [
                    {"firstName": "", "lastName": "Unesco", "creatorType": "author"},
                    {"name": "World Health Organization", "creatorType": "author"},
                ],
            },
        },
        {
            "key": "K3",
            "meta": {"parsedDate": "2024"},
            "data": {
                "key": "K3",
                "itemType": "book",
                "title": "Książka pod redakcją",
                "date": "",
                "creators": [
                    {"firstName": "Barbara", "lastName": "Nowak", "creatorType": "editor"},
                ],
            },
        },
    ]

    session = _FakeSession(lambda *_: _FakeResponse(200, raw_items))
    client = ZoteroClient(session=session)

    results = client.search_items(query="Kowalski 2023", limit=10)
    assert len(results) == 3
    assert results[0] == ZoteroItem(
        key="K1",
        title="Klasyfikacja AI",
        authors=("Kowalski, Adam",),
        year="2023",
        item_type="journalArticle",
    )
    assert results[1] == ZoteroItem(
        key="K2",
        title="Raport UNESCO",
        authors=("Unesco", "World Health Organization"),
        year="2013",
        item_type="webpage",
    )
    assert results[2] == ZoteroItem(
        key="K3",
        title="Książka pod redakcją",
        authors=("Nowak, Barbara",),
        year="2024",
        item_type="book",
    )

    url, params, _ = session.calls[0]
    assert url.endswith("/items/top")
    assert params is not None
    assert params["q"] == "Kowalski 2023"
    assert params["qmode"] == "titleCreatorYear"
    assert params["format"] == "json"
    assert params["include"] == "data"


def test_search_items_respects_limit_and_empty_query() -> None:
    raw_items = [
        {"key": f"K{i}", "data": {"key": f"K{i}", "itemType": "journalArticle", "title": f"T{i}", "date": "2024"}}
        for i in range(5)
    ]
    session = _FakeSession(lambda *_: _FakeResponse(200, raw_items))
    client = ZoteroClient(session=session)

    assert client.search_items(limit=0) == []
    assert len(session.calls) == 0

    limited = client.search_items(query="", limit=2)
    assert [item.key for item in limited] == ["K0", "K1"]
    _, params, _ = session.calls[0]
    assert params is not None
    assert "q" not in params
    assert params["limit"] == 2


def test_get_formatted_ref_ieee_v4_and_apa() -> None:
    def handler(url: str, params: dict[str, Any] | None, _: float | None) -> _FakeResponse:
        assert params is not None
        style = params["style"]
        if url.endswith("/items/DOVE2013") and style == STYLE_IEEE_V4:
            return _FakeResponse(
                200,
                {
                    "key": "DOVE2013",
                    "data": {
                        "key": "DOVE2013",
                        "itemType": "journalArticle",
                        "title": "Ethics safe harbor",
                        "date": "2013",
                        "creators": [
                            {"firstName": "Edward S.", "lastName": "Dove", "creatorType": "author"}
                        ],
                    },
                    "citation": "[1]",
                    "bib": (
                        '<div class="csl-bib-body"><div class="csl-entry">'
                        '<div class="csl-left-margin">[1]</div>'
                        '<div class="csl-right-inline">E. S. Dove, "Ethics safe harbor," '
                        "<i>Genome Med</i>, 2013.</div>"
                        "</div></div>"
                    ),
                },
            )
        if url.endswith("/items/DOVE2013") and style == STYLE_APA:
            return _FakeResponse(
                200,
                {
                    "key": "DOVE2013",
                    "data": {
                        "key": "DOVE2013",
                        "itemType": "journalArticle",
                        "title": "Ethics safe harbor",
                        "date": "2013",
                        "creators": [
                            {"firstName": "Edward S.", "lastName": "Dove", "creatorType": "author"}
                        ],
                    },
                    "citation": "(Dove &#38; Zawati, 2013)",
                    "bib": (
                        '<div class="csl-bib-body"><div class="csl-entry">'
                        "Dove, E. S., &amp; Zawati, M. H. (2013). Ethics safe harbor. "
                        "<i>Genome Medicine</i>."
                        "</div></div>"
                    ),
                },
            )
        return _FakeResponse(404, "Not found")

    client = ZoteroClient(session=_FakeSession(handler))

    ref_v4 = client.get_formatted_ref("DOVE2013", STYLE_IEEE_V4)
    assert isinstance(ref_v4, FormattedRef)
    assert ref_v4.key == "DOVE2013"
    assert ref_v4.style_id == STYLE_IEEE_V4
    assert ref_v4.in_text == "[1]"
    assert ref_v4.bib_body_spans == (
        TextSpan('E. S. Dove, "Ethics safe harbor," '),
        TextSpan("Genome Med", italic=True),
        TextSpan(", 2013."),
    )
    assert ref_v4.sort_key == "Dove, Edward S. (2013)"

    ref_apa = client.get_formatted_ref("DOVE2013", STYLE_APA)
    assert ref_apa.in_text == "(Dove & Zawati, 2013)"
    assert ref_apa.bib_body_spans == (
        TextSpan("Dove, E. S., & Zawati, M. H. (2013). Ethics safe harbor. "),
        TextSpan("Genome Medicine", italic=True),
        TextSpan("."),
    )


def test_get_formatted_ref_missing_key_raises_key_error() -> None:
    client = ZoteroClient(session=_FakeSession(lambda *_: _FakeResponse(404, "Not found")))
    with pytest.raises(KeyError):
        client.get_formatted_ref("MISSING", STYLE_IEEE_V4)

    with pytest.raises(KeyError):
        client.get_formatted_ref("   ", STYLE_IEEE_V4)


def test_get_formatted_refs_batch_dedup_and_missing() -> None:
    def handler(url: str, params: dict[str, Any] | None, _: float | None) -> _FakeResponse:
        if url.endswith("/items"):
            assert params is not None
            return _FakeResponse(
                200,
                [
                    {
                        "key": "ATT_CHILD",
                        "data": {"key": "ATT_CHILD", "itemType": "attachment"},
                        "citation": "",
                        "bib": "",
                    },
                    {
                        "key": "K1",
                        "citation": "[1]",
                        "bib": '<div class="csl-entry">[1] Autor A, <i>Czasopismo</i>, 2023.</div>',
                    },
                    {
                        "key": "K2",
                        "citation": "[1]",
                        "bib": '<div class="csl-entry">[1] Autor B, 2024.</div>',
                    },
                ],
            )
        if url.endswith("/items/MISSING"):
            return _FakeResponse(404, "Not found")
        return _FakeResponse(404, "Not found")

    session = _FakeSession(handler)
    client = ZoteroClient(session=session)

    assert client.get_formatted_refs([], STYLE_IEEE_V4) == {}
    assert len(session.calls) == 0

    refs = client.get_formatted_refs(["K1", "K2", "K1"], STYLE_IEEE_V4)
    assert list(refs.keys()) == ["K1", "K2"]
    assert refs["K1"].bib_plain == "Autor A, Czasopismo, 2023."
    assert refs["K2"].bib_plain == "Autor B, 2024."

    with pytest.raises(KeyError):
        client.get_formatted_refs(["K1", "MISSING"], STYLE_IEEE_V4)


def test_get_formatted_refs_normalizes_and_deduplicates_whitespace_keys() -> None:
    def handler(url: str, params: dict[str, Any] | None, _: float | None) -> _FakeResponse:
        if url.endswith("/items"):
            assert params is not None
            assert params["itemKey"] == "K1,K2"
            return _FakeResponse(
                200,
                [
                    {
                        "key": "K1",
                        "citation": "[1]",
                        "bib": '<div class="csl-entry">[1] Autor A, 2023.</div>',
                    },
                    {
                        "key": "K2",
                        "citation": "[1]",
                        "bib": '<div class="csl-entry">[1] Autor B, 2024.</div>',
                    },
                ],
            )
        return _FakeResponse(404, "Not found")

    session = _FakeSession(handler)
    client = ZoteroClient(session=session)

    refs = client.get_formatted_refs([" K1 ", "K2", "K1"], STYLE_IEEE_V4)
    assert list(refs.keys()) == ["K1", "K2"]
    assert len(session.calls) == 1


@pytest.mark.parametrize(
    "exc_or_response",
    [
        requests.ConnectionError("Connection refused"),
        requests.Timeout("Read timed out"),
        _FakeResponse(500, "Internal Server Error"),
        _FakeResponse(403, "Forbidden"),
        _FakeResponse(200, json_error=True),
    ],
)
def test_offline_errors_raise_zotero_offline_error_with_polish_message(
    exc_or_response: Any,
) -> None:
    client = ZoteroClient(session=_FakeSession(lambda *_: exc_or_response))

    with pytest.raises(ZoteroOfflineError, match="Zotero nie odpowiada"):
        client.search_items("test")

    with pytest.raises(ZoteroOfflineError, match="Zotero nie odpowiada"):
        client.get_formatted_ref("K1", STYLE_IEEE_V4)

    with pytest.raises(ZoteroOfflineError, match="Zotero nie odpowiada"):
        client.get_formatted_refs(["K1", "K2"], STYLE_IEEE_V4)


def test_list_styles_contains_ieee_v4_and_apa() -> None:
    client = ZoteroClient(session=_FakeSession(lambda *_: _FakeResponse(200, [])))
    styles = client.list_styles()
    by_id = {s.style_id: s for s in styles}
    assert STYLE_IEEE_V4 in by_id
    assert by_id[STYLE_IEEE_V4].is_numeric is True
    assert STYLE_APA in by_id
    assert by_id[STYLE_APA].is_numeric is False


def test_search_items_multi_word_cross_field_fallback_and_empty_query_topup() -> None:
    """Gdy Zotero dla zapytania wielowyrazowego zwraca [], klient dociąga po głównym tokenie i filtruje."""
    raw_items = [
        {
            "key": "K1",
            "data": {
                "key": "K1",
                "itemType": "journalArticle",
                "title": "An ethics safe harbor for international genomics",
                "date": "2013",
                "creators": [
                    {"firstName": "Edward S.", "lastName": "Dove", "creatorType": "author"}
                ],
            },
        },
        {
            "key": "K2",
            "data": {
                "key": "K2",
                "itemType": "journalArticle",
                "title": "Other genomics paper",
                "date": "2020",
                "creators": [
                    {"firstName": "Jan", "lastName": "Nowak", "creatorType": "author"}
                ],
            },
        },
    ]

    def handler(url: str, params: dict[str, Any] | None, _: float | None) -> _FakeResponse:
        assert params is not None
        q = params.get("q", "")
        if q == "Dove genomics 2013":
            return _FakeResponse(200, [])
        if q == "genomics":
            return _FakeResponse(200, raw_items)
        return _FakeResponse(200, [])

    client = ZoteroClient(session=_FakeSession(handler))
    results = client.search_items("Dove genomics 2013", limit=10)
    assert [it.key for it in results] == ["K1"]

    # Puste zapytanie: pierwszy batch ma załącznik, drugi dociąga pełny limit
    def empty_q_handler(url: str, params: dict[str, Any] | None, _: float | None) -> _FakeResponse:
        assert params is not None
        if params.get("limit") == 2:
            return _FakeResponse(
                200,
                [
                    {"key": "ATT1", "data": {"key": "ATT1", "itemType": "attachment"}},
                    raw_items[0],
                ],
            )
        return _FakeResponse(200, raw_items)

    client_topup = ZoteroClient(session=_FakeSession(empty_q_handler))
    topup_results = client_topup.search_items("", limit=2)
    assert [it.key for it in topup_results] == ["K1", "K2"]


def test_get_formatted_ref_falls_back_to_standard_ieee_when_custom_v4_missing() -> None:
    """Gdy niestandardowy styl IEEE v4 nie jest zainstalowany w Zotero (HTTP 400), klient używa standardowego IEEE."""
    from azp.contracts import STYLE_IEEE_FALLBACK

    def handler(url: str, params: dict[str, Any] | None, _: float | None) -> _FakeResponse:
        assert params is not None
        if params.get("style") == STYLE_IEEE_V4:
            return _FakeResponse(400, "Invalid style")
        assert params.get("style") == STYLE_IEEE_FALLBACK
        if url.endswith("/items/K1"):
            return _FakeResponse(
                200,
                {
                    "key": "K1",
                    "citation": "[1]",
                    "bib": '<div class="csl-entry">[1] A. Kowalski, <i>Test</i>, 2023.</div>',
                },
            )
        return _FakeResponse(
            200,
            [
                {
                    "key": "K1",
                    "citation": "[1]",
                    "bib": '<div class="csl-entry">[1] A. Kowalski, <i>Test</i>, 2023.</div>',
                },
                {
                    "key": "K2",
                    "citation": "[1]",
                    "bib": '<div class="csl-entry">[1] B. Nowak, 2024.</div>',
                },
            ],
        )

    client = ZoteroClient(session=_FakeSession(handler))
    single = client.get_formatted_ref("K1", STYLE_IEEE_V4)
    assert single.in_text == "[1]"
    assert single.bib_plain == "A. Kowalski, Test, 2023."

    batch = client.get_formatted_refs(["K1", "K2"], STYLE_IEEE_V4)
    assert set(batch.keys()) == {"K1", "K2"}


def test_list_collections_parses_hierarchy_sorts_and_handles_cycles_d25() -> None:
    raw_collections = [
        {
            "key": "COL_ML",
            "data": {
                "key": "COL_ML",
                "name": "Uczenie maszynowe",
                "parentCollection": "COL_AI",
            },
        },
        {
            "key": "COL_AI",
            "data": {
                "key": "COL_AI",
                "name": "Sztuczna inteligencja",
                "parentCollection": False,
            },
        },
        {
            "key": "COL_ETH",
            "data": {
                "key": "COL_ETH",
                "name": "Etyka i prawo",
                "parentCollection": "false",
            },
        },
        {
            "key": "COL_DEEP",
            "data": {
                "key": "COL_DEEP",
                "name": "Sieci głębokie",
                "parentCollection": "COL_ML",
            },
        },
        {
            "key": "CYC_A",
            "data": {
                "key": "CYC_A",
                "name": "Cykl A",
                "parentCollection": "CYC_B",
            },
        },
        {
            "key": "CYC_B",
            "data": {
                "key": "CYC_B",
                "name": "Cykl B",
                "parentCollection": "CYC_A",
            },
        },
    ]

    session = _FakeSession(lambda *_: _FakeResponse(200, raw_collections))
    client = ZoteroClient(session=session)
    collections = client.list_collections()

    url, params, _ = session.calls[0]
    assert url.endswith("/collections")
    assert params == {"format": "json", "include": "data", "limit": 200}

    by_key = {c.key: c for c in collections}
    assert by_key["COL_AI"] == ZoteroCollection(
        key="COL_AI",
        name="Sztuczna inteligencja",
        parent_key="",
    )
    assert by_key["COL_ETH"] == ZoteroCollection(
        key="COL_ETH",
        name="Etyka i prawo",
        parent_key="",
    )
    assert by_key["COL_ML"] == ZoteroCollection(
        key="COL_ML",
        name="Sztuczna inteligencja / Uczenie maszynowe",
        parent_key="COL_AI",
    )
    assert by_key["COL_DEEP"] == ZoteroCollection(
        key="COL_DEEP",
        name="Sztuczna inteligencja / Uczenie maszynowe / Sieci głębokie",
        parent_key="COL_ML",
    )
    assert "CYC_A" in by_key and "CYC_B" in by_key
    assert [c.name for c in collections] == sorted(
        (c.name for c in collections), key=str.casefold
    )


def test_search_items_with_collection_key_uses_collection_endpoint_d25() -> None:
    raw_items = [
        {
            "key": "K1",
            "data": {
                "key": "K1",
                "itemType": "journalArticle",
                "title": "Deep Learning in Practice",
                "date": "2024",
                "creators": [
                    {"firstName": "Barbara", "lastName": "Nowak", "creatorType": "author"}
                ],
            },
        }
    ]

    def handler(url: str, params: dict[str, Any] | None, _: float | None) -> _FakeResponse:
        assert params is not None
        if params.get("q") == "Nowak Learning":
            return _FakeResponse(200, [])
        return _FakeResponse(200, raw_items)

    session = _FakeSession(handler)
    client = ZoteroClient(session=session)

    # 1. Z podanym collection_key (również w fallbacku wielowyrazowym)
    results = client.search_items("Nowak Learning", limit=10, collection_key=" COL_ML ")
    assert [it.key for it in results] == ["K1"]
    assert len(session.calls) == 2
    assert all(
        call[0].endswith("/collections/COL_ML/items/top") for call in session.calls
    )

    # 2. Gdy collection_key jest None lub pusty -> /items/top
    session.calls.clear()
    client.search_items("", limit=5, collection_key="   ")
    assert len(session.calls) == 1
    assert session.calls[0][0].endswith("/items/top")
    assert "/collections/" not in session.calls[0][0]


def test_list_collections_resilience_malformed_json_and_self_cycle_m3() -> None:
    # 1. Payload niebędący listą (np. słownik) -> ZoteroOfflineError
    bad_payload_client = ZoteroClient(
        session=_FakeSession(lambda *_: _FakeResponse(200, {"error": "not a list"}))
    )
    with pytest.raises(ZoteroOfflineError, match="oczekiwano listy kolekcji"):
        bad_payload_client.list_collections()

    # 2. Błąd HTTP -> ZoteroOfflineError
    http_err_client = ZoteroClient(
        session=_FakeSession(lambda *_: _FakeResponse(502, "Bad Gateway"))
    )
    with pytest.raises(ZoteroOfflineError, match="HTTP 502"):
        http_err_client.list_collections()

    # 3. Elementy niebędące słownikami, rekordy bez key/name, płaskie obiekty i cykl key == parentCollection
    mixed_payload: list[Any] = [
        "invalid_entry",
        123,
        {"key": "", "data": {"name": "Bez klucza"}},
        {"key": "NO_NAME", "data": {"name": "   "}},
        {
            "key": "SELF_CYC",
            "data": {
                "key": "SELF_CYC",
                "name": "Samorodzic",
                "parentCollection": "SELF_CYC",
            },
        },
        {
            "key": "FLAT_COL",
            "name": "Płaska kolekcja",
            "parentCollection": None,
        },
    ]
    mixed_client = ZoteroClient(
        session=_FakeSession(lambda *_: _FakeResponse(200, mixed_payload))
    )
    cols = mixed_client.list_collections()
    assert cols == [
        ZoteroCollection(key="FLAT_COL", name="Płaska kolekcja", parent_key=""),
        ZoteroCollection(key="SELF_CYC", name="Samorodzic", parent_key="SELF_CYC"),
    ]


def test_formatted_refs_caching_and_clear_cache() -> None:
    """Testuje, że get_formatted_ref/get_formatted_refs buforuje wyniki w pamięci i clear_cache czyści bufor."""
    requests_log: list[str] = []

    def _handler(url: str, params: dict[str, Any] | None, _timeout: float | None) -> _FakeResponse:
        requests_log.append(url)
        if "/items/KOWAL2023" in url:
            return _FakeResponse(
                200,
                {
                    "csl-json": {"author": [{"family": "Kowalski"}]},
                    "csl-citation": "[1]",
                    "csl-bib": '<div class="csl-bib-body"><div class="csl-entry">[1] J. Kowalski, 2023.</div></div>',
                },
            )
        if url.endswith("/items") and params and "NOWAK2024" in params.get("itemKey", ""):
            return _FakeResponse(
                200,
                [
                    {
                        "key": "NOWAK2024",
                        "citation": "[2]",
                        "bib": '<div class="csl-entry">[2] A. Nowak, 2024.</div>',
                    }
                ],
            )
        return _FakeResponse(404, {})

    client = ZoteroClient(session=_FakeSession(_handler))

    # Pierwsze pobranie: zapytanie sieciowe
    ref1 = client.get_formatted_ref("KOWAL2023", STYLE_IEEE_V4)
    assert ref1.key == "KOWAL2023"
    assert len(requests_log) == 1

    # Drugie pobranie: z pamięci podręcznej (brak nowego zapytania)
    ref1_cached = client.get_formatted_ref("KOWAL2023", STYLE_IEEE_V4)
    assert ref1_cached == ref1
    assert len(requests_log) == 1

    # get_formatted_refs dla KOWAL2023 i NOWAK2024: tylko NOWAK2024 powinien wykonać zapytanie
    batch = client.get_formatted_refs(["KOWAL2023", "NOWAK2024"], STYLE_IEEE_V4)
    assert set(batch.keys()) == {"KOWAL2023", "NOWAK2024"}
    assert len(requests_log) == 2

    # clear_cache() czyści bufor
    client.clear_cache()
    ref1_after_clear = client.get_formatted_ref("KOWAL2023", STYLE_IEEE_V4)
    assert ref1_after_clear.key == "KOWAL2023"
    assert len(requests_log) == 3


# ---------------------------------------------------------------------------
# Test integracyjny z żywym lokalnym API Zotero 10
# ---------------------------------------------------------------------------


def test_live_zotero_api_search_and_formatting() -> None:
    client = ZoteroClient()
    if not client.is_available():
        pytest.skip("Lokalne API Zotero 10 nie działa pod 127.0.0.1:23119.")

    items = client.search_items(limit=5)
    assert len(items) >= 1
    first = items[0]
    assert first.key
    assert first.item_type not in {"attachment", "note", "annotation"}

    # Wyszukanie po fragmencie tytułu, autorze i roku pierwszego elementu
    query_word = next((w for w in first.title.split() if len(w) >= 4 and w.isalnum()), "")
    if query_word:
        found = client.search_items(query=query_word, limit=10)
        assert any(item.key == first.key for item in found)

    if first.authors:
        surname = first.authors[0].split(",")[0].strip()
        if len(surname) >= 3:
            found_author = client.search_items(query=surname, limit=10)
            assert any(item.key == first.key for item in found_author)

    if first.year:
        found_year = client.search_items(query=first.year, limit=25)
        assert any(item.key == first.key for item in found_year)

    ref_v4 = client.get_formatted_ref(first.key, STYLE_IEEE_V4)
    assert ref_v4.key == first.key
    assert ref_v4.in_text == "[1]"
    assert ref_v4.bib_plain
    assert not ref_v4.bib_plain.startswith("[1]")
    assert any(span.italic for span in ref_v4.bib_body_spans)

    ref_apa = client.get_formatted_ref(first.key, STYLE_APA)
    assert ref_apa.key == first.key
    assert ref_apa.in_text.startswith("(") and ref_apa.in_text.endswith(")")
    assert ref_apa.bib_plain

    if len(items) >= 2:
        batch = client.get_formatted_refs([items[0].key, items[1].key], STYLE_IEEE_V4)
        assert set(batch.keys()) == {items[0].key, items[1].key}

    with pytest.raises(KeyError):
        client.get_formatted_ref("NONEXISTENT_KEY_AZP_TEST", STYLE_IEEE_V4)
