"""Testy jednostkowe i integracyjne silnika numeracji i stylów (Etap B2.2)."""

from __future__ import annotations

from dataclasses import replace
import pytest

from azp.contracts import (
    STYLE_APA,
    STYLE_IEEE_V4,
    AzpError,
    Citation,
    FormattedRef,
    NoActiveSlideError,
    NoTextCursorError,
    Settings,
    SlideRecord,
    StyleInfo,
    TextSpan,
    ZoteroOfflineError,
)
from azp.engine import (
    CitationEngine,
    assign_global_numbers,
    assign_slide_numbers,
    build_slide_footer_entries,
    compute_presentation_plan,
    format_citation_text,
    is_numeric_style,
    merge_citation_keys,
)
from tests.fakes.fake_ppt import FakeSlideDoc
from tests.fakes.fake_zotero import FakeZoteroSource


class MergingFakeSlideDoc(FakeSlideDoc):
    """Podklasa `FakeSlideDoc` symulująca automatyczne scalanie sąsiedniego cytowania przy kursorze (D13)."""

    def __init__(self, *args: object, merge_adjacent: bool = True, **kwargs: object) -> None:
        super().__init__(*args, **kwargs)  # type: ignore[arg-type]
        self.merge_adjacent = merge_adjacent

    def insert_citation_at_cursor(self, citation: Citation) -> None:
        slide_idx = self.get_active_slide_index()
        if not self.has_cursor:
            raise NoTextCursorError(
                "Ustaw kursor w polu tekstowym na slajdzie przed wstawieniem cytowania."
            )
        slide = self._get_slide(slide_idx)
        if self.merge_adjacent and slide.citations:
            prev = slide.citations[-1]
            merged_keys = merge_citation_keys(prev.keys, citation.keys)
            temp_text = f"{prev.text}{citation.text}"
            if prev.text and prev.text in slide.body_text:
                pos = slide.body_text.rfind(prev.text)
                slide.body_text = (
                    slide.body_text[:pos]
                    + temp_text
                    + slide.body_text[pos + len(prev.text) :]
                )
            slide.citations[-1] = Citation(
                citation_id=prev.citation_id,
                keys=merged_keys,
                text=temp_text,
            )
            return
        super().insert_citation_at_cursor(citation)


# ---------------------------------------------------------------------------
# 1. Testy czystych funkcji numeracji i formatowania
# ---------------------------------------------------------------------------


def test_is_numeric_style_recognizes_builtin_and_custom_styles() -> None:
    assert is_numeric_style(STYLE_IEEE_V4) is True
    assert is_numeric_style(STYLE_APA) is False

    custom_styles = [
        StyleInfo(style_id="http://www.zotero.org/styles/custom-num", title="Custom Num", is_numeric=True),
        StyleInfo(style_id="http://www.zotero.org/styles/custom-auth", title="Custom Auth", is_numeric=False),
    ]
    assert is_numeric_style("http://www.zotero.org/styles/custom-num", custom_styles) is True
    assert is_numeric_style("http://www.zotero.org/styles/custom-auth", custom_styles) is False


def test_assign_global_and_slide_numbers_first_appearance_and_repeats() -> None:
    slide1 = SlideRecord(
        slide_index=1,
        slide_id=101,
        citations=(
            Citation(citation_id="c1", keys=("KOWAL2023",), text="[1]"),
            Citation(citation_id="c2", keys=("NOWAK2024", "KOWAL2023"), text="[1, 2]"),
        ),
    )
    slide2 = SlideRecord(
        slide_index=2,
        slide_id=102,
        citations=(
            Citation(citation_id="c3", keys=("NOWAK2024",), text="[2]"),
            Citation(citation_id="c4", keys=("DOVE2013", "ZIEL2022"), text="[3, 4]"),
            Citation(citation_id="c5", keys=("KOWAL2023",), text="[1]"),
        ),
    )

    global_map = assign_global_numbers([slide1, slide2])
    assert global_map == {
        "KOWAL2023": 1,
        "NOWAK2024": 2,
        "DOVE2013": 3,
        "ZIEL2022": 4,
    }

    # Tryb per_slide (D4): na slajdzie 2 numeracja rusza od [1]
    slide2_map = assign_slide_numbers(slide2.citations)
    assert slide2_map == {
        "NOWAK2024": 1,
        "DOVE2013": 2,
        "ZIEL2022": 3,
        "KOWAL2023": 4,
    }


def test_merge_citation_keys_preserves_order_and_deduplicates() -> None:
    merged = merge_citation_keys(("KOWAL2023", "NOWAK2024"), ("NOWAK2024", "DOVE2013", "KOWAL2023"))
    assert merged == ("KOWAL2023", "NOWAK2024", "DOVE2013")

    with pytest.raises(ValueError):
        merge_citation_keys((), ())


def test_format_citation_text_ieee_and_apa() -> None:
    zotero = FakeZoteroSource()
    all_keys = ("KOWAL2023", "NOWAK2024", "DOVE2013", "ZIEL2022")
    ieee_refs = zotero.get_formatted_refs(all_keys, STYLE_IEEE_V4)
    apa_refs = zotero.get_formatted_refs(all_keys, STYLE_APA)
    num_map = {"KOWAL2023": 1, "NOWAK2024": 2, "DOVE2013": 3, "ZIEL2022": 4}

    # IEEE pojedyncze
    assert format_citation_text(("KOWAL2023",), ieee_refs, is_numeric=True, number_map=num_map) == "[1]"

    # IEEE zbiorcze/scalone: sortowane rosnąco po numerach i bez duplikatów (D13)
    assert (
        format_citation_text(
            ("ZIEL2022", "KOWAL2023", "DOVE2013", "KOWAL2023"),
            ieee_refs,
            is_numeric=True,
            number_map=num_map,
        )
        == "[1, 3, 4]"
    )

    # APA pojedyncze i zbiorcze (D3, D13)
    assert (
        format_citation_text(("KOWAL2023",), apa_refs, is_numeric=False)
        == "(Kowalski, 2023)"
    )
    assert (
        format_citation_text(
            ("KOWAL2023", "NOWAK2024", "KOWAL2023"),
            apa_refs,
            is_numeric=False,
        )
        == "(Kowalski, 2023; Nowak & Wiśniewski, 2024)"
    )

    with pytest.raises(ValueError):
        format_citation_text((), ieee_refs, is_numeric=True, number_map=num_map)

    with pytest.raises(ValueError):
        format_citation_text(("KOWAL2023",), ieee_refs, is_numeric=True, number_map=None)


def test_build_slide_footer_entries_ieee_numeric_order_and_italics() -> None:
    zotero = FakeZoteroSource()
    all_keys = ("KOWAL2023", "NOWAK2024", "DOVE2013")
    ieee_refs = zotero.get_formatted_refs(all_keys, STYLE_IEEE_V4)
    num_map = {"KOWAL2023": 1, "NOWAK2024": 2, "DOVE2013": 3}

    # Na slajdzie najpierw zacytowano pozycję [3], a potem [1, 2]
    citations = (
        Citation(citation_id="c1", keys=("DOVE2013",), text="[3]"),
        Citation(citation_id="c2", keys=("KOWAL2023", "NOWAK2024"), text="[1, 2]"),
    )
    entries = build_slide_footer_entries(citations, ieee_refs, is_numeric=True, number_map=num_map)

    assert [e.key for e in entries] == ["KOWAL2023", "NOWAK2024", "DOVE2013"]
    assert [e.prefix for e in entries] == ["[1] ", "[2] ", "[3] "]
    assert entries[0].plain_text.startswith('[1] A. Kowalski, "Klasyfikacja')
    # T4, T8: zachowana kursywa na tytule czasopisma/książki w spans
    assert any(span.italic and "Przegląd Informatyczny" in span.text for span in entries[0].spans)
    assert any(span.italic and "Algorytmy uczenia maszynowego" in span.text for span in entries[1].spans)


def test_build_slide_footer_entries_apa_alphabetical_order() -> None:
    zotero = FakeZoteroSource()
    all_keys = ("ZIEL2022", "NOWAK2024", "KOWAL2023", "DOVE2013")
    apa_refs = zotero.get_formatted_refs(all_keys, STYLE_APA)

    # Cytowane w odwrotnej kolejności alfabetycznej na slajdzie
    citations = (
        Citation(citation_id="c1", keys=("ZIEL2022", "NOWAK2024"), text="(Zieliński, 2022; Nowak & Wiśniewski, 2024)"),
        Citation(citation_id="c2", keys=("KOWAL2023", "DOVE2013"), text="(Kowalski, 2023; Dove et al., 2013)"),
    )
    entries = build_slide_footer_entries(citations, apa_refs, is_numeric=False)

    # W stopce APA muszą być posortowane alfabetycznie bez prefiksu [n] (D3, T8)
    assert [e.key for e in entries] == ["DOVE2013", "KOWAL2023", "NOWAK2024", "ZIEL2022"]
    assert all(e.prefix == "" for e in entries)
    assert entries[0].plain_text.startswith("Dove, E. S.")
    assert entries[1].plain_text.startswith("Kowalski, A.")
    assert entries[2].plain_text.startswith("Nowak, B.")
    assert entries[3].plain_text.startswith("Zieliński, M.")


def test_build_slide_footer_entries_empty_and_tie_breaker() -> None:
    assert build_slide_footer_entries((), {}, is_numeric=False) == ()

    ref_b = FormattedRef(
        key="B_KEY",
        style_id=STYLE_APA,
        in_text="(Anon, 2020b)",
        bib_body_spans=(TextSpan("Anon. (2020). Drugi artykuł."),),
        sort_key="Anon (2020)",
    )
    ref_a = FormattedRef(
        key="A_KEY",
        style_id=STYLE_APA,
        in_text="(Anon, 2020a)",
        bib_body_spans=(TextSpan("Anon. (2020). Pierwszy artykuł."),),
        sort_key="Anon (2020)",
    )
    citations = (Citation(citation_id="c1", keys=("A_KEY", "B_KEY")),)
    entries = build_slide_footer_entries(
        citations,
        {"A_KEY": ref_a, "B_KEY": ref_b},
        is_numeric=False,
    )
    # "Drugi artykuł" < "Pierwszy artykuł" alfabetycznie w tie-breakerze po bib_plain
    assert [e.key for e in entries] == ["B_KEY", "A_KEY"]


def test_compute_presentation_plan_global_per_slide_and_apa() -> None:
    zotero = FakeZoteroSource()
    all_keys = ("KOWAL2023", "NOWAK2024", "DOVE2013")
    slides = [
        SlideRecord(
            slide_index=1,
            slide_id=101,
            citations=(
                Citation(citation_id="c1", keys=("NOWAK2024",), text=""),
                Citation(citation_id="c2", keys=("KOWAL2023", "NOWAK2024"), text=""),
            ),
        ),
        SlideRecord(
            slide_index=2,
            slide_id=102,
            citations=(
                Citation(citation_id="c3", keys=("DOVE2013",), text=""),
                Citation(citation_id="c4", keys=("KOWAL2023",), text=""),
            ),
        ),
    ]

    # 1. Global IEEE
    ieee_refs = zotero.get_formatted_refs(all_keys, STYLE_IEEE_V4)
    plan_global = compute_presentation_plan(
        slides,
        ieee_refs,
        Settings(style_id=STYLE_IEEE_V4, numbering_mode="global"),
    )
    assert [c.text for c in plan_global[0].citations] == ["[1]", "[1, 2]"]
    assert [e.prefix for e in plan_global[0].footer_entries] == ["[1] ", "[2] "]
    assert [e.key for e in plan_global[0].footer_entries] == ["NOWAK2024", "KOWAL2023"]
    assert [c.text for c in plan_global[1].citations] == ["[3]", "[2]"]
    assert [e.prefix for e in plan_global[1].footer_entries] == ["[2] ", "[3] "]
    assert [e.key for e in plan_global[1].footer_entries] == ["KOWAL2023", "DOVE2013"]

    # 2. Per-slide IEEE
    plan_per_slide = compute_presentation_plan(
        slides,
        ieee_refs,
        Settings(style_id=STYLE_IEEE_V4, numbering_mode="per_slide"),
    )
    assert [c.text for c in plan_per_slide[0].citations] == ["[1]", "[1, 2]"]
    assert [c.text for c in plan_per_slide[1].citations] == ["[1]", "[2]"]
    assert [e.prefix for e in plan_per_slide[1].footer_entries] == ["[1] ", "[2] "]
    assert [e.key for e in plan_per_slide[1].footer_entries] == ["DOVE2013", "KOWAL2023"]

    # 3. APA 7
    apa_refs = zotero.get_formatted_refs(all_keys, STYLE_APA)
    plan_apa = compute_presentation_plan(
        slides,
        apa_refs,
        Settings(style_id=STYLE_APA, numbering_mode="global"),
    )
    assert [c.text for c in plan_apa[0].citations] == [
        "(Nowak & Wiśniewski, 2024)",
        "(Kowalski, 2023; Nowak & Wiśniewski, 2024)",
    ]
    assert [e.key for e in plan_apa[0].footer_entries] == ["KOWAL2023", "NOWAK2024"]
    assert [c.text for c in plan_apa[1].citations] == [
        "(Dove et al., 2013)",
        "(Kowalski, 2023)",
    ]
    assert [e.key for e in plan_apa[1].footer_entries] == ["DOVE2013", "KOWAL2023"]


# ---------------------------------------------------------------------------
# 2. Testy CitationEngine z FakeZoteroSource i FakeSlideDoc
# ---------------------------------------------------------------------------


def test_engine_cite_without_cursor_makes_zero_changes_d10() -> None:
    zotero = FakeZoteroSource()
    doc = FakeSlideDoc(slide_count=2, active_slide_index=1, has_cursor=False)
    engine = CitationEngine(zotero, doc)

    before_slides = doc.read_slides()
    before_body = doc.get_slide_body_text(1)

    with pytest.raises(NoTextCursorError):
        engine.cite(["KOWAL2023"])

    assert doc.read_slides() == before_slides
    assert doc.get_slide_body_text(1) == before_body
    assert doc.load_settings() is None
    assert doc.backups_created == []


def test_engine_cite_and_rebuild_without_open_presentation() -> None:
    zotero = FakeZoteroSource()
    doc = FakeSlideDoc(slide_count=2, active_slide_index=None)
    engine = CitationEngine(zotero, doc)

    with pytest.raises(NoActiveSlideError):
        engine.cite(["KOWAL2023"])

    with pytest.raises(NoActiveSlideError):
        engine.rebuild()

    with pytest.raises(ValueError):
        engine.cite([])


def test_engine_cite_single_multi_repeats_and_earlier_slide_insertion() -> None:
    zotero = FakeZoteroSource()
    doc = FakeSlideDoc(slide_count=3, active_slide_index=1, has_cursor=True)
    engine = CitationEngine(zotero, doc)

    # Slajd 1: pojedyncze cytowanie [1]
    s1 = engine.cite(["KOWAL2023"])
    assert [c.text for c in s1.citations] == ["[1]"]
    assert doc.get_slide_body_text(1) == "Tekst slajdu 1 [1]"
    assert [e.plain_text[:4] for e in s1.footer_entries] == ["[1] "]

    # Slajd 1: cytowanie zbiorcze (D13) zawierające powtórzenie KOWAL2023 oraz nową pozycję NOWAK2024
    s1 = engine.cite(["KOWAL2023", "NOWAK2024"])
    assert [c.text for c in s1.citations] == ["[1]", "[1, 2]"]
    assert doc.get_slide_body_text(1) == "Tekst slajdu 1 [1] [1, 2]"
    assert [e.key for e in s1.footer_entries] == ["KOWAL2023", "NOWAK2024"]

    # Slajd 3: cytujemy DOVE2013 -> [3] oraz powtórzenie KOWAL2023 -> [1]
    doc.active_slide_index = 3
    s3 = engine.cite(["DOVE2013"])
    assert [c.text for c in s3.citations] == ["[3]"]
    s3 = engine.cite(["KOWAL2023"])
    assert [c.text for c in s3.citations] == ["[3]", "[1]"]
    # Stopka na slajdzie 3 ma wpisy w kolejności numerycznej [1], [3]
    assert [e.prefix for e in s3.footer_entries] == ["[1] ", "[3] "]
    assert [e.key for e in s3.footer_entries] == ["KOWAL2023", "DOVE2013"]

    # Wstawienie nowej pozycji ZIEL2022 na slajdzie 2 (przed slajdem 3!) przesuwa numer DOVE2013 na slajdzie 3 z [3] na [4]
    doc.active_slide_index = 2
    s2 = engine.cite(["ZIEL2022"])
    assert [c.text for c in s2.citations] == ["[3]"]
    slides_now = doc.read_slides()
    assert [c.text for c in slides_now[2].citations] == ["[4]", "[1]"]
    assert doc.get_slide_body_text(3) == "Tekst slajdu 3 [4] [1]"
    assert [e.prefix for e in slides_now[2].footer_entries] == ["[1] ", "[4] "]

    # Pojedyncze „Cytuj” nie tworzy kopii zapasowej .bak.pptx
    assert doc.backups_created == []


def test_engine_cite_adjacent_merging_ieee_and_apa_d13() -> None:
    zotero = FakeZoteroSource()
    doc = MergingFakeSlideDoc(slide_count=2, active_slide_index=1, has_cursor=True)
    engine = CitationEngine(zotero, doc)

    # Slajd 1: pierwsze cytowanie NOWAK2024 -> [1], drugie KOWAL2023 bez scalania -> [2]
    doc.merge_adjacent = False
    engine.cite(["NOWAK2024"])
    engine.cite(["KOWAL2023"])
    assert doc.get_slide_body_text(1) == "Tekst slajdu 1 [1] [2]"

    # Slajd 2: najpierw cytujemy KOWAL2023 ([2]), a potem przy kursorze stoi sąsiednie cytowanie
    # i dokładamy NOWAK2024 ([1]) oraz DOVE2013 ([3]) -> scalone i posortowane [1, 2, 3]
    doc.active_slide_index = 2
    doc.merge_adjacent = False
    engine.cite(["KOWAL2023"])
    assert doc.get_slide_body_text(2) == "Tekst slajdu 2 [2]"

    doc.merge_adjacent = True
    s2 = engine.cite(["NOWAK2024", "DOVE2013"])
    assert len(s2.citations) == 1
    assert s2.citations[0].keys == ("KOWAL2023", "NOWAK2024", "DOVE2013")
    assert s2.citations[0].text == "[1, 2, 3]"
    assert doc.get_slide_body_text(2) == "Tekst slajdu 2 [1, 2, 3]"
    assert [e.prefix for e in s2.footer_entries] == ["[1] ", "[2] ", "[3] "]

    # Przełączenie na APA i kolejne scalenie przy kursorze z ZIEL2022
    apa_settings = Settings(style_id=STYLE_APA)
    engine.apply_settings(apa_settings)
    s2_apa = engine.cite(["ZIEL2022"], settings=apa_settings)
    assert len(s2_apa.citations) == 1
    assert (
        s2_apa.citations[0].text
        == "(Kowalski, 2023; Nowak & Wiśniewski, 2024; Dove et al., 2013; Zieliński, 2022)"
    )
    assert [e.key for e in s2_apa.footer_entries] == [
        "DOVE2013",
        "KOWAL2023",
        "NOWAK2024",
        "ZIEL2022",
    ]


def test_engine_rebuild_after_slide_reordering_renumbers_and_creates_backup() -> None:
    zotero = FakeZoteroSource()
    doc = FakeSlideDoc(slide_count=3, active_slide_index=1, has_cursor=True)
    engine = CitationEngine(zotero, doc)

    # Slajd 1: KOWAL2023 -> [1], NOWAK2024 -> [2]
    engine.cite(["KOWAL2023"])
    engine.cite(["NOWAK2024"])
    # Slajd 2: NOWAK2024 -> [2], KOWAL2023 -> [1], DOVE2013 -> [3]
    doc.active_slide_index = 2
    engine.cite(["NOWAK2024"])
    engine.cite(["KOWAL2023"])
    engine.cite(["DOVE2013"])
    assert doc.get_slide_body_text(1) == "Tekst slajdu 1 [1] [2]"
    assert doc.get_slide_body_text(2) == "Tekst slajdu 2 [2] [1] [3]"

    # Użytkownik zamienia kolejność slajdów 1 i 2 w PowerPoincie
    doc.reorder_slides([2, 1, 3])

    rebuilt = engine.rebuild()
    assert len(doc.backups_created) == 1
    assert doc.backups_created[0].name == "prezentacja_testowa.bak.pptx"

    # Po zamianie slajdów NOWAK2024 jest pierwsze -> [1], KOWAL2023 drugie -> [2], DOVE2013 -> [3]
    assert [c.text for c in rebuilt[0].citations] == ["[1]", "[2]", "[3]"]
    assert doc.get_slide_body_text(1) == "Tekst slajdu 2 [1] [2] [3]"
    assert [e.key for e in rebuilt[0].footer_entries] == ["NOWAK2024", "KOWAL2023", "DOVE2013"]

    assert [c.text for c in rebuilt[1].citations] == ["[2]", "[1]"]
    assert doc.get_slide_body_text(2) == "Tekst slajdu 1 [2] [1]"
    assert [e.key for e in rebuilt[1].footer_entries] == ["NOWAK2024", "KOWAL2023"]
    assert [e.prefix for e in rebuilt[1].footer_entries] == ["[1] ", "[2] "]


def test_engine_apply_settings_switches_styles_modes_and_footer_options_d14() -> None:
    zotero = FakeZoteroSource()
    doc = FakeSlideDoc(slide_count=2, active_slide_index=1, has_cursor=True)
    engine = CitationEngine(zotero, doc)

    # Slajd 1: ZIEL2022 ([1]), KOWAL2023 ([2])
    engine.cite(["ZIEL2022"])
    engine.cite(["KOWAL2023"])
    # Slajd 2: NOWAK2024 ([3]), DOVE2013 ([4]), ZIEL2022 ([1])
    doc.active_slide_index = 2
    engine.cite(["NOWAK2024", "DOVE2013"])
    engine.cite(["ZIEL2022"])

    assert doc.get_slide_body_text(1) == "Tekst slajdu 1 [1] [2]"
    assert doc.get_slide_body_text(2) == "Tekst slajdu 2 [3, 4] [1]"

    # 1. Przełączenie na tryb per_slide (D4, D14)
    s_per_slide = Settings(
        style_id=STYLE_IEEE_V4,
        numbering_mode="per_slide",
        font_name="Arial",
        font_size_pt=9,
        footer_position="bottom_left",
        footer_header=False,
    )
    slides_ps = engine.apply_settings(s_per_slide)
    assert len(doc.backups_created) == 1
    assert doc.get_slide_body_text(2) == "Tekst slajdu 2 [1, 2] [3]"
    assert [e.prefix for e in slides_ps[1].footer_entries] == ["[1] ", "[2] ", "[3] "]
    assert doc.get_slide_footer_settings(2) == s_per_slide
    assert doc.load_settings() == s_per_slide

    # 2. Przełączenie na styl APA 7 oraz położenie replace_footer (D3, D9, D14)
    s_apa = replace(
        s_per_slide,
        style_id=STYLE_APA,
        footer_position="replace_footer",
        footer_header=True,
    )
    slides_apa = engine.apply_settings(s_apa)
    assert len(doc.backups_created) == 2
    assert doc.get_slide_body_text(1) == "Tekst slajdu 1 (Zieliński, 2022) (Kowalski, 2023)"
    assert (
        doc.get_slide_body_text(2)
        == "Tekst slajdu 2 (Nowak & Wiśniewski, 2024; Dove et al., 2013) (Zieliński, 2022)"
    )
    # Alfabetyczna kolejność w stopkach APA
    assert [e.key for e in slides_apa[0].footer_entries] == ["KOWAL2023", "ZIEL2022"]
    assert [e.key for e in slides_apa[1].footer_entries] == ["DOVE2013", "NOWAK2024", "ZIEL2022"]
    assert doc.get_slide_footer_settings(1) == s_apa

    # 3. Powrót do IEEE v4 global (D2, D4, D14)
    s_ieee_global = Settings(style_id=STYLE_IEEE_V4, numbering_mode="global")
    engine.apply_settings(s_ieee_global)
    assert len(doc.backups_created) == 3
    assert doc.get_slide_body_text(1) == "Tekst slajdu 1 [1] [2]"
    assert doc.get_slide_body_text(2) == "Tekst slajdu 2 [3, 4] [1]"


def test_engine_resilience_to_zotero_offline_error() -> None:
    zotero = FakeZoteroSource()
    doc = FakeSlideDoc(slide_count=2, active_slide_index=1, has_cursor=True)
    engine = CitationEngine(zotero, doc)

    engine.cite(["KOWAL2023"])
    snapshot_slides = doc.read_slides()
    snapshot_body = doc.get_slide_body_text(1)

    # Zotero zostaje wyłączone
    zotero.offline = True

    with pytest.raises(ZoteroOfflineError):
        engine.cite(["NOWAK2024"])

    # Prezentacja nie została częściowo zmieniona
    assert doc.read_slides() == snapshot_slides
    assert doc.get_slide_body_text(1) == snapshot_body

    with pytest.raises(ZoteroOfflineError):
        engine.rebuild()

    # Nie utworzono zbędnej kopii zapasowej, gdy Zotero nie odpowiada
    assert doc.backups_created == []


def test_engine_callback_adapters() -> None:
    zotero = FakeZoteroSource()
    doc = FakeSlideDoc(slide_count=2, active_slide_index=1, has_cursor=True)
    engine = CitationEngine(zotero, doc)

    on_cite = engine.make_cite_callback()
    on_rebuild = engine.make_rebuild_callback()
    settings = Settings()

    cite_msg = on_cite(("KOWAL2023", "NOWAK2024"), settings)
    assert cite_msg == "Wstawiono cytowanie [1, 2] na slajdzie 1."

    rebuild_msg = on_rebuild(replace(settings, style_id=STYLE_APA))
    assert rebuild_msg == "Przebudowano prezentację (kopia zapasowa: prezentacja_testowa.bak.pptx)."
    assert doc.get_slide_body_text(1) == "Tekst slajdu 1 (Kowalski, 2023; Nowak & Wiśniewski, 2024)"


def test_h1_cite_uses_get_adjacent_citation_at_cursor_and_slides_up_to_active() -> None:
    class AdjacentAwareSlideDoc(FakeSlideDoc):
        def __init__(self) -> None:
            super().__init__(slide_count=2, active_slide_index=1, has_cursor=True)
            self.inserted_citations: list[Citation] = []
            self.update_calls: list[int] = []

        def get_adjacent_citation_at_cursor(self) -> Citation | None:
            slide = self._get_slide(self.get_active_slide_index())
            return slide.citations[-1] if slide.citations else None

        def insert_citation_at_cursor(self, citation: Citation) -> None:
            self.inserted_citations.append(citation)
            slide = self._get_slide(self.get_active_slide_index())
            if slide.citations and slide.citations[-1].citation_id == citation.citation_id:
                prev = slide.citations[-1]
                pos = slide.body_text.rfind(prev.text)
                if pos != -1:
                    slide.body_text = (
                        slide.body_text[:pos]
                        + citation.text
                        + slide.body_text[pos + len(prev.text) :]
                    )
                slide.citations[-1] = citation
                return
            super().insert_citation_at_cursor(citation)

        def update_slide_citations(self, slide_index: int, citations: object) -> None:
            self.update_calls.append(slide_index)
            super().update_slide_citations(slide_index, citations)  # type: ignore[arg-type]

    zotero = FakeZoteroSource()
    doc = AdjacentAwareSlideDoc()
    engine = CitationEngine(zotero, doc)

    # Najpierw cytujemy DOVE2013 na slajdzie 2 -> [1]
    doc.active_slide_index = 2
    engine.cite(["DOVE2013"])
    assert doc.inserted_citations[-1].text == "[1]"

    # Teraz przechodzimy na slajd 1 (przed slajdem 2!) i wstawiamy KOWAL2023 + NOWAK2024:
    # provisional_map dla slajdu 1 nie może brać kluczy ze slajdu 2, więc od razu dostaje "[1, 2]"
    doc.active_slide_index = 1
    engine.cite(["KOWAL2023", "NOWAK2024"])
    first_inserted = doc.inserted_citations[-1]
    assert first_inserted.text == "[1, 2]"

    # Przy kursorze stoi istniejące cytowanie [1, 2] ("KOWAL2023", "NOWAK2024");
    # ponowne zacytowanie "NOWAK2024" (duplikat) musi przekazać do insert_citation_at_cursor
    # to samo citation_id oraz docelowy tekst "[1, 2]" (bez tymczasowego "[1, 2, 2]")
    doc.update_calls.clear()
    engine.cite(["NOWAK2024"])
    second_inserted = doc.inserted_citations[-1]
    assert second_inserted.citation_id == first_inserted.citation_id
    assert second_inserted.keys == ("KOWAL2023", "NOWAK2024")
    assert second_inserted.text == "[1, 2]"
    # L1: skoro tekst cytowania na slajdzie 1 był od razu docelowy, update_slide_citations nie było potrzebne
    assert 1 not in doc.update_calls


def test_m1_cite_checks_has_text_cursor_before_get_active_slide_index_d10() -> None:
    class CursorFirstDoc(FakeSlideDoc):
        def has_open_presentation(self) -> bool:
            return True

        def has_text_cursor(self) -> bool:
            return False

        def get_active_slide_index(self) -> int:
            raise NoActiveSlideError("Brak zaznaczonego slajdu.")

    zotero = FakeZoteroSource()
    doc = CursorFirstDoc(slide_count=2, active_slide_index=None, has_cursor=False)
    engine = CitationEngine(zotero, doc)

    with pytest.raises(NoTextCursorError):
        engine.cite(["KOWAL2023"])


def test_m2_make_cite_callback_reports_inserted_citation_before_merged_citation() -> None:
    class PrependingSlideDoc(FakeSlideDoc):
        def insert_citation_at_cursor(self, citation: Citation) -> None:
            slide = self._get_slide(self.get_active_slide_index())
            slide.body_text = f"{citation.text} {slide.body_text}"
            slide.citations.insert(0, citation)

    zotero = FakeZoteroSource()
    doc = PrependingSlideDoc(slide_count=1, active_slide_index=1, has_cursor=True)
    engine = CitationEngine(zotero, doc)
    on_cite = engine.make_cite_callback()
    settings = Settings()

    # Najpierw na slajdzie jest cytowanie zbiorcze [1, 2]
    assert on_cite(("KOWAL2023", "NOWAK2024"), settings) == "Wstawiono cytowanie [1, 2] na slajdzie 1."

    # Teraz przed [1, 2] wstawiamy pojedyncze cytowanie ("KOWAL2023",) -> [1]
    msg = on_cite(("KOWAL2023",), settings)
    assert msg == "Wstawiono cytowanie [1] na slajdzie 1."
    assert engine.last_citation is not None
    assert engine.last_citation.text == "[1]"


def test_m3_callbacks_convert_key_error_to_azp_error() -> None:
    zotero = FakeZoteroSource()
    doc = FakeSlideDoc(slide_count=1, active_slide_index=1, has_cursor=True)
    engine = CitationEngine(zotero, doc)
    on_cite = engine.make_cite_callback()
    on_rebuild = engine.make_rebuild_callback()
    settings = Settings()

    with pytest.raises(AzpError, match="Nie znaleziono opisu pozycji w Zotero"):
        on_cite(("NIEISTNIEJACY_KLUCZ",), settings)

    # Ręczne wstrzyknięcie nieznanego klucza do slajdu i wywołanie on_rebuild
    doc.insert_citation_at_cursor(
        Citation(citation_id="c-bad", keys=("NIEISTNIEJACY_KLUCZ",), text="[1]")
    )
    with pytest.raises(AzpError, match="Nie znaleziono opisu pozycji w Zotero"):
        on_rebuild(settings)

