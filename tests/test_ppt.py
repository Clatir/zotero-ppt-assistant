"""Testy jednostkowe oraz integracyjne COM warstwy PowerPoint (`azp/ppt/`, Etap B2.1)."""

from __future__ import annotations

from pathlib import Path
import subprocess

import pytest

from azp.contracts import (
    STYLE_APA,
    STYLE_IEEE_V4,
    Citation,
    FooterEntry,
    NoActiveSlideError,
    NoTextCursorError,
    Settings,
    SlideDoc,
    TextSpan,
)
from azp.ppt import (
    FOOTER_HEADER_TEXT,
    PP_PLACEHOLDER_FOOTER,
    PP_SELECTION_TEXT,
    PowerPointSlideDoc,
    TAG_ROLE,
    TAG_ROLE_FOOTER,
    apply_citation_at_cursor_in_text,
    build_footer_text_and_ranges,
    compute_footer_geometry,
    deserialize_citations,
    deserialize_footer_entries,
    deserialize_settings,
    find_adjacent_citation_in_text,
    merge_citation_texts,
    merge_citations,
    replace_citations_in_text,
    replace_citations_two_pass,
    serialize_citations,
    serialize_footer_entries,
    serialize_settings,
)
from azp.contracts import BackupError
from azp.ppt.com_doc import (
    BIB_COLUMN_GAP_PT,
    BIB_MARGIN_BOTTOM_PT,
    BIB_MARGIN_TOP_PT,
    BIB_MARGIN_X_PT,
    PP_SELECTION_NONE,
    PP_SELECTION_SHAPES,
    PP_SELECTION_SLIDES,
    TAG_BIB_COLUMN,
    TAG_BIB_CONTINUATION,
    TAG_FOOTER_ENTRIES,
    TAG_ROLE_BIBLIOGRAPHY,
    TextEdit,
    apply_text_edits,
    choose_powerpoint_window,
    compute_bibliography_column_geometries,
    estimate_bibliography_font_size,
    plan_citation_insert,
    plan_citation_replacements,
    split_bibliography_entries_across_slides,
    split_bibliography_entries_into_columns,
)
from tests.fakes.fake_com import MockPresentation, MockShape, MockSlide, MockTextRange

TEST_PPTX_PATH = Path(__file__).resolve().parent / "data" / "test_ppt_b21.pptx"


def _is_powerpoint_running() -> bool:
    """Sprawdza, czy proces POWERPNT.EXE jest uruchomiony."""
    res = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
        capture_output=True,
        text=True,
        check=False,
    )
    return "powerpnt.exe" in res.stdout.lower()


# ---------------------------------------------------------------------------
# Testy jednostkowe (bez COM)
# ---------------------------------------------------------------------------


def test_powerpoint_slide_doc_satisfies_slide_doc_protocol() -> None:
    doc = PowerPointSlideDoc()
    assert isinstance(doc, SlideDoc)


def test_compute_footer_geometry_three_positions_d9() -> None:
    slide_w, slide_h, box_h = 960.0, 540.0, 40.0

    g_full = compute_footer_geometry(slide_w, slide_h, "bottom_full", height=box_h)
    g_left = compute_footer_geometry(slide_w, slide_h, "bottom_left", height=box_h)
    g_repl = compute_footer_geometry(slide_w, slide_h, "replace_footer", height=box_h)

    assert g_full.left == 36.0
    assert g_full.width == 960.0 - 72.0
    assert g_full.top == 540.0 - 40.0 - 28.0
    assert g_full.height == 40.0

    # bottom_left: ten sam lewy margines i wysokość od dołu, węższa szerokość (62%)
    assert g_left.left == g_full.left
    assert g_left.top == g_full.top
    assert g_left.width == pytest.approx(g_full.width * 0.62)
    assert g_left.width < g_full.width

    # replace_footer: pełna szerokość, niżej na slajdzie (w miejscu stopki PowerPointa)
    assert g_repl.left == g_full.left
    assert g_repl.width == g_full.width
    assert g_repl.top == 540.0 - 40.0 - 10.0
    assert g_repl.top > g_full.top

    with pytest.raises(ValueError, match="Nieprawidłowe położenie"):
        compute_footer_geometry(slide_w, slide_h, "top_right")  # type: ignore[arg-type]


def test_build_footer_text_and_ranges_with_and_without_header() -> None:
    entries = [
        FooterEntry(
            key="KOWAL2023",
            prefix="[1] ",
            spans=(
                TextSpan("A. Kowalski, "),
                TextSpan("Genome Med", italic=True),
                TextSpan(", 2023."),
            ),
        ),
        FooterEntry(
            key="NOWAK2024",
            prefix="[2] ",
            spans=(
                TextSpan("B. Nowak, "),
                TextSpan("AI Review", italic=True, bold=True),
                TextSpan(", 2024."),
            ),
        ),
    ]

    # Wariant z nagłówkiem "Źródła:"
    text_with_hdr, ranges_with_hdr = build_footer_text_and_ranges(entries, footer_header=True)
    lines = text_with_hdr.split("\r")
    assert len(lines) == 3
    assert lines[0] == FOOTER_HEADER_TEXT
    assert lines[1] == "[1] A. Kowalski, Genome Med, 2023."
    assert lines[2] == "[2] B. Nowak, AI Review, 2024."

    # Weryfikacja 1-based indeksów znakowych względem tekstu z pojedynczym '\r'
    italic_ranges = [r for r in ranges_with_hdr if r.italic]
    assert len(italic_ranges) == 2

    r1_start, r1_len, r1_italic, r1_bold = italic_ranges[0]
    assert r1_italic is True and r1_bold is False
    assert text_with_hdr[r1_start - 1 : r1_start - 1 + r1_len] == "Genome Med"

    r2_start, r2_len, r2_italic, r2_bold = italic_ranges[1]
    assert r2_italic is True and r2_bold is True
    assert text_with_hdr[r2_start - 1 : r2_start - 1 + r2_len] == "AI Review"

    # Wariant bez nagłówka "Źródła:"
    text_no_hdr, ranges_no_hdr = build_footer_text_and_ranges(entries, footer_header=False)
    assert not text_no_hdr.startswith(FOOTER_HEADER_TEXT)
    assert text_no_hdr.split("\r") == [
        "[1] A. Kowalski, Genome Med, 2023.",
        "[2] B. Nowak, AI Review, 2024.",
    ]
    italic_no_hdr = [r for r in ranges_no_hdr if r.italic]
    assert len(italic_no_hdr) == 2
    assert text_no_hdr[italic_no_hdr[0].start - 1 : italic_no_hdr[0].start - 1 + italic_no_hdr[0].length] == "Genome Med"
    assert text_no_hdr[italic_no_hdr[1].start - 1 : italic_no_hdr[1].start - 1 + italic_no_hdr[1].length] == "AI Review"

    # Puste wpisy
    empty_text, empty_ranges = build_footer_text_and_ranges([], footer_header=True)
    assert empty_text == ""
    assert empty_ranges == []


def test_cursor_adjacency_and_merging_citations_d13() -> None:
    c1 = Citation(citation_id="cit-1", keys=("KOWAL2023",), text="[1]")
    c2 = Citation(citation_id="cit-2", keys=("NOWAK2024",), text="[2]")

    # 1. Wstawienie pierwszego cytowania na końcu tekstu (automatyczna spacja)
    text0 = "Wyniki eksperymentu"
    text1, cits1, pos1 = apply_citation_at_cursor_in_text(text0, len(text0), c1, [])
    assert text1 == "Wyniki eksperymentu [1]"
    assert cits1 == [c1]
    assert pos1 == len(text1)

    # 2. Kursor stoi bezpośrednio po "[1]" -> wykrywa sąsiedztwo i scala w "[1, 2]"
    adj = find_adjacent_citation_in_text(text1, pos1, cits1)
    assert adj is not None
    assert adj[1] == c1

    text2, cits2, pos2 = apply_citation_at_cursor_in_text(text1, pos1, c2, cits1)
    assert text2 == "Wyniki eksperymentu [1, 2]"
    assert len(cits2) == 1
    assert cits2[0].citation_id == "cit-1"
    assert cits2[0].keys == ("KOWAL2023", "NOWAK2024")
    assert cits2[0].text == "[1, 2]"
    assert pos2 == len(text2)

    # 3. Kursor stoi ze spacją po cytowaniu lub bezpośrednio przed cytowaniem
    text_with_space = "Wyniki eksperymentu [1, 2]   dalszy tekst"
    pos_after_spaces = len("Wyniki eksperymentu [1, 2]  ")
    adj_space = find_adjacent_citation_in_text(text_with_space, pos_after_spaces, cits2)
    assert adj_space is not None and adj_space[1].citation_id == "cit-1"

    pos_before_cit = len("Wyniki eksperymentu")
    adj_before = find_adjacent_citation_in_text(text_with_space, pos_before_cit, cits2)
    assert adj_before is not None and adj_before[1].citation_id == "cit-1"

    # 4. Kursor oddzielony słowami od cytowania -> brak sąsiedztwa, osobne cytowanie
    pos_far = len(text_with_space)
    assert find_adjacent_citation_in_text(text_with_space, pos_far, cits2) is None
    c3 = Citation(citation_id="cit-3", keys=("WISN2025",), text="[3]")
    text3, cits3, _ = apply_citation_at_cursor_in_text(text_with_space, pos_far, c3, cits2)
    assert text3 == "Wyniki eksperymentu [1, 2]   dalszy tekst [3]"
    assert len(cits3) == 2
    assert cits3[1] == c3

    # 5. Scalanie w stylu APA (nawiasy okrągłe) oraz podstawienie przy tym samym citation_id
    apa1 = Citation(citation_id="apa-1", keys=("KOWAL2023",), text="(Kowalski, 2023)")
    apa2 = Citation(citation_id="apa-2", keys=("NOWAK2024",), text="(Nowak, 2024)")
    merged_apa = merge_citations(apa1, apa2)
    assert merged_apa.citation_id == "apa-1"
    assert merged_apa.keys == ("KOWAL2023", "NOWAK2024")
    assert merged_apa.text == "(Kowalski, 2023; Nowak, 2024)"

    # Brak duplikatów kluczy i podmiana przy zgodnym citation_id
    override = Citation(
        citation_id="apa-1",
        keys=("NOWAK2024", "WISN2025"),
        text="(Kowalski, 2023; Nowak, 2024; Wiśniewski, 2025)",
    )
    merged_override = merge_citations(merged_apa, override)
    assert merged_override.citation_id == "apa-1"
    assert merged_override.keys == ("KOWAL2023", "NOWAK2024", "WISN2025")
    assert merged_override.text == "(Kowalski, 2023; Nowak, 2024; Wiśniewski, 2025)"

    assert merge_citation_texts("[1]", "[2]") == "[1, 2]"
    assert merge_citation_texts("(A, 2020)", "(B, 2021)") == "(A, 2020; B, 2021)"


def test_two_pass_citation_replacement_handles_duplicates_and_swaps() -> None:
    # Przypadek 1: dwa cytowania wstawione z tym samym tekstem "[1]"
    old_dups = [
        Citation(citation_id="c1", keys=("K1",), text="[1]"),
        Citation(citation_id="c2", keys=("K2",), text="[1]"),
    ]
    new_numbered = [
        Citation(citation_id="c1", keys=("K1",), text="[1]"),
        Citation(citation_id="c2", keys=("K2",), text="[2]"),
    ]
    text = "Pierwsze zdanie [1] oraz drugie zdanie [1]."
    updated = replace_citations_in_text(text, old_dups, new_numbered)
    assert updated == "Pierwsze zdanie [1] oraz drugie zdanie [2]."

    # Przypadek 2: zamiana numerów miejscami [1] <-> [2]
    swapped = [
        Citation(citation_id="c1", keys=("K1",), text="[2]"),
        Citation(citation_id="c2", keys=("K2",), text="[1]"),
    ]
    updated_swapped = replace_citations_in_text(updated, new_numbered, swapped)
    assert updated_swapped == "Pierwsze zdanie [2] oraz drugie zdanie [1]."

    # Przypadek 3: podmiana rozłożona na wiele kształtów tekstowych na slajdzie
    shapes_texts = ["Tytuł slajdu [2]", "Treść główna [1]"]
    reverted = replace_citations_two_pass(shapes_texts, swapped, new_numbered)
    assert reverted == ["Tytuł slajdu [1]", "Treść główna [2]"]


def test_tag_serialization_roundtrip_citations_footer_and_settings() -> None:
    citations = (
        Citation(citation_id="c1", keys=("KOWAL2023", "NOWAK2024"), text="[1, 2]"),
        Citation(citation_id="c2", keys=("WISN2025",), text="[3]"),
    )
    raw_cits = serialize_citations(citations)
    assert deserialize_citations(raw_cits) == citations
    assert deserialize_citations("") == ()

    entries = (
        FooterEntry(
            key="KOWAL2023",
            prefix="[1] ",
            spans=(
                TextSpan("A. Kowalski, "),
                TextSpan("Genome Med", italic=True),
                TextSpan(", 2023."),
            ),
        ),
    )
    raw_entries = serialize_footer_entries(entries)
    assert deserialize_footer_entries(raw_entries) == entries
    assert deserialize_footer_entries("") == ()

    settings = Settings(
        style_id=STYLE_APA,
        numbering_mode="per_slide",
        font_name="Arial",
        font_size_pt=12,
        footer_position="bottom_left",
        footer_header=False,
        always_on_top=False,
    )
    raw_settings = serialize_settings(settings)
    assert deserialize_settings(raw_settings) == settings
    assert deserialize_settings("") is None


def test_active_window_selection_mode_with_mock_com_objects() -> None:
    """Sprawdza współpracę `PowerPointSlideDoc` z `app.ActiveWindow.Selection` na atrapie COM."""

    class _MockTags:
        def __init__(self) -> None:
            self._d: dict[str, str] = {}

        def __call__(self, key: str) -> str:
            return self._d.get(key, "")

        def Add(self, key: str, val: str) -> None:
            self._d[key] = val

        def Delete(self, key: str) -> None:
            self._d.pop(key, None)

    class _MockTextRange:
        def __init__(self, text: str = "", start: int = 1) -> None:
            self.Text = text
            self.Start = start

        def Characters(self, start: int, length: int = 0) -> _MockTextRange:
            return _MockTextRange(self.Text[start - 1 : start - 1 + length], start=start)

        def Select(self) -> None:
            pass

    class _MockTextFrame:
        def __init__(self, text: str) -> None:
            self.TextRange = MockTextRange(text)

    class _MockShape:
        def __init__(self, text: str) -> None:
            self.HasTextFrame = -1
            self.TextFrame = _MockTextFrame(text)
            self.Tags = _MockTags()

    class _MockSlide:
        def __init__(self, slide_index: int, slide_id: int, shape: _MockShape) -> None:
            self.SlideIndex = slide_index
            self.SlideID = slide_id
            self._shapes = [shape]
            self.Tags = _MockTags()

        @property
        def Shapes(self):  # type: ignore[no-untyped-def]
            class _ShapesCol:
                def __init__(self, items: list[_MockShape]) -> None:
                    self._items = items
                    self.Count = len(items)

                def __call__(self, idx: int) -> _MockShape:
                    return self._items[idx - 1]

            return _ShapesCol(self._shapes)

    shape = _MockShape("Analiza danych")
    slide = _MockSlide(1, 256, shape)

    class _MockSelection:
        def __init__(self) -> None:
            self.Type = PP_SELECTION_TEXT
            self.TextRange = _MockTextRange("", start=len(shape.TextFrame.TextRange.Text) + 1)

        def ShapeRange(self, idx: int) -> _MockShape:
            return shape

        def SlideRange(self, idx: int) -> _MockSlide:
            return slide

        def Unselect(self) -> None:
            self.Type = 0

    selection = _MockSelection()

    class _MockPres:
        def __init__(self) -> None:
            self.Path = ""
            self.FullName = "Prezentacja1"
            self.Tags = _MockTags()

        @property
        def Slides(self):  # type: ignore[no-untyped-def]
            class _SlidesCol:
                Count = 1

                def __call__(self, idx: int) -> _MockSlide:
                    return slide

            return _SlidesCol()

    class _MockApp:
        def __init__(self) -> None:
            self.ActivePresentation = _MockPres()

            class _Win:
                Selection = selection

            self.ActiveWindow = _Win()

            class _PresCol:
                Count = 1

            self.Presentations = _PresCol()

    app = _MockApp()
    doc = PowerPointSlideDoc(app=app, presentation=app.ActivePresentation)
    assert doc.has_open_presentation() is True
    assert doc.get_presentation_path() is None
    assert doc.get_active_slide_index() == 1
    assert doc.has_text_cursor() is True

    cit1 = Citation(citation_id="c1", keys=("KOWAL2023",), text="[1]")
    doc.insert_citation_at_cursor(cit1)
    assert shape.TextFrame.TextRange.Text == "Analiza danych [1]"

    # Przesuwamy kursor w ActiveWindow.Selection na koniec nowego tekstu i scalamy z c2
    selection.TextRange.Start = len(shape.TextFrame.TextRange.Text) + 1
    assert doc.get_adjacent_citation_at_cursor() == cit1

    cit2 = Citation(citation_id="c2", keys=("NOWAK2024",), text="[2]")
    doc.insert_citation_at_cursor(cit2)
    assert shape.TextFrame.TextRange.Text == "Analiza danych [1, 2]"

    # Wyłączenie kursora (Selection.Type = 0) -> D10 NoTextCursorError
    doc.set_cursor_target(None)
    assert doc.has_text_cursor() is False
    with pytest.raises(NoTextCursorError):
        doc.insert_citation_at_cursor(Citation(citation_id="c3", keys=("WISN2025",), text="[3]"))


def test_h1_insert_citation_before_existing_citation_preserves_text_order() -> None:
    """H1: Wstawienie cytowania PRZED istniejącym w tym samym polu zachowuje kolejność w tekście."""
    c1 = Citation(citation_id="c1", keys=("K1",), text="[1]")
    c2 = Citation(citation_id="c2", keys=("K2",), text="[1]")
    c3 = Citation(citation_id="c3", keys=("K3",), text="[2]")

    text0 = "Zdanie pierwsze. Zdanie drugie. [1]"
    pos_after_first = len("Zdanie pierwsze.")
    text1, cits1, pos1 = apply_citation_at_cursor_in_text(text0, pos_after_first, c2, [c1])

    assert text1 == "Zdanie pierwsze. [1] Zdanie drugie. [1]"
    # Kolejność cytowań na liście musi odpowiadać kolejności występowania w tekście: [c2, c1]
    assert cits1 == [c2, c1]

    # Kolejne wstawienie w tym samym miejscu kursora (tuż po pierwszym "[1]") musi scalić z c2, a nie z c1
    text2, cits2, _ = apply_citation_at_cursor_in_text(text1, pos1, c3, cits1)
    assert text2 == "Zdanie pierwsze. [1, 2] Zdanie drugie. [1]"
    assert cits2 == [
        Citation(citation_id="c2", keys=("K2", "K3"), text="[1, 2]"),
        c1,
    ]


def test_h2_multiple_text_shapes_on_single_slide_isolation_and_ordering() -> None:
    """H2: Wiele kształtów tekstowych na jednym slajdzie – izolacja kursora i kolejność kształtów."""

    class _MockTags:
        def __init__(self) -> None:
            self._d: dict[str, str] = {}

        def __call__(self, key: str) -> str:
            return self._d.get(key, "")

        def Add(self, key: str, val: str) -> None:
            self._d[key] = val

        def Delete(self, key: str) -> None:
            self._d.pop(key, None)

    class _MockTextRange:
        def __init__(self, text: str = "") -> None:
            self.Text = text

    class _MockTextFrame:
        def __init__(self, text: str) -> None:
            self.TextRange = MockTextRange(text)

    class _MockShape:
        def __init__(self, name: str, text: str) -> None:
            self.Name = name
            self.HasTextFrame = -1
            self.TextFrame = _MockTextFrame(text)
            self.Tags = _MockTags()

    class _MockSlide:
        def __init__(self, shapes: list[_MockShape]) -> None:
            self.SlideIndex = 1
            self.SlideID = 101
            self._shapes = shapes
            self.Tags = _MockTags()

        @property
        def Shapes(self):  # type: ignore[no-untyped-def]
            class _ShapesCol:
                def __init__(self, items: list[_MockShape]) -> None:
                    self._items = items
                    self.Count = len(items)

                def __call__(self, key: int | str) -> _MockShape:
                    if isinstance(key, int):
                        return self._items[key - 1]
                    for sh in self._items:
                        if sh.Name == key:
                            return sh
                    raise KeyError(key)

            return _ShapesCol(self._shapes)

    class _MockPres:
        def __init__(self, slide: _MockSlide) -> None:
            self.Path = ""
            self.FullName = "Prezentacja1"
            self.Tags = _MockTags()
            self._slide = slide

        @property
        def Slides(self):  # type: ignore[no-untyped-def]
            slide_ref = self._slide

            class _SlidesCol:
                Count = 1

                def __call__(self, idx: int) -> _MockSlide:
                    return slide_ref

            return _SlidesCol()

    # Część A: Shape1 ma c1 ("[1]"), Shape2 ma c2 ("[1]"). Scalenie w Shape2 musi scalić z c2, nie z c1.
    sh1 = _MockShape("Shape1", "Tekst pierwszy")
    sh2 = _MockShape("Shape2", "Tekst drugi")
    footer_sh = _MockShape("FooterShape", "Źródła:\r[1] Stopka")
    footer_sh.Tags.Add(TAG_ROLE, TAG_ROLE_FOOTER)

    slide_a = _MockSlide([sh1, sh2, footer_sh])
    pres_a = _MockPres(slide_a)
    doc_a = PowerPointSlideDoc(presentation=pres_a, active_slide_index=1)

    c1 = Citation(citation_id="c1", keys=("K1",), text="[1]")
    c2 = Citation(citation_id="c2", keys=("K2",), text="[1]")
    c3 = Citation(citation_id="c3", keys=("K3",), text="[2]")

    doc_a.set_cursor_target(1, "Shape1", None)
    doc_a.insert_citation_at_cursor(c1)
    assert sh1.TextFrame.TextRange.Text == "Tekst pierwszy [1]"

    doc_a.set_cursor_target(1, "Shape2", None)
    doc_a.insert_citation_at_cursor(c2)
    assert sh2.TextFrame.TextRange.Text == "Tekst drugi [1]"

    # Kursor nadal na końcu Shape2 -> sąsiednie cytowanie to c2, a wstawienie c3 scala z c2
    assert doc_a.get_adjacent_citation_at_cursor() == c2
    doc_a.insert_citation_at_cursor(c3)
    assert sh1.TextFrame.TextRange.Text == "Tekst pierwszy [1]"
    assert sh2.TextFrame.TextRange.Text == "Tekst drugi [1, 2]"
    assert doc_a.read_slides()[0].citations == (
        c1,
        Citation(citation_id="c2", keys=("K2", "K3"), text="[1, 2]"),
    )

    # Część B: Najpierw wstawienie c1 w Shape2, a potem c2 w Shape1 -> kolejność na slajdzie to (c2, c1)
    sh1_b = _MockShape("Shape1", "Nagłówek")
    sh2_b = _MockShape("Shape2", "Treść")
    slide_b = _MockSlide([sh1_b, sh2_b, footer_sh])
    pres_b = _MockPres(slide_b)
    doc_b = PowerPointSlideDoc(presentation=pres_b, active_slide_index=1)

    doc_b.set_cursor_target(1, "Shape2", None)
    doc_b.insert_citation_at_cursor(c1)
    doc_b.set_cursor_target(1, "Shape1", None)
    doc_b.insert_citation_at_cursor(c2)

    assert doc_b.read_slides()[0].citations == (c2, c1)

    # Przenumerowanie: c2 (w Shape1) -> [1], c1 (w Shape2) -> [2]; stopka pozostaje nietknięta
    doc_b.update_slide_citations(
        1,
        [
            Citation(citation_id="c2", keys=("K2",), text="[1]"),
            Citation(citation_id="c1", keys=("K1",), text="[2]"),
        ],
    )
    assert sh1_b.TextFrame.TextRange.Text == "Nagłówek [1]"
    assert sh2_b.TextFrame.TextRange.Text == "Treść [2]"
    assert footer_sh.TextFrame.TextRange.Text == "Źródła:\r[1] Stopka"


def test_m1_native_footer_placeholder_preserves_preexisting_hidden_state() -> None:
    """M1: Natywny placeholder stopki ukryty przez użytkownika przed AZP nie jest odkrywany."""

    class _MockTags:
        def __init__(self) -> None:
            self._d: dict[str, str] = {}

        def __call__(self, key: str) -> str:
            return self._d.get(key, "")

        def Add(self, key: str, val: str) -> None:
            self._d[key] = val

        def Delete(self, key: str) -> None:
            self._d.pop(key, None)

    class _MockPlaceholderFormat:
        Type = PP_PLACEHOLDER_FOOTER

    class _MockPlaceholderShape:
        def __init__(self, visible: int) -> None:
            self.Visible = visible
            self.PlaceholderFormat = _MockPlaceholderFormat()
            self.Tags = _MockTags()

    class _MockSlide:
        def __init__(self, shapes: list[_MockPlaceholderShape]) -> None:
            self._shapes = shapes

        @property
        def Shapes(self):  # type: ignore[no-untyped-def]
            items = self._shapes

            class _Col:
                Count = len(items)

                def __call__(self, idx: int) -> _MockPlaceholderShape:
                    return items[idx - 1]

            return _Col()

    user_hidden = _MockPlaceholderShape(visible=0)
    originally_visible = _MockPlaceholderShape(visible=-1)
    slide = _MockSlide([user_hidden, originally_visible])

    doc = PowerPointSlideDoc()
    doc._update_native_footer_placeholder_visibility(slide, hide_native_footer=True)
    assert user_hidden.Visible == 0
    assert originally_visible.Visible == 0

    # Powrót do bottom_full (hide_native_footer=False): tylko uprzednio widoczny wraca na -1
    doc._update_native_footer_placeholder_visibility(slide, hide_native_footer=False)
    assert user_hidden.Visible == 0
    assert originally_visible.Visible == -1


def test_m2_build_footer_text_normalizes_crlf_and_lf_inside_spans() -> None:
    """M2: Ewentualne `\\r\\n` lub `\\n` w `TextSpan.text` są normalizowane do pojedynczego `\\r`."""
    entries = [
        FooterEntry(
            key="K1",
            prefix="[1] ",
            spans=(
                TextSpan("Linia A\r\nLinia B "),
                TextSpan("Kursywa\nDruga", italic=True),
            ),
        )
    ]
    full_text, ranges = build_footer_text_and_ranges(entries, footer_header=True)
    assert "\n" not in full_text
    italic_ranges = [r for r in ranges if r.italic]
    assert len(italic_ranges) == 1
    r = italic_ranges[0]
    assert full_text[r.start - 1 : r.start - 1 + r.length] == "Kursywa\rDruga"


# ---------------------------------------------------------------------------
# C1: edycje tylko w obrębie cytowań (T10), kopia SaveCopyAs (D17), klucz prezentacji (D19),
# powrót do okna PowerPointa (D20)
# ---------------------------------------------------------------------------


def test_c1_plan_citation_insert_returns_minimal_edit() -> None:
    """Wstawienie i scalenie są opisane jedną małą edycją, zgodną z dotychczasowym wynikiem tekstu."""
    c1 = Citation(citation_id="c1", keys=("K1",), text="[1]")
    c2 = Citation(citation_id="c2", keys=("K2",), text="[2]")

    text = "Alfa Beta Gamma"
    edit, cits, pos = plan_citation_insert(text, len("Alfa Beta"), c1, [])
    assert edit == TextEdit(start=len("Alfa Beta"), length=0, new_text=" [1]")
    assert apply_text_edits(text, [edit]) == "Alfa Beta [1] Gamma"
    assert cits == [c1]
    assert pos == len("Alfa Beta [1]")

    text2 = "Alfa Beta [1] Gamma"
    edit2, cits2, pos2 = plan_citation_insert(text2, len("Alfa Beta [1]"), c2, [c1])
    assert edit2 == TextEdit(start=len("Alfa Beta "), length=3, new_text="[1, 2]")
    assert apply_text_edits(text2, [edit2]) == "Alfa Beta [1, 2] Gamma"
    assert cits2 == [Citation(citation_id="c1", keys=("K1", "K2"), text="[1, 2]")]
    assert pos2 == len("Alfa Beta [1, 2]")

    # Zgodność z dotychczasową funkcją tekstową
    assert apply_citation_at_cursor_in_text(text2, len("Alfa Beta [1]"), c2, [c1]) == (
        "Alfa Beta [1, 2] Gamma",
        cits2,
        pos2,
    )


@pytest.mark.parametrize(
    ("text", "old", "new"),
    [
        (
            "A [1] oraz B [1]",
            [Citation("a", ("K1",), "[1]"), Citation("b", ("K2",), "[1]")],
            [Citation("a", ("K1",), "[1]"), Citation("b", ("K2",), "[2]")],
        ),
        (
            "A [1] oraz B [2] i C [1]",
            [
                Citation("a", ("K1",), "[1]"),
                Citation("b", ("K2",), "[2]"),
                Citation("c", ("K1",), "[1]"),
            ],
            [
                Citation("a", ("K1",), "[2]"),
                Citation("b", ("K2",), "[1]"),
                Citation("c", ("K1",), "[2]"),
            ],
        ),
        (
            "Tekst [1, 2] dalej",
            [Citation("a", ("K1", "K2"), "[1, 2]")],
            [Citation("a", ("K1", "K2"), "(Kowalski, 2023; Nowak, 2024)")],
        ),
        (
            "Bez zmian [1]",
            [Citation("a", ("K1",), "[1]")],
            [Citation("a", ("K1",), "[1]")],
        ),
        (
            "Usunięty ręcznie tekst",
            [Citation("a", ("K1",), "[1]")],
            [Citation("a", ("K1",), "[2]")],
        ),
    ],
)
def test_c1_plan_citation_replacements_matches_two_pass_replacement(
    text: str, old: list[Citation], new: list[Citation]
) -> None:
    """Edycje cytowań dają ten sam tekst co dwuprzebiegowa podmiana, nie dotykając tekstu między nimi."""
    edits = plan_citation_replacements(text, old, new)
    assert apply_text_edits(text, edits) == replace_citations_in_text(text, old, new)
    # Edycje są rozłączne, posortowane malejąco (można je stosować od końca bez przeliczania)
    starts = [e.start for e in edits]
    assert starts == sorted(starts, reverse=True)
    for e in edits:
        assert text[e.start : e.start + e.length] in {c.text for c in old}
        assert e.new_text != text[e.start : e.start + e.length]


def _mock_doc_with_shape(text: str) -> tuple[PowerPointSlideDoc, MockShape, MockPresentation]:
    shape = MockShape("Body", text)
    slide = MockSlide([shape])
    pres = MockPresentation([slide])
    doc = PowerPointSlideDoc(presentation=pres, active_slide_index=1)
    return doc, shape, pres


def test_c1_insert_and_renumber_do_not_overwrite_whole_shape_text_t10() -> None:
    """T10: wstawienie, scalenie i przenumerowanie zmieniają tylko zakres cytowania (formatowanie zostaje)."""
    doc, shape, _ = _mock_doc_with_shape("Alfa Beta Gamma\rDrugi akapit")
    tr = shape.TextFrame.TextRange

    doc.set_cursor_target(1, "Body", len("Alfa Beta"))
    doc.insert_citation_at_cursor(Citation("c1", ("K1",), "[1]"))
    assert tr.Text == "Alfa Beta [1] Gamma\rDrugi akapit"

    doc.set_cursor_target(1, "Body", None)
    doc.insert_citation_at_cursor(Citation("c2", ("K2",), "[1]"))
    assert tr.Text == "Alfa Beta [1] Gamma\rDrugi akapit [1]"

    doc.update_slide_citations(
        1, [Citation("c1", ("K1",), "[2]"), Citation("c2", ("K2",), "[1]")]
    )
    assert tr.Text == "Alfa Beta [2] Gamma\rDrugi akapit [1]"

    doc.set_cursor_target(1, "Body", len("Alfa Beta [2]"))
    doc.insert_citation_at_cursor(Citation("c3", ("K3",), "[3]"))
    assert tr.Text == "Alfa Beta [2, 3] Gamma\rDrugi akapit [1]"

    assert tr.whole_text_assignments == 0
    for start, length, new_text in tr.edits:
        assert length in (0, 3)
        assert new_text in {" [1]", "[1]", "[2]", "[2, 3]"}


def test_c1_insert_in_selection_mode_places_caret_after_citation() -> None:
    """Tryb z oknem: po wstawieniu kursor (Select) stoi tuż za cytowaniem."""
    doc, shape, pres = _mock_doc_with_shape("Alfa Beta Gamma")
    tr = shape.TextFrame.TextRange

    class _Sel:
        Type = PP_SELECTION_TEXT
        TextRange = MockTextRange("", start=len("Alfa Beta") + 1)

        def ShapeRange(self, idx: int) -> MockShape:
            return shape

        def SlideRange(self, idx: int):  # type: ignore[no-untyped-def]
            return pres.Slides(1)

    class _Win:
        Selection = _Sel()

    class _App:
        ActiveWindow = _Win()
        ActivePresentation = pres

        class Presentations:
            Count = 1

    doc = PowerPointSlideDoc(app=_App(), presentation=pres)
    doc.insert_citation_at_cursor(Citation("c1", ("K1",), "[1]"))
    assert tr.Text == "Alfa Beta [1] Gamma"
    assert tr.selected_at == len("Alfa Beta [1]") + 1
    assert tr.whole_text_assignments == 0


def test_c1_create_backup_uses_save_copy_as_without_saving_original_d17() -> None:
    """D17: kopia przez SaveCopyAs obok pliku; oryginał nie jest zapisywany (brak Save)."""
    shape = MockShape("Body", "Tekst")
    pres = MockPresentation(
        [MockSlide([shape])],
        full_name="C:\\Prace\\Wyklad.pptx",
        path="C:\\Prace",
    )
    doc = PowerPointSlideDoc(presentation=pres, active_slide_index=1)

    backup = doc.create_backup()
    assert backup == Path("C:\\Prace\\Wyklad.bak.pptx")
    assert pres.save_copy_as_calls == ["C:\\Prace\\Wyklad.bak.pptx"]
    assert pres.save_calls == 0

    # Plik z makrami: kopia zachowuje rozszerzenie (.bak.pptm), żeby dało się ją otworzyć
    pres_m = MockPresentation(
        [MockSlide([MockShape("Body", "x")])],
        full_name="C:\\Prace\\Makra.pptm",
        path="C:\\Prace",
    )
    assert PowerPointSlideDoc(presentation=pres_m).create_backup() == Path(
        "C:\\Prace\\Makra.bak.pptm"
    )

    # Niezapisana prezentacja: brak kopii (silnik i okienko blokują przebudowę – D18)
    pres_new = MockPresentation([MockSlide([MockShape("Body", "x")])])
    assert PowerPointSlideDoc(presentation=pres_new).create_backup() is None
    assert pres_new.save_copy_as_calls == []


def test_c1_create_backup_failure_raises_backup_error() -> None:
    """Błąd SaveCopyAs (np. brak uprawnień) -> `BackupError`, a nie surowy błąd COM."""
    pres = MockPresentation(
        [MockSlide([MockShape("Body", "x")])],
        full_name="C:\\Prace\\Wyklad.pptx",
        path="C:\\Prace",
    )
    pres.save_copy_as_error = RuntimeError("Access denied")
    doc = PowerPointSlideDoc(presentation=pres)
    with pytest.raises(BackupError):
        doc.create_backup()


def test_c1_presentation_key_identifies_active_presentation_d19() -> None:
    """D19: klucz aktywnej prezentacji (FullName) odróżnia także dwie niezapisane prezentacje."""
    pres1 = MockPresentation([MockSlide([MockShape("Body", "x")])], full_name="Prezentacja1")
    pres2 = MockPresentation([MockSlide([MockShape("Body", "y")])], full_name="Prezentacja2")

    class _App:
        ActivePresentation = pres1

        class Presentations:
            Count = 2

    app = _App()
    doc = PowerPointSlideDoc(app=app)
    assert doc.get_presentation_key() == "Prezentacja1"
    app.ActivePresentation = pres2
    assert doc.get_presentation_key() == "Prezentacja2"

    class _NoPres:
        class Presentations:
            Count = 0

    assert PowerPointSlideDoc(app=_NoPres()).get_presentation_key() is None


def test_c1_choose_powerpoint_window_prefers_presentation_caption_d20() -> None:
    """D20: wybór okna PowerPointa do przywrócenia fokusu – najpierw to z nazwą prezentacji."""
    windows = [
        (101, "Inna.pptx - PowerPoint"),
        (202, "Wyklad.pptx - PowerPoint"),
        (303, ""),
    ]
    assert choose_powerpoint_window(windows, "Wyklad.pptx") == 202
    # Eksplorator może ukrywać rozszerzenia – dopasowanie po nazwie bez rozszerzenia
    assert choose_powerpoint_window([(7, "Wyklad - PowerPoint")], "Wyklad.pptx") == 7
    # Brak dopasowania: pierwsze okno z tytułem
    assert choose_powerpoint_window(windows, "Brak.pptx") == 101
    assert choose_powerpoint_window([], "Wyklad.pptx") is None


def test_c1_h1_choose_powerpoint_window_ignores_azp_own_window_d20() -> None:
    """H1 (C1): okno AZP („Asystent Zotero dla PowerPointa”) nie może być wybrane zamiast PowerPointa."""
    windows = [
        (999, "Asystent Zotero dla PowerPointa"),
        (202, "Zotero.pptx - PowerPoint"),
    ]
    assert choose_powerpoint_window(windows, "Zotero.pptx") == 202
    assert choose_powerpoint_window(windows, "Inna.pptx") == 202
    assert choose_powerpoint_window(windows, "") == 202
    assert (
        choose_powerpoint_window(
            [(999, "Asystent Zotero dla PowerPointa")], "Zotero.pptx"
        )
        is None
    )


def test_c1_h2_cite_before_existing_citations_keeps_cursor_after_updated_citation_d20() -> None:
    """H2 (C1): po przenumerowaniu podczas cite() kursor przesuwa się o zmianę długości cytowania."""
    shape = MockShape("Body", "Wstęp.Zdanie [1, 2, 3, 4, 5, 6, 7, 8, 9]")
    slide = MockSlide([shape])
    pres = MockPresentation([slide])
    doc = PowerPointSlideDoc(presentation=pres, active_slide_index=1)

    c_tail = Citation(
        citation_id="c-tail",
        keys=tuple(f"K{i}" for i in range(1, 10)),
        text="[1, 2, 3, 4, 5, 6, 7, 8, 9]",
    )
    doc.set_cursor_target(1, "Body", len("Wstęp.Zdanie [1, 2, 3, 4, 5, 6, 7, 8, 9]"))
    # Zapisz istniejące cytowanie na końcu tekstu
    shape.Tags.Add("AZP_CITATIONS", serialize_citations([c_tail]))
    slide.Tags.Add("AZP_CITATIONS", serialize_citations([c_tail]))

    # Stawiamy kursor po "Wstęp." (przed "Zdanie")
    doc.set_cursor_target(1, "Body", len("Wstęp."))
    # Krok 1 w cite(): insert_citation_at_cursor wstawia tymczasowe "[10]" (długość 4)
    doc.insert_citation_at_cursor(Citation("c-head", ("K10",), "[10]"))
    assert (
        shape.TextFrame.TextRange.Text
        == "Wstęp. [10]Zdanie [1, 2, 3, 4, 5, 6, 7, 8, 9]"
    )
    # Krok 2 w cite(): update_slide_citations zamienia "[10]" -> "[1]" (długość 3)
    doc.update_slide_citations(
        1,
        [
            Citation("c-head", ("K10",), "[1]"),
            Citation(
                "c-tail",
                tuple(f"K{i}" for i in range(1, 10)),
                "[2, 3, 4, 5, 6, 7, 8, 9, 10]",
            ),
        ],
    )
    assert (
        shape.TextFrame.TextRange.Text
        == "Wstęp. [1]Zdanie [2, 3, 4, 5, 6, 7, 8, 9, 10]"
    )

    # Kursor musi stać tuż za "[1]", więc kolejne wstawienie scala z c-head zamiast trafić w środek "Zdanie"
    assert doc.get_adjacent_citation_at_cursor() == Citation("c-head", ("K10",), "[1]")
    doc.insert_citation_at_cursor(Citation("c-head", ("K10", "K11"), "[1, 2]"))
    assert (
        shape.TextFrame.TextRange.Text
        == "Wstęp. [1, 2]Zdanie [2, 3, 4, 5, 6, 7, 8, 9, 10]"
    )

    # M2 (runda 2): dwie edycje przed/na kursorze w tym samym kształcie (współrzędne względem old_t)
    shape2 = MockShape("Body2", "A [1, 2] BKoniec")
    slide2 = MockSlide([shape2])
    pres2 = MockPresentation([slide2])
    doc2 = PowerPointSlideDoc(presentation=pres2, active_slide_index=1)
    c_first = Citation("c1", ("K1",), "[1, 2]")
    shape2.Tags.Add("AZP_CITATIONS", serialize_citations([c_first]))
    slide2.Tags.Add("AZP_CITATIONS", serialize_citations([c_first]))

    doc2.set_cursor_target(1, "Body2", len("A [1, 2] B"))
    doc2.insert_citation_at_cursor(Citation("c2", ("K2",), "[10]"))
    assert shape2.TextFrame.TextRange.Text == "A [1, 2] B [10]Koniec"

    doc2.update_slide_citations(
        1,
        [
            Citation("c1", ("K1",), "[1]"),
            Citation("c2", ("K2",), "[2]"),
        ],
    )
    assert shape2.TextFrame.TextRange.Text == "A [1] B [2]Koniec"
    assert doc2.get_adjacent_citation_at_cursor() == Citation("c2", ("K2",), "[2]")


def test_manual_deletion_and_reordering_of_citations_in_slide_text_prunes_ghosts() -> None:
    """Usunięcie cytowania z tekstu slajdu (Backspace) lub przeklejenie kolejności aktualizuje odczyt slajdu."""
    sh1 = MockShape("Shape1", "Pierwsze [1] oraz drugie [2]")
    sh2 = MockShape("Shape2", "Trzecie [3]")
    slide = MockSlide([sh1, sh2])
    pres = MockPresentation([slide])
    doc = PowerPointSlideDoc(presentation=pres, active_slide_index=1)

    c1 = Citation("c1", ("K1",), "[1]")
    c2 = Citation("c2", ("K2",), "[2]")
    c3 = Citation("c3", ("K3",), "[3]")
    sh1.Tags.Add("AZP_CITATIONS", serialize_citations([c1, c2]))
    sh2.Tags.Add("AZP_CITATIONS", serialize_citations([c3]))
    slide.Tags.Add("AZP_CITATIONS", serialize_citations([c1, c2, c3]))

    # 1. Użytkownik ręcznie usuwa "[1]" z tekstu Shape1
    sh1.TextFrame.TextRange.Text = "Pierwsze oraz drugie [2]"
    slides_after_del = doc.read_slides()
    assert slides_after_del[0].citations == (c2, c3)

    # 2. Użytkownik usuwa tekst "[3]" z Shape2 -> nie może on zostać doklejony do Shape1!
    sh2.TextFrame.TextRange.Text = "Trzecie bez cytowania"
    assert doc.read_slides()[0].citations == (c2,)

    # 3. Użytkownik przekleja zdania w Shape1 tak, że [2] jest po [1] lub odwrotnie
    sh1.TextFrame.TextRange.Text = "Najpierw [2], potem [1]"
    sh1.Tags.Add("AZP_CITATIONS", serialize_citations([c1, c2]))
    assert doc.read_slides()[0].citations == (c2, c1)


def test_corrupted_tags_deserialization_does_not_crash() -> None:
    """Uszkodzony JSON lub błędne typy w tagach AZP_* zwracają puste wartości zamiast wyjątku."""
    assert deserialize_citations("{uszkodzony") == ()
    assert deserialize_citations('[{"citation_id": "c1", "keys": [""]}]') == ()
    assert deserialize_footer_entries("{uszkodzony") == ()
    assert deserialize_footer_entries('[{"key": "K1", "spans": "nie-lista"}]') == ()
    assert deserialize_settings("{uszkodzony") is None
    assert deserialize_settings('{"font_size_pt": 999}') is None


def test_cursor_in_grouped_shape_table_and_notes_pane_filtering() -> None:
    """Kursor w zgrupowanym kształcie i tabeli jest wykrywany, a panel notatek prelegenta jest odrzucany."""
    child_shape = MockShape("ChildText", "Tekst w grupie")

    class _GroupItems:
        Count = 1

        def __call__(self, idx: int) -> MockShape:
            return child_shape

    class _GroupShape:
        Name = "Group1"
        HasTextFrame = 0
        Tags = child_shape.Tags
        GroupItems = _GroupItems()

    slide = MockSlide([_GroupShape()])  # type: ignore[list-item]
    pres = MockPresentation([slide])

    class _Sel:
        Type = PP_SELECTION_TEXT
        HasChildShapeRange = -1
        TextRange = MockTextRange("", start=len("Tekst w grupie") + 1)

        def ChildShapeRange(self, idx: int) -> MockShape:
            return child_shape

        def ShapeRange(self, idx: int) -> _GroupShape:
            return _GroupShape()

        def SlideRange(self, idx: int) -> MockSlide:
            return slide

    class _Pane:
        ViewType = 1  # ppViewSlide

    class _Win:
        Selection = _Sel()
        ActivePane = _Pane()

    class _App:
        ActiveWindow = _Win()
        ActivePresentation = pres

        class Presentations:
            Count = 1

    doc = PowerPointSlideDoc(app=_App(), presentation=pres)
    assert doc.has_text_cursor() is True
    doc.insert_citation_at_cursor(Citation("cg", ("K1",), "[1]"))
    assert child_shape.TextFrame.TextRange.Text == "Tekst w grupie [1]"

    # Gdy aktywny jest panel notatek prelegenta (ppViewNotesPage = 3), kursor jest odrzucany
    _Pane.ViewType = 3
    assert doc.has_text_cursor() is False


def test_powerpoint_addin_ppam_and_vbs_builder(tmp_path: Path) -> None:
    """Budowanie i instalacja dodatku `.ppam` na wstążkę PowerPointa tworzy poprawny pakiet OpenXML."""
    import zipfile
    from azp.ppt import (
        ADDIN_FILENAME,
        build_ppam_addin,
        build_vbs_launcher,
        install_powerpoint_addin,
        is_powerpoint_addin_installed,
        uninstall_powerpoint_addin,
    )

    addin_dir = tmp_path / "AddIns"
    appdata_dir = tmp_path / "AsystentZotero"
    ppam_path = install_powerpoint_addin(
        addin_dir=addin_dir,
        appdata_dir=appdata_dir,
        register_hkcu=False,
    )
    assert ppam_path == addin_dir / ADDIN_FILENAME
    assert ppam_path.is_file()
    assert is_powerpoint_addin_installed(addin_dir=addin_dir, check_hkcu=False) is True
    assert (appdata_dir / "launch_azp.vbs").is_file()
    assert "start_azp.pyw" in (appdata_dir / "launch_azp.vbs").read_text(encoding="utf-8")

    with zipfile.ZipFile(ppam_path, "r") as zf:
        names = set(zf.namelist())
        assert {
            "[Content_Types].xml",
            "_rels/.rels",
            "ppt/presentation.xml",
            "ppt/_rels/presentation.xml.rels",
            "ppt/vbaProject.bin",
            "customUI/customUI.xml",
            "customUI/customUI14.xml",
        }.issubset(names)
        ui14 = zf.read("customUI/customUI14.xml").decode("utf-8")
        assert 'idMso="TabAddIns"' in ui14
        assert 'onAction="OnLaunchAzp"' in ui14

    uninstall_powerpoint_addin(
        addin_dir=addin_dir,
        appdata_dir=appdata_dir,
        unregister_hkcu=False,
    )
    assert is_powerpoint_addin_installed(addin_dir=addin_dir, check_hkcu=False) is False


def test_add_footer_citation_on_picture_only_and_text_slides_d26() -> None:
    """D26: add_footer_citation działa na slajdzie ze samymi zdjęciami (brak pól tekstowych) oraz z polami tekstowymi."""
    # Slajd 1: tylko zdjęcie (HasTextFrame = 0, brak jakichkolwiek body_shapes)
    picture_shape = MockShape("Picture1", "", has_text_frame=0)
    slide1 = MockSlide([picture_shape], slide_id=101)
    # Slajd 2: zwykłe pole tekstowe z cytowaniem w tekście
    body_shape = MockShape("Body2", "Tekst z cytowaniem")
    slide2 = MockSlide([body_shape], slide_id=102)
    pres = MockPresentation([slide1, slide2])
    doc = PowerPointSlideDoc(presentation=pres, active_slide_index=1)

    assert doc.has_text_cursor() is False
    c_foot1 = Citation(citation_id="cf1", keys=("KOWAL2023",), text="", in_text=False)
    doc.add_footer_citation(c_foot1)

    slides = doc.read_slides()
    assert len(slides[0].citations) == 1
    assert slides[0].citations[0] == c_foot1

    # Kolejne wywołanie add_footer_citation na tym samym slajdzie scala klucze
    c_foot2 = Citation(citation_id="cf2", keys=("NOWAK2024", "KOWAL2023"), text="", in_text=False)
    doc.add_footer_citation(c_foot2)
    slides = doc.read_slides()
    assert len(slides[0].citations) == 1
    assert slides[0].citations[0] == Citation(
        citation_id="cf1",
        keys=("KOWAL2023", "NOWAK2024"),
        text="",
        in_text=False,
    )

    # Slajd 2: najpierw cytowanie w tekście (z kursorem), potem cytowanie stopkowe (bez kursora)
    doc.set_cursor_target(2, "Body2", None)
    c_in = Citation(citation_id="cin1", keys=("DOVE2013",), text="[3]", in_text=True)
    doc.insert_citation_at_cursor(c_in)
    assert body_shape.TextFrame.TextRange.Text == "Tekst z cytowaniem [3]"

    doc.clear_cursor_target()
    doc.set_active_slide(2)
    c_foot_s2 = Citation(citation_id="cf_s2", keys=("ZIEL2022",), text="", in_text=False)
    doc.add_footer_citation(c_foot_s2)

    assert body_shape.TextFrame.TextRange.Text == "Tekst z cytowaniem [3]"
    s2_cits = doc.read_slides()[1].citations
    assert s2_cits == (c_in, c_foot_s2)

    # update_slide_citations aktualizuje tekst dla in_text=True i zachowuje in_text=False na slajdzie
    doc.update_slide_citations(
        2,
        [
            Citation(citation_id="cin1", keys=("DOVE2013",), text="(Dove et al., 2013)", in_text=True),
            c_foot_s2,
        ],
    )
    assert body_shape.TextFrame.TextRange.Text == "Tekst z cytowaniem (Dove et al., 2013)"
    assert doc.read_slides()[1].citations == (
        Citation(citation_id="cin1", keys=("DOVE2013",), text="(Dove et al., 2013)", in_text=True),
        c_foot_s2,
    )


def test_get_active_slide_index_with_picture_slide_or_none_selection_d26() -> None:
    """D26: get_active_slide_index wykrywa slajd przy zaznaczonym zdjęciu, miniaturze slajdu i braku zaznaczenia."""
    pic = MockShape("Pic", "", has_text_frame=0)
    slide1 = MockSlide([MockShape("Body1", "A")], slide_id=101)
    slide2 = MockSlide([pic], slide_id=102)
    pres = MockPresentation([slide1, slide2])

    class _View:
        Slide = slide2

    class _Sel:
        Type = PP_SELECTION_SHAPES

        def SlideRange(self, idx: int) -> MockSlide:
            _ = idx
            return slide2

        def ShapeRange(self, idx: int) -> MockShape:
            _ = idx
            return pic

    class _Win:
        Selection = _Sel()
        View = _View()

    class _App:
        ActiveWindow = _Win()
        ActivePresentation = pres

        class Presentations:
            Count = 1

    doc = PowerPointSlideDoc(app=_App(), presentation=pres)
    assert doc.get_active_slide_index() == 2

    # Zaznaczona miniatura slajdu (PP_SELECTION_SLIDES)
    _Sel.Type = PP_SELECTION_SLIDES
    assert doc.get_active_slide_index() == 2

    # Brak zaznaczenia (PP_SELECTION_NONE), gdzie SlideRange rzuca błąd COM, a View.Slide wskazuje slajd 2
    class _EmptySel:
        Type = PP_SELECTION_NONE

        def SlideRange(self, idx: int) -> MockSlide:
            raise RuntimeError("Nothing selected")

    _Win.Selection = _EmptySel()  # type: ignore[assignment]
    assert doc.get_active_slide_index() == 2


def test_bibliography_layout_helpers_columns_fit_and_split_d27() -> None:
    """D27: czyste funkcje podziału na 1..4 kolumny, geometrii, zmniejszania czcionki i podziału na slajdy."""
    entries = [
        FooterEntry(
            key=f"K{i}",
            prefix=f"[{i}] ",
            spans=(TextSpan(f"Autor {i}, Tytuł pracy naukowej numer {i}, Wydawnictwo 202{i % 10}."),),
        )
        for i in range(1, 8)
    ]

    # Podział 7 wpisów na 1, 2, 3, 4 kolumny
    assert [len(c) for c in split_bibliography_entries_into_columns(entries, 1)] == [7]
    assert [len(c) for c in split_bibliography_entries_into_columns(entries, 2)] == [4, 3]
    assert [len(c) for c in split_bibliography_entries_into_columns(entries, 3)] == [3, 2, 2]
    assert [len(c) for c in split_bibliography_entries_into_columns(entries, 4)] == [2, 2, 2, 1]
    assert split_bibliography_entries_into_columns([], 2) == []

    with pytest.raises(ValueError):
        split_bibliography_entries_into_columns(entries, 5)

    # Geometria kolumn na slajdzie 960 x 540
    geoms_3 = compute_bibliography_column_geometries(960.0, 540.0, 3)
    assert len(geoms_3) == 3
    expected_w = (960.0 - 2 * BIB_MARGIN_X_PT - 2 * BIB_COLUMN_GAP_PT) / 3.0
    expected_h = 540.0 - BIB_MARGIN_TOP_PT - BIB_MARGIN_BOTTOM_PT
    for idx, g in enumerate(geoms_3):
        assert g.top == BIB_MARGIN_TOP_PT
        assert g.width == pytest.approx(expected_w)
        assert g.height == pytest.approx(expected_h)
        assert g.left == pytest.approx(BIB_MARGIN_X_PT + idx * (expected_w + BIB_COLUMN_GAP_PT))

    # estimate_bibliography_font_size: mała lista nie zmniejsza czcionki, duża lista zmniejsza
    assert estimate_bibliography_font_size(entries[:3], 960.0, 540.0, 12, 1) == 12
    many_entries = [
        FooterEntry(
            key=f"M{i}",
            prefix=f"[{i}] ",
            spans=(TextSpan("Bardzo długi opis bibliograficzny " * 5),),
        )
        for i in range(1, 25)
    ]
    shrunk_1col = estimate_bibliography_font_size(many_entries, 960.0, 540.0, 16, 1)
    shrunk_2col = estimate_bibliography_font_size(many_entries, 960.0, 540.0, 16, 2)
    assert 6 <= shrunk_1col < 16
    assert shrunk_2col >= shrunk_1col

    # split_bibliography_entries_across_slides: dzieli dużą listę na wiele slajdów
    slide_chunks = split_bibliography_entries_across_slides(many_entries, 960.0, 540.0, 16, 1)
    assert len(slide_chunks) > 1
    assert [e for chunk in slide_chunks for e in chunk] == many_entries


def test_set_slide_bibliography_no_header_columns_and_modes_d27() -> None:
    """D27: set_slide_bibliography tworzy kolumny bez nagłówka 'Źródła:', obsługuje czcionki i tryby dopasowania."""
    slide1 = MockSlide([MockShape("Body1", "Slajd 1")], slide_id=101)
    slide2 = MockSlide([], slide_id=102)
    pres = MockPresentation([slide1, slide2])
    doc = PowerPointSlideDoc(presentation=pres, active_slide_index=2)

    entries = [
        FooterEntry(
            key=f"K{i}",
            prefix=f"[{i}] ",
            spans=(
                TextSpan(f"A. Autor {i}, "),
                TextSpan(f"Czasopismo {i}", italic=True),
                TextSpan(", 2024."),
            ),
        )
        for i in range(1, 6)
    ]

    # 1. Tryb fixed_font w 2 kolumnach bez nagłówka
    settings_2col = Settings(
        bib_font_name="Georgia",
        bib_font_size_pt=14,
        bib_fit_mode="fixed_font",
        bib_columns=2,
    )
    used = doc.set_slide_bibliography(2, entries, settings_2col)
    assert used == 1
    assert pres.Slides.Count == 2

    bib_shapes = [
         slide2.Shapes(i)
        for i in range(1, slide2.Shapes.Count + 1)
        if slide2.Shapes(i).Tags(TAG_ROLE) == TAG_ROLE_BIBLIOGRAPHY
    ]
    assert len(bib_shapes) == 2
    assert bib_shapes[0].Tags(TAG_BIB_COLUMN) == "1"
    assert bib_shapes[1].Tags(TAG_BIB_COLUMN) == "2"
    for sh in bib_shapes:
        assert not sh.TextFrame.TextRange.Text.startswith(FOOTER_HEADER_TEXT)
        assert sh.TextFrame.TextRange.Font.Name == "Georgia"
        assert sh.TextFrame.TextRange.Font.Size == 14
        # Sprawdzenie ustawienia kursywy na podzakresie
        assert any(italic == -1 for _, _, italic, _ in sh.TextFrame.TextRange.subrange_formats)

    slides_read = doc.read_slides()
    assert slides_read[0].has_bibliography is False
    assert slides_read[1].has_bibliography is True

    # 2. Tryb split_slides z dużą czcionką -> tworzy slajdy kontynuacji z TAG_BIB_CONTINUATION="1"
    many_entries = [
        FooterEntry(
            key=f"L{i}",
            prefix=f"[{i}] ",
            spans=(TextSpan("Obszerny wpis bibliograficzny w spisie literatury " * 4),),
        )
        for i in range(1, 15)
    ]
    settings_split = Settings(
        bib_font_name="Calibri",
        bib_font_size_pt=20,
        bib_fit_mode="split_slides",
        bib_columns=1,
    )
    used_split = doc.set_slide_bibliography(2, many_entries, settings_split)
    assert used_split > 1
    assert pres.Slides.Count == 1 + used_split
    for cont_idx in range(3, pres.Slides.Count + 1):
        assert pres.Slides(cont_idx).Tags(TAG_BIB_CONTINUATION) == "1"
    # Slajdy kontynuacji mają has_bibliography=False, a główny slajd bibliografii ma True
    records_split = doc.read_slides()
    assert records_split[1].has_bibliography is True
    assert all(not r.has_bibliography for r in records_split[2:])

    # 3. Ponowna aktualizacja w trybie shrink_to_fit (4 kolumny) usuwa stare slajdy kontynuacji
    settings_shrink_4col = Settings(
        bib_font_name="Tahoma",
        bib_font_size_pt=18,
        bib_fit_mode="shrink_to_fit",
        bib_columns=4,
    )
    used_shrink = doc.set_slide_bibliography(2, many_entries, settings_shrink_4col)
    assert used_shrink == 1
    assert pres.Slides.Count == 2
    bib_shapes_4 = [
        slide2.Shapes(i)
        for i in range(1, slide2.Shapes.Count + 1)
        if slide2.Shapes(i).Tags(TAG_ROLE) == TAG_ROLE_BIBLIOGRAPHY
    ]
    assert len(bib_shapes_4) == 4
    assert all(sh.TextFrame.TextRange.Font.Size <= 18 for sh in bib_shapes_4)

    # 4. Wyczyszczenie bibliografii (entries=[]) usuwa kształty bibliografii i zwraca 0
    assert doc.set_slide_bibliography(2, [], settings_shrink_4col) == 0
    assert doc.read_slides()[1].has_bibliography is False


def test_add_footer_citation_after_in_text_citation_does_not_restore_stale_cursor_d26() -> None:
    """H1: po cytowaniu w tekście, odznaczeniu kursora i add_footer_citation stary kursor nie jest przywracany."""
    shape = MockShape("Body", "Treść slajdu")
    slide = MockSlide([shape], slide_id=101)
    pres = MockPresentation([slide])
    tr = shape.TextFrame.TextRange

    class _View:
        Slide = slide

    class _Sel:
        Type = PP_SELECTION_TEXT
        TextRange = MockTextRange("", start=len("Treść slajdu") + 1)

        def ShapeRange(self, idx: int) -> MockShape:
            _ = idx
            return shape

        def SlideRange(self, idx: int) -> MockSlide:
            _ = idx
            return slide

    sel = _Sel()

    class _Win:
        Selection = sel
        View = _View()

        def Activate(self) -> None:
            pass

    class _App:
        ActiveWindow = _Win()
        ActivePresentation = pres
        Visible = -1

        class Presentations:
            Count = 1

    doc = PowerPointSlideDoc(app=_App(), presentation=pres)
    assert doc.has_text_cursor() is True
    doc.insert_citation_at_cursor(Citation("c1", ("K1",), "[1]", in_text=True))
    assert tr.Text == "Treść slajdu [1]"
    assert tr.selected_at == len("Treść slajdu [1]") + 1

    # Użytkownik odznacza pole tekstowe (PP_SELECTION_NONE) i dodaje cytowanie stopkowe bez kursora
    sel.Type = PP_SELECTION_NONE
    tr.selected_at = None
    assert doc.has_text_cursor() is False

    doc.add_footer_citation(Citation("cf", ("K2",), "", in_text=False))
    doc.set_slide_footer(
        1,
        [
            FooterEntry("K1", (TextSpan("Opis 1"),), prefix="[1] "),
            FooterEntry("K2", (TextSpan("Opis 2"),), prefix="[2] "),
        ],
        Settings(),
    )
    doc.focus_presentation_window()

    # Stary kursor w polu tekstowym nie może zostać przywrócony
    assert tr.selected_at is None


def test_set_slide_bibliography_on_continuation_slide_and_index_shift_h2() -> None:
    """H2: wywołanie set_slide_bibliography na slajdzie kontynuacji cofa do głównego slajdu i aktualizuje indeksy."""
    slide1 = MockSlide([MockShape("Body1", "Slajd 1")], slide_id=101)
    slide2 = MockSlide([], slide_id=102)
    slide3 = MockSlide([MockShape("Body3", "Slajd 3")], slide_id=103)
    pres = MockPresentation([slide1, slide2, slide3])
    doc = PowerPointSlideDoc(presentation=pres, active_slide_index=3)

    many_entries = [
        FooterEntry(
            key=f"L{i}",
            prefix=f"[{i}] ",
            spans=(TextSpan("Obszerny wpis bibliograficzny w spisie literatury " * 4),),
        )
        for i in range(1, 10)
    ]
    settings_split = Settings(
        bib_font_name="Calibri",
        bib_font_size_pt=24,
        bib_fit_mode="split_slides",
        bib_columns=1,
    )
    # Wstawiamy wieloslajdową bibliografię na slajdzie 2 (przed slajdem 3!)
    used = doc.set_slide_bibliography(2, many_entries, settings_split)
    assert used > 1
    added = used - 1
    # Aktywny slajd (dawny slajd 3, slide_id=103) przesunął się o `added` pozycji
    assert doc.get_active_slide_index() == 3 + added
    assert pres.Slides(3 + added).SlideID == 103

    # Wywołanie set_slide_bibliography na slajdzie kontynuacji (indeks 3) aktualizuje główny slajd 2
    settings_single = Settings(
        bib_font_name="Calibri",
        bib_font_size_pt=12,
        bib_fit_mode="shrink_to_fit",
        bib_columns=2,
    )
    used_after = doc.set_slide_bibliography(3, many_entries[:3], settings_single)
    assert used_after == 1
    assert pres.Slides.Count == 3
    assert doc.get_active_slide_index() == 3
    assert pres.Slides(3).SlideID == 103
    assert doc.read_slides()[1].has_bibliography is True


def test_get_active_slide_index_prefers_visible_slide_over_stale_thumbnail_and_reads_panes_d28() -> None:
    """D28/T15: po przewinięciu widoku na slajd 2 (gdy miniatura slajdu 1 jest nadal zaznaczona)
    get_active_slide_index zwraca widoczny slajd 2, także przez win.Panes (ppViewSlide = 1).
    """
    slide1 = MockSlide([MockShape("Body1", "Slajd 1")], slide_id=101)
    slide2 = MockSlide([], slide_id=102)
    pres = MockPresentation([slide1, slide2])

    class _View2:
        Slide = slide2

    class _StaleSlideSelection:
        Type = PP_SELECTION_SLIDES

        def SlideRange(self, idx: int) -> MockSlide:
            _ = idx
            # Miniatura slajdu 1 pozostała zaznaczona na pasku bocznym
            return slide1

    class _WinScrolled:
        Selection = _StaleSlideSelection()
        View = _View2()

    class _AppScrolled:
        ActiveWindow = _WinScrolled()
        ActivePresentation = pres

        class Presentations:
            Count = 1

    doc = PowerPointSlideDoc(app=_AppScrolled(), presentation=pres)
    assert doc.get_active_slide_index() == 2

    # Przypadek 2: ActivePane to panel miniatur (ppViewThumbnails = 11), więc win.View.Slide rzuca wyjątek,
    # a właściwy slajd 2 znajduje się w win.Panes w panelu o ViewType == 1 (ppViewSlide)
    class _BrokenTopView:
        @property
        def Slide(self) -> MockSlide:
            raise RuntimeError("View.Slide niedostępne w aktywnym panelu miniatur")

    class _ThumbPane:
        ViewType = 11
        View = _BrokenTopView()

    class _SlidePane:
        ViewType = 1
        View = _View2()

    class _PanesCol:
        Count = 2

        def __call__(self, idx: int) -> object:
            return (_ThumbPane(), _SlidePane())[idx - 1]

    class _WinWithPanes:
        Selection = _StaleSlideSelection()
        View = _BrokenTopView()
        Panes = _PanesCol()

    class _AppWithPanes:
        ActiveWindow = _WinWithPanes()
        ActivePresentation = pres

        class Presentations:
            Count = 1

    doc_panes = PowerPointSlideDoc(app=_AppWithPanes(), presentation=pres)
    assert doc_panes.get_active_slide_index() == 2

    # Przypadek 3 (M1): sel_type == PP_SELECTION_TEXT, ale ShapeRange, SlideRange i win.View.Slide rzucają wyjątek,
    # więc get_active_slide_index przechodzi do win.Panes (ViewType == 1) i zwraca slajd 2
    class _BrokenTextSelection:
        Type = PP_SELECTION_TEXT

        def ShapeRange(self, idx: int) -> MockShape:
            raise RuntimeError("ShapeRange niedostępny")

        def SlideRange(self, idx: int) -> MockSlide:
            raise RuntimeError("SlideRange niedostępny")

    class _WinTextInPanes:
        Selection = _BrokenTextSelection()
        View = _BrokenTopView()
        Panes = _PanesCol()

    class _AppTextInPanes:
        ActiveWindow = _WinTextInPanes()
        ActivePresentation = pres

        class Presentations:
            Count = 1

    doc_text_panes = PowerPointSlideDoc(app=_AppTextInPanes(), presentation=pres)
    assert doc_text_panes.get_active_slide_index() == 2


def test_h1_manual_deletion_of_in_text_citation_clears_footer_and_frees_number_one() -> None:
    """H1 / H1-R2: ręczne usunięcie jedynego cytowania [1] z tekstu slajdu 1 mającego stopkę
    czyści stopkę i zwalnia [1] zarówno podczas cite() na kolejnym slajdzie, jak i podczas rebuild().
    """
    from azp.engine import CitationEngine
    from tests.fakes.fake_zotero import FakeZoteroSource

    # Scenariusz A (H1-R2): cite() na slajdzie 2 zaraz po ręcznym usunięciu "[1]" ze slajdu 1
    sh1_a = MockShape("Body1", "Tekst slajdu 1")
    sh2_a = MockShape("Body2", "Tekst slajdu 2")
    slide1_a = MockSlide([sh1_a], slide_id=101)
    slide2_a = MockSlide([sh2_a], slide_id=102)
    pres_a = MockPresentation(
        [slide1_a, slide2_a],
        full_name="C:\\Prace\\TestH1A.pptx",
        path="C:\\Prace",
    )
    doc_a = PowerPointSlideDoc(presentation=pres_a, active_slide_index=1)
    zotero = FakeZoteroSource()
    engine_a = CitationEngine(zotero, doc_a)

    doc_a.set_cursor_target(1, "Body1", None)
    engine_a.cite(["KOWAL2023"])
    assert sh1_a.TextFrame.TextRange.Text == "Tekst slajdu 1 [1]"
    assert doc_a._find_footer_shape(slide1_a) is not None

    # Ręczne usunięcie "[1]" ze slajdu 1 i od razu cite() na slajdzie 2 (bez rebuild!)
    sh1_a.TextFrame.TextRange.Text = "Tekst slajdu 1"
    doc_a.set_cursor_target(2, "Body2", None)
    s2_planned = engine_a.cite(["NOWAK2024"])
    assert doc_a._find_footer_shape(slide1_a) is None
    assert sh2_a.TextFrame.TextRange.Text == "Tekst slajdu 2 [1]"
    assert [c.text for c in s2_planned.citations] == ["[1]"]
    assert [e.prefix for e in s2_planned.footer_entries] == ["[1] "]
    assert [e.key for e in s2_planned.footer_entries] == ["NOWAK2024"]

    # Scenariusz B: rebuild() po ręcznym usunięciu "[1]" ze slajdu 1
    sh1 = MockShape("Body1", "Tekst slajdu 1")
    sh2 = MockShape("Body2", "Tekst slajdu 2")
    slide1 = MockSlide([sh1], slide_id=101)
    slide2 = MockSlide([sh2], slide_id=102)
    pres = MockPresentation(
        [slide1, slide2],
        full_name="C:\\Prace\\TestH1.pptx",
        path="C:\\Prace",
    )
    doc = PowerPointSlideDoc(presentation=pres, active_slide_index=1)
    engine = CitationEngine(zotero, doc)

    # Slajd 1: cytujemy KOWAL2023 -> [1] + stopka na slajdzie 1
    doc.set_cursor_target(1, "Body1", None)
    engine.cite(["KOWAL2023"])
    assert sh1.TextFrame.TextRange.Text == "Tekst slajdu 1 [1]"
    assert doc._find_footer_shape(slide1) is not None

    # Slajd 2: cytujemy NOWAK2024 -> [2] + stopka na slajdzie 2
    doc.set_cursor_target(2, "Body2", None)
    engine.cite(["NOWAK2024"])
    assert sh2.TextFrame.TextRange.Text == "Tekst slajdu 2 [2]"

    # Użytkownik ręcznie usuwa "[1]" z tekstu slajdu 1 w PowerPoincie
    sh1.TextFrame.TextRange.Text = "Tekst slajdu 1"

    # Przebudowa prezentacji: stopka na slajdzie 1 znika, a NOWAK2024 na slajdzie 2 dostaje numer [1]
    rebuilt = engine.rebuild()
    assert rebuilt[0].citations == ()
    assert rebuilt[0].footer_entries == ()
    assert doc._find_footer_shape(slide1) is None
    assert sh2.TextFrame.TextRange.Text == "Tekst slajdu 2 [1]"
    assert [e.prefix for e in rebuilt[1].footer_entries] == ["[1] "]
    assert [e.key for e in rebuilt[1].footer_entries] == ["NOWAK2024"]


def test_get_slide_bibliography_keys_from_tag_and_continuation_slides_d28() -> None:
    """D28/T15: get_slide_bibliography_keys odczytuje klucze z TAG_BIB_KEYS oraz ze slajdów kontynuacji."""
    from azp.ppt import TAG_BIB_KEYS

    slide1 = MockSlide([MockShape("Body1", "Slajd 1")], slide_id=101)
    slide2 = MockSlide([], slide_id=102)
    pres = MockPresentation([slide1, slide2])
    doc = PowerPointSlideDoc(presentation=pres, active_slide_index=2)

    many_entries = [
        FooterEntry(
            key=f"K{i}",
            prefix=f"[{i}] ",
            spans=(TextSpan("Obszerny wpis bibliograficzny w spisie literatury " * 4),),
        )
        for i in range(1, 10)
    ]
    expected_keys = tuple(f"K{i}" for i in range(1, 10))
    settings_split = Settings(
        bib_font_name="Calibri",
        bib_font_size_pt=24,
        bib_fit_mode="split_slides",
        bib_columns=1,
    )
    used = doc.set_slide_bibliography(2, many_entries, settings_split)
    assert used > 1
    assert slide2.Tags(TAG_BIB_KEYS) != ""

    # Odczyt zarówno z głównego slajdu (2), jak i ze slajdu kontynuacji (3)
    assert doc.get_slide_bibliography_keys(2) == expected_keys
    assert doc.get_slide_bibliography_keys(3) == expected_keys

    # Fallback: gdy TAG_BIB_KEYS jest pusty, odczytuje klucze z kształtów na slajdzie głównym i kontynuacji
    slide2.Tags.Delete(TAG_BIB_KEYS)
    assert doc.get_slide_bibliography_keys(2) == expected_keys
    assert doc.get_slide_bibliography_keys(3) == expected_keys

    # Wyczyszczenie bibliografii czyści również TAG_BIB_KEYS
    doc.set_slide_bibliography(2, [], settings_split)
    assert slide2.Tags(TAG_BIB_KEYS) == ""
    assert doc.get_slide_bibliography_keys(2) == ()

    # M2: gdy _insert_blank_continuation_slide używa fallbacku Duplicate(),
    # skopiowany slajd kontynuacji ma wyczyszczony TAG_BIB_KEYS
    orig_add = pres.Slides.__class__.Add
    try:
        def _fail_add(self_slides: object, index: int, layout: int = 12) -> MockSlide:
            raise RuntimeError("Slides.Add disabled to force Duplicate fallback")

        pres.Slides.__class__.Add = _fail_add  # type: ignore[method-assign]
        used_dup = doc.set_slide_bibliography(2, many_entries, settings_split)
        assert used_dup > 1
        assert pres.Slides(2).Tags(TAG_BIB_KEYS) != ""
        for cont_idx in range(3, pres.Slides.Count + 1):
            assert pres.Slides(cont_idx).Tags(TAG_BIB_CONTINUATION) == "1"
            assert pres.Slides(cont_idx).Tags(TAG_BIB_KEYS) == ""
    finally:
        pres.Slides.__class__.Add = orig_add  # type: ignore[method-assign]


# ---------------------------------------------------------------------------
# Testy integracyjne COM (@pytest.mark.com)
# ---------------------------------------------------------------------------


@pytest.mark.com
def test_powerpoint_com_b21_full_workflow() -> None:
    """Kompleksowy test integracyjny COM dla etapu B2.1 na pliku `tests/data/test_ppt_b21.pptx`."""
    if _is_powerpoint_running():
        pytest.fail("PowerPoint jest uruchomiony! Zamknij PowerPointa przed testami COM.")

    import pythoncom
    import win32com.client

    TEST_PPTX_PATH.parent.mkdir(parents=True, exist_ok=True)
    backup_expected_path = TEST_PPTX_PATH.with_suffix(".bak.pptx")
    for p in (TEST_PPTX_PATH, backup_expected_path):
        if p.exists():
            p.unlink()

    pythoncom.CoInitialize()
    app = None
    pres = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        # Tworzymy nową prezentację bez okna (WithWindow=0)
        pres = app.Presentations.Add(0)

        # Slajd 1: układ z tytułem i stopką (ppLayoutTitleOnly = 11) + pole tekstowe
        slide1 = pres.Slides.Add(1, 11)
        box1 = slide1.Shapes.AddTextbox(1, 60, 120, 600, 120)
        box1.Name = "BodyBox1"
        box1.TextFrame.TextRange.Text = "Pierwszy akapit"

        # Dodajemy na slajdzie 1 placeholder stopki, jeśli układ go nie ma, lub drugi slajd
        slide2 = pres.Slides.Add(2, 12)
        box2 = slide2.Shapes.AddTextbox(1, 60, 120, 600, 120)
        box2.Name = "BodyBox2"
        box2.TextFrame.TextRange.Text = "Drugi slajd: punkt A oraz punkt B"

        slide1_original_id = int(slide1.SlideID)
        slide2_original_id = int(slide2.SlideID)

        doc = PowerPointSlideDoc(app=app, presentation=pres, active_slide_index=1)
        assert isinstance(doc, SlideDoc)
        assert doc.has_open_presentation() is True
        # Przed zapisaniem pliku na dysk get_presentation_path() zwraca None
        assert doc.get_presentation_path() is None

        # --- 1. Blokada insert_citation_at_cursor przy braku kursora (D10) ---
        assert doc.has_text_cursor() is False
        with pytest.raises(NoTextCursorError):
            doc.insert_citation_at_cursor(
                Citation(citation_id="cit-blocked", keys=("KOWAL2023",), text="[1]")
            )
        assert box1.TextFrame.TextRange.Text == "Pierwszy akapit"
        assert doc.read_slides()[0].citations == ()

        # --- 2. Wstawienie cytowania i automatyczne scalenie sąsiedniego (D13) ---
        doc.set_cursor_target(slide_index=1, shape_index_or_name="BodyBox1", char_pos=None)
        assert doc.has_text_cursor() is True
        assert doc.get_active_slide_index() == 1

        c1 = Citation(citation_id="cit-1", keys=("KOWAL2023",), text="[1]")
        doc.insert_citation_at_cursor(c1)
        assert box1.TextFrame.TextRange.Text == "Pierwszy akapit [1]"
        assert doc.get_adjacent_citation_at_cursor() == c1

        # Drugie wstawienie w tym samym miejscu kursora -> automatyczne scalenie w [1, 2]
        c2 = Citation(citation_id="cit-2", keys=("NOWAK2024",), text="[2]")
        doc.insert_citation_at_cursor(c2)
        assert box1.TextFrame.TextRange.Text == "Pierwszy akapit [1, 2]"
        slides_state = doc.read_slides()
        assert len(slides_state[0].citations) == 1
        assert slides_state[0].citations[0] == Citation(
            citation_id="cit-1",
            keys=("KOWAL2023", "NOWAK2024"),
            text="[1, 2]",
        )

        # --- 3. Pole źródeł (stopka) w 3 położeniach, z nagłówkiem i bez, kursywa (D5, D9, T4) ---
        entries_s1 = [
            FooterEntry(
                key="KOWAL2023",
                prefix="[1] ",
                spans=(
                    TextSpan("A. Kowalski, "),
                    TextSpan("Genome Med", italic=True),
                    TextSpan(", 2023."),
                ),
            ),
            FooterEntry(
                key="NOWAK2024",
                prefix="[2] ",
                spans=(
                    TextSpan("B. Nowak, "),
                    TextSpan("Przegląd AI", italic=True),
                    TextSpan(", 2024."),
                ),
            ),
        ]

        # 3a. bottom_full z nagłówkiem "Źródła:"
        settings_full = Settings(
            style_id=STYLE_IEEE_V4,
            font_name="Arial",
            font_size_pt=11,
            footer_position="bottom_full",
            footer_header=True,
        )
        doc.set_slide_footer(1, entries_s1, settings_full)
        footer_shape = None
        for i in range(1, int(slide1.Shapes.Count) + 1):
            sh = slide1.Shapes(i)
            if sh.Tags(TAG_ROLE) == TAG_ROLE_FOOTER:
                footer_shape = sh
                break
        assert footer_shape is not None
        full_width = float(footer_shape.Width)
        full_top = float(footer_shape.Top)

        tr = footer_shape.TextFrame.TextRange
        assert tr.Text.startswith("Źródła:\r[1] A. Kowalski, Genome Med, 2023.")
        assert tr.Font.Name == "Arial"
        assert int(tr.Font.Size) == 11

        # Weryfikacja kursywy na dokładnym podzakresie Characters(start, length) w COM
        _, ranges = build_footer_text_and_ranges(entries_s1, footer_header=True)
        italic_ranges = [r for r in ranges if r.italic]
        non_italic_ranges = [r for r in ranges if not r.italic]
        assert len(italic_ranges) == 2

        sub_italic_1 = tr.Characters(italic_ranges[0].start, italic_ranges[0].length)
        assert sub_italic_1.Text == "Genome Med"
        assert int(sub_italic_1.Font.Italic) == -1

        sub_italic_2 = tr.Characters(italic_ranges[1].start, italic_ranges[1].length)
        assert sub_italic_2.Text == "Przegląd AI"
        assert int(sub_italic_2.Font.Italic) == -1

        sub_plain = tr.Characters(non_italic_ranges[0].start, non_italic_ranges[0].length)
        assert sub_plain.Text == "Źródła:"
        assert int(sub_plain.Font.Italic) == 0

        # 3b. bottom_left bez nagłówka "Źródła:"
        settings_left = Settings(
            style_id=STYLE_IEEE_V4,
            font_name="Calibri",
            font_size_pt=10,
            footer_position="bottom_left",
            footer_header=False,
        )
        doc.set_slide_footer(1, entries_s1, settings_left)
        left_width = float(footer_shape.Width)
        assert left_width < full_width
        assert not footer_shape.TextFrame.TextRange.Text.startswith("Źródła:")
        assert footer_shape.TextFrame.TextRange.Font.Name == "Calibri"
        assert int(footer_shape.TextFrame.TextRange.Font.Size) == 10

        # 3c. replace_footer z nagłówkiem "Źródła:"
        settings_replace = Settings(
            style_id=STYLE_IEEE_V4,
            font_name="Calibri",
            font_size_pt=10,
            footer_position="replace_footer",
            footer_header=True,
        )
        doc.set_slide_footer(1, entries_s1, settings_replace)
        replace_width = float(footer_shape.Width)
        replace_top = float(footer_shape.Top)
        assert replace_width == pytest.approx(full_width)
        assert replace_top > full_top
        for i in range(1, int(slide1.Shapes.Count) + 1):
            sh = slide1.Shapes(i)
            try:
                if int(sh.PlaceholderFormat.Type) == PP_PLACEHOLDER_FOOTER:
                    assert int(sh.Visible) == 0
            except Exception:
                pass

        # --- 4. Slajd 2: dwa cytowania (duplikaty [1], [1]) i dwuprzebiegowa podmiana ---
        pos_after_a = len("Drugi slajd: punkt A")
        doc.set_cursor_target(slide_index=2, shape_index_or_name="BodyBox2", char_pos=pos_after_a)
        doc.insert_citation_at_cursor(Citation(citation_id="s2-c1", keys=("KOWAL2023",), text="[1]"))

        doc.set_cursor_target(slide_index=2, shape_index_or_name="BodyBox2", char_pos=None)
        doc.insert_citation_at_cursor(Citation(citation_id="s2-c2", keys=("WISN2025",), text="[1]"))
        assert box2.TextFrame.TextRange.Text == "Drugi slajd: punkt A [1] oraz punkt B [1]"

        # Przenumerowanie [1], [1] -> [1], [2]
        doc.update_slide_citations(
            2,
            [
                Citation(citation_id="s2-c1", keys=("KOWAL2023",), text="[1]"),
                Citation(citation_id="s2-c2", keys=("WISN2025",), text="[2]"),
            ],
        )
        assert box2.TextFrame.TextRange.Text == "Drugi slajd: punkt A [1] oraz punkt B [2]"

        # Zamiana numerów miejscami [1] <-> [2]
        doc.update_slide_citations(
            2,
            [
                Citation(citation_id="s2-c1", keys=("KOWAL2023",), text="[2]"),
                Citation(citation_id="s2-c2", keys=("WISN2025",), text="[1]"),
            ],
        )
        assert box2.TextFrame.TextRange.Text == "Drugi slajd: punkt A [2] oraz punkt B [1]"

        # --- 5. Zapis ustawień, zapis pliku .pptx, kopia zapasowa .bak.pptx ---
        saved_settings = Settings(
            style_id=STYLE_APA,
            numbering_mode="per_slide",
            font_name="Calibri",
            font_size_pt=10,
            footer_position="replace_footer",
            footer_header=True,
            always_on_top=False,
        )
        doc.save_settings(saved_settings)
        assert doc.load_settings() == saved_settings

        pres.SaveAs(str(TEST_PPTX_PATH))
        assert doc.get_presentation_path() == TEST_PPTX_PATH

        backup_path = doc.create_backup()
        assert backup_path == backup_expected_path
        assert backup_path is not None and backup_path.exists()

        # --- 6. Zamknięcie i ponowne otwarcie pliku: trwałość tagów ---
        pres.Close()
        pres = None

        pres = app.Presentations.Open(
            str(TEST_PPTX_PATH),
            ReadOnly=False,
            Untitled=False,
            WithWindow=0,
        )
        reopened_doc = PowerPointSlideDoc(app=app, presentation=pres, active_slide_index=1)
        assert reopened_doc.load_settings() == saved_settings

        reopened_slides = reopened_doc.read_slides()
        assert len(reopened_slides) == 2
        assert reopened_slides[0].slide_id == slide1_original_id
        assert reopened_slides[0].citations == (
            Citation(citation_id="cit-1", keys=("KOWAL2023", "NOWAK2024"), text="[1, 2]"),
        )
        assert reopened_slides[0].footer_entries == tuple(entries_s1)

        assert reopened_slides[1].slide_id == slide2_original_id
        assert reopened_slides[1].citations == (
            Citation(citation_id="s2-c1", keys=("KOWAL2023",), text="[2]"),
            Citation(citation_id="s2-c2", keys=("WISN2025",), text="[1]"),
        )

        # --- 7. Przestawienie kolejności slajdów (slajd 2 na pozycję 1) ---
        pres.Slides(2).MoveTo(1)
        reordered_slides = reopened_doc.read_slides()
        assert reordered_slides[0].slide_id == slide2_original_id
        assert reordered_slides[0].slide_index == 1
        assert reordered_slides[1].slide_id == slide1_original_id
        assert reordered_slides[1].slide_index == 2

        # Usunięcie stopki gdy entries=[]
        reopened_doc.set_slide_footer(2, [], saved_settings)
        assert reopened_doc.read_slides()[1].footer_entries == ()

        # --- 8. Cytowanie stopkowe bez kursora (D26) i bibliografia na slajdzie (D27) ---
        reopened_doc.clear_cursor_target()
        reopened_doc.set_active_slide(2)
        assert reopened_doc.has_text_cursor() is False
        reopened_doc.add_footer_citation(
            Citation(citation_id="foot-com-1", keys=("DOVE2013",), text="", in_text=False)
        )
        s2_cits_after_foot = reopened_doc.read_slides()[1].citations
        assert any(not c.in_text and c.keys == ("DOVE2013",) for c in s2_cits_after_foot)

        bib_slide = pres.Slides.Add(3, 12)
        _ = bib_slide
        bib_settings = Settings(
            style_id=STYLE_IEEE_V4,
            bib_font_name="Calibri",
            bib_font_size_pt=11,
            bib_fit_mode="shrink_to_fit",
            bib_columns=2,
        )
        slides_used = reopened_doc.set_slide_bibliography(3, entries_s1, bib_settings)
        assert slides_used == 1
        assert reopened_doc.read_slides()[2].has_bibliography is True

        pres.Save()
        pres.Close()
        pres = None
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        app = None
        import gc
        import time

        gc.collect()
        pythoncom.CoUninitialize()
        for _ in range(25):
            if not _is_powerpoint_running():
                break
            time.sleep(0.2)
        if backup_expected_path.exists():
            try:
                backup_expected_path.unlink()
            except Exception:
                pass
