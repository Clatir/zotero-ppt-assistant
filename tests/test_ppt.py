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
            self.TextRange = _MockTextRange(text)

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
            self.TextRange = _MockTextRange(text)

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
        pythoncom.CoUninitialize()
        if backup_expected_path.exists():
            try:
                backup_expected_path.unlink()
            except Exception:
                pass
