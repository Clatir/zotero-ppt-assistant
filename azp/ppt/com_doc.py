"""Implementacja warstwy COM Microsoft PowerPoint (`PowerPointSlideDoc`, Etap B2.1)."""

from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
import shutil
from typing import Any, NamedTuple, Sequence

from azp.contracts import (
    VALID_FOOTER_POSITIONS,
    BackupError,
    Citation,
    FooterEntry,
    FooterPosition,
    NoActiveSlideError,
    NoTextCursorError,
    Settings,
    SlideDoc,
    SlideRecord,
    TextSpan,
)

# Stałe COM PowerPoint / Office
PP_SELECTION_NONE = 0
PP_SELECTION_SLIDES = 1
PP_SELECTION_SHAPES = 2
PP_SELECTION_TEXT = 3

PP_PLACEHOLDER_FOOTER = 15
PP_AUTOSIZE_SHAPE_TO_FIT_TEXT = 1
MSO_TEXT_ORIENTATION_HORIZONTAL = 1
MSO_TRUE = -1
MSO_FALSE = 0

# Klucze tagów AZP w prezentacji, slajdach i kształtach (T3)
TAG_SETTINGS = "AZP_SETTINGS"
TAG_CITATIONS = "AZP_CITATIONS"
TAG_ROLE = "AZP_ROLE"
TAG_ROLE_FOOTER = "FOOTER"
TAG_FOOTER_ENTRIES = "AZP_FOOTER_ENTRIES"
TAG_HIDDEN_NATIVE_FOOTER = "AZP_HIDDEN_NATIVE_FOOTER"

FOOTER_HEADER_TEXT = "Źródła:"
FOOTER_MARGIN_X_PT = 36.0
FOOTER_BOTTOM_OFFSET_ABOVE_PT = 28.0
FOOTER_BOTTOM_OFFSET_REPLACE_PT = 10.0
FOOTER_LEFT_WIDTH_RATIO = 0.62
DEFAULT_FOOTER_HEIGHT_PT = 36.0


class FooterGeometry(NamedTuple):
    """Wyliczone położenie i wymiary kształtu pola źródeł na slajdzie (w punktach)."""

    left: float
    top: float
    width: float
    height: float


class CharFormatRange(NamedTuple):
    """1-based zakres znaków dla `TextRange.Characters(start, length)` z formatowaniem."""

    start: int
    length: int
    italic: bool = False
    bold: bool = False


class TextEdit(NamedTuple):
    """0-based edycja podzakresu tekstu kształtu (`start`, `length`, `new_text`) – T10."""

    start: int
    length: int
    new_text: str


@dataclass
class _CursorTarget:
    """Programowy cel kursora (używany w testach COM bez okna `WithWindow=0`)."""

    slide_index: int
    shape_index_or_name: int | str | None = None
    char_pos: int | None = None


def compute_footer_geometry(
    slide_width: float,
    slide_height: float,
    position: FooterPosition,
    height: float = DEFAULT_FOOTER_HEIGHT_PT,
) -> FooterGeometry:
    """Wylicza geometrię `(left, top, width, height)` pola źródeł dla 3 położeń (D9)."""
    if position not in VALID_FOOTER_POSITIONS:
        raise ValueError(f"Nieprawidłowe położenie stopki: {position!r}")

    w_slide = float(slide_width)
    h_slide = float(slide_height)
    h_box = float(height)
    full_width = max(50.0, w_slide - 2.0 * FOOTER_MARGIN_X_PT)

    if position == "bottom_full":
        left = FOOTER_MARGIN_X_PT
        width = full_width
        top = h_slide - h_box - FOOTER_BOTTOM_OFFSET_ABOVE_PT
    elif position == "bottom_left":
        left = FOOTER_MARGIN_X_PT
        width = full_width * FOOTER_LEFT_WIDTH_RATIO
        top = h_slide - h_box - FOOTER_BOTTOM_OFFSET_ABOVE_PT
    else:  # "replace_footer"
        left = FOOTER_MARGIN_X_PT
        width = full_width
        top = h_slide - h_box - FOOTER_BOTTOM_OFFSET_REPLACE_PT

    return FooterGeometry(left=left, top=top, width=width, height=h_box)


def build_footer_text_and_ranges(
    entries: Sequence[FooterEntry],
    footer_header: bool = True,
) -> tuple[str, list[CharFormatRange]]:
    """Składa tekst pola źródeł oraz 1-based zakresy formatowania (`Characters(start, len)`).

    W PowerPoint COM wewnątrz `TextRange.Text` znak nowej linii jest pojedynczym
    znakiem `\\r` (CR, długość 1). Funkcja używa `\\r` jako separatora wierszy,
    dzięki czemu offsety znakowe 1-based odpowiadają dokładnie indeksacji COM.
    """
    if not entries:
        return "", []

    parts: list[str] = []
    ranges: list[CharFormatRange] = []
    cursor_1based = 1

    if footer_header:
        parts.append(FOOTER_HEADER_TEXT)
        ranges.append(
            CharFormatRange(
                start=cursor_1based,
                length=len(FOOTER_HEADER_TEXT),
                italic=False,
                bold=False,
            )
        )
        cursor_1based += len(FOOTER_HEADER_TEXT)
        parts.append("\r")
        cursor_1based += 1

    for idx, entry in enumerate(entries):
        if idx > 0:
            parts.append("\r")
            cursor_1based += 1
        for span in entry.full_spans():
            norm_text = span.text.replace("\r\n", "\r").replace("\n", "\r")
            if not norm_text:
                continue
            span_len = len(norm_text)
            parts.append(norm_text)
            ranges.append(
                CharFormatRange(
                    start=cursor_1based,
                    length=span_len,
                    italic=bool(span.italic),
                    bold=bool(span.bold),
                )
            )
            cursor_1based += span_len

    return "".join(parts), ranges


def merge_citation_texts(existing_text: str, new_text: str) -> str:
    """Łączy dwa teksty cytowań: `[a]` + `[b]` -> `[a, b]`, `(A)` + `(B)` -> `(A; B)` (D13)."""
    e = existing_text.strip()
    n = new_text.strip()
    if not e:
        return n
    if not n:
        return e
    if e.startswith("[") and e.endswith("]") and n.startswith("[") and n.endswith("]"):
        a = e[1:-1].strip()
        b = n[1:-1].strip()
        return f"[{a}, {b}]"
    if e.startswith("(") and e.endswith(")") and n.startswith("(") and n.endswith(")"):
        a = e[1:-1].strip()
        b = n[1:-1].strip()
        return f"({a}; {b})"
    return f"{e} {n}"


def merge_citations(existing_cit: Citation, new_cit: Citation) -> Citation:
    """Scala istniejące cytowanie z nowym z zachowaniem `existing_cit.citation_id` (D13)."""
    merged_keys = tuple(dict.fromkeys((*existing_cit.keys, *new_cit.keys)))
    if new_cit.citation_id == existing_cit.citation_id:
        merged_text = new_cit.text
    elif merged_keys == existing_cit.keys and new_cit.text == existing_cit.text:
        merged_text = existing_cit.text
    else:
        merged_text = merge_citation_texts(existing_cit.text, new_cit.text)
    return Citation(
        citation_id=existing_cit.citation_id,
        keys=merged_keys,
        text=merged_text,
    )


def _locate_citation_spans_in_text(
    shape_text: str,
    citations: Sequence[Citation],
) -> list[tuple[int, Citation, int, int]]:
    """Znajduje niepokrywające się wystąpienia `(cit_idx, cit, start, end)` w tekście kształtu."""
    located: list[tuple[int, Citation, int, int]] = []
    used_intervals: list[tuple[int, int]] = []

    for idx, cit in enumerate(citations):
        if not cit.text:
            continue
        search_from = 0
        while search_from <= len(shape_text) - len(cit.text):
            pos = shape_text.find(cit.text, search_from)
            if pos == -1:
                break
            end = pos + len(cit.text)
            overlaps = any(not (end <= u_start or pos >= u_end) for u_start, u_end in used_intervals)
            if not overlaps:
                used_intervals.append((pos, end))
                located.append((idx, cit, pos, end))
                break
            search_from = pos + 1

    return located


def find_adjacent_citation_in_text(
    shape_text: str,
    char_pos: int,
    citations: Sequence[Citation],
) -> tuple[int, Citation, int, int] | None:
    """Sprawdza, czy kursor `char_pos` stoi bezpośrednio przy lub wewnątrz istniejącego cytowania.

    Dopuszcza spacje/tabulatory pomiędzy kursorem a granicą cytowania `existing_cit.text`.
    Zwraca `(cit_idx, existing_cit, start, end)` lub `None`.
    """
    pos = max(0, min(int(char_pos), len(shape_text)))
    located = _locate_citation_spans_in_text(shape_text, citations)
    best: tuple[int, int, tuple[int, Citation, int, int]] | None = None

    for item in located:
        _, _, start, end = item
        distance: int | None = None
        preference = 0  # 0 = cytowanie przed/pod kursorem, 1 = cytowanie za kursorem
        if start <= pos <= end:
            distance = 0
            preference = 0
        elif pos > end:
            between = shape_text[end:pos]
            if between.strip(" \t") == "":
                distance = pos - end
                preference = 0
        elif pos < start:
            between = shape_text[pos:start]
            if between.strip(" \t") == "":
                distance = start - pos
                preference = 1

        if distance is not None:
            candidate_key = (distance, preference)
            if best is None or candidate_key < (best[0], best[1]):
                best = (distance, preference, item)

    return best[2] if best is not None else None


def apply_text_edits(text: str, edits: Sequence[TextEdit]) -> str:
    """Stosuje rozłączne edycje `TextEdit` w kolejności od końca do początku tekstu (T10)."""
    result = text
    for edit in sorted(edits, key=lambda e: e.start, reverse=True):
        result = f"{result[:edit.start]}{edit.new_text}{result[edit.start + edit.length:]}"
    return result


def plan_citation_insert(
    shape_text: str,
    char_pos: int,
    citation: Citation,
    existing_citations: Sequence[Citation],
) -> tuple[TextEdit, list[Citation], int]:
    """Wyznacza pojedynczą edycję `TextEdit` dla wstawienia lub scalenia cytowania (D13, T10).

    Zwraca `(edit, updated_citations, new_char_pos)`.
    """
    pos = max(0, min(int(char_pos), len(shape_text)))
    adj = find_adjacent_citation_in_text(shape_text, pos, existing_citations)

    if adj is not None:
        cit_idx, existing_cit, start, end = adj
        merged = merge_citations(existing_cit, citation)
        edit = TextEdit(start=start, length=end - start, new_text=merged.text)
        new_citations = list(existing_citations)
        new_citations[cit_idx] = merged
        new_pos = start + len(merged.text)
        return edit, new_citations, new_pos

    needs_space = (
        pos > 0
        and not shape_text[pos - 1].isspace()
        and shape_text[pos - 1] not in "([{"
    )
    prefix = " " if needs_space else ""
    inserted = f"{prefix}{citation.text}"
    edit = TextEdit(start=pos, length=0, new_text=inserted)

    located = _locate_citation_spans_in_text(shape_text, existing_citations)
    start_by_idx = {cit_idx: start for cit_idx, _, start, _ in located}
    insert_list_idx = len(existing_citations)
    for cit_idx in range(len(existing_citations)):
        cit_start = start_by_idx.get(cit_idx)
        if cit_start is not None and cit_start >= pos:
            insert_list_idx = cit_idx
            break

    new_citations = list(existing_citations)
    new_citations.insert(insert_list_idx, citation)
    new_pos = pos + len(inserted)
    return edit, new_citations, new_pos


def apply_citation_at_cursor_in_text(
    shape_text: str,
    char_pos: int,
    citation: Citation,
    existing_citations: Sequence[Citation],
) -> tuple[str, list[Citation], int]:
    """Wstawia cytowanie na pozycji `char_pos` lub scala je z sąsiednim cytowaniem (D13).

    Zwraca `(new_shape_text, updated_citations, new_char_pos)`.
    """
    edit, new_citations, new_pos = plan_citation_insert(
        shape_text=shape_text,
        char_pos=char_pos,
        citation=citation,
        existing_citations=existing_citations,
    )
    return apply_text_edits(shape_text, [edit]), new_citations, new_pos


def plan_citation_replacements(
    text: str,
    old_citations: Sequence[Citation],
    new_citations: Sequence[Citation],
) -> list[TextEdit]:
    """Wyznacza listę edycji `TextEdit` (posortowaną malejąco po `start`) dla podmiany cytowań (T10).

    Lokalizuje niepokrywające się wystąpienia starych cytowań w kolejności `old_citations`
    i zwraca tylko te zakresy, których tekst faktycznie się zmienia.
    """
    new_by_id = {c.citation_id: c.text for c in new_citations}
    used_intervals: list[tuple[int, int]] = []
    edits: list[TextEdit] = []

    for old_cit in old_citations:
        if old_cit.citation_id not in new_by_id or not old_cit.text:
            continue
        target_new_text = new_by_id[old_cit.citation_id]
        search_from = 0
        while search_from <= len(text) - len(old_cit.text):
            pos = text.find(old_cit.text, search_from)
            if pos == -1:
                break
            end = pos + len(old_cit.text)
            overlaps = any(
                not (end <= u_start or pos >= u_end)
                for u_start, u_end in used_intervals
            )
            if not overlaps:
                used_intervals.append((pos, end))
                if target_new_text != old_cit.text:
                    edits.append(
                        TextEdit(
                            start=pos,
                            length=len(old_cit.text),
                            new_text=target_new_text,
                        )
                    )
                break
            search_from = pos + 1

    edits.sort(key=lambda e: e.start, reverse=True)
    return edits


def _apply_text_edit_to_text_range(tr: Any, edit: TextEdit) -> None:
    """Stosuje pojedynczą edycję `TextEdit` na obiekcie `TextRange` przez podzakres `Characters` (T10)."""
    if hasattr(tr, "Characters"):
        if edit.length == 0:
            sub = tr.Characters(edit.start + 1, 0)
            if hasattr(sub, "InsertAfter"):
                sub.InsertAfter(edit.new_text)
            else:
                sub.Text = edit.new_text
        else:
            tr.Characters(edit.start + 1, edit.length).Text = edit.new_text
    else:
        current = str(tr.Text)
        tr.Text = f"{current[:edit.start]}{edit.new_text}{current[edit.start + edit.length:]}"


def choose_powerpoint_window(
    windows: Sequence[tuple[int, str]],
    presentation_name: str = "",
) -> int | None:
    """Wybiera uchwyt okna PowerPointa (`hwnd`) pasujący do nazwy prezentacji (D20)."""
    filtered = [
        (hwnd, title)
        for hwnd, title in windows
        if "asystent zotero dla powerpointa" not in (title or "").casefold()
    ]
    if not filtered:
        return None

    titled = [(hwnd, title) for hwnd, title in filtered if title and title.strip()]
    clean_name = presentation_name.strip()
    if clean_name:
        target_full = clean_name.casefold()
        target_stem = Path(clean_name).stem.casefold()
        for hwnd, title in titled:
            if target_full in title.casefold():
                return hwnd
        if target_stem:
            for hwnd, title in titled:
                if target_stem in title.casefold():
                    return hwnd

    if titled:
        return titled[0][0]
    return filtered[0][0]


def replace_citations_two_pass(
    texts: Sequence[str],
    old_citations: Sequence[Citation],
    new_citations: Sequence[Citation],
) -> list[str]:
    """Dwuprzebiegowa podmiana tekstów cytowań w liście tekstów kształtów slajdu.

    Przebieg 1 podmienia stare teksty cytowań w kolejności wystąpień na unikalne
    znaczniki tymczasowe. Przebieg 2 podstawia nowe teksty, dzięki czemu duplikaty
    (`[1]`, `[1]`) oraz zamiany miejscami (`[1] <-> [2]`) nigdy się nie nadpisują.
    """
    new_by_id = {c.citation_id: c.text for c in new_citations}
    working = list(texts)
    tokens: list[tuple[str, str]] = []

    token_prefix = "\ufffdAZP_CIT_"
    while any(token_prefix in t for t in working):
        token_prefix = f"\ufffd{token_prefix}"

    for idx, old_cit in enumerate(old_citations):
        if old_cit.citation_id not in new_by_id or not old_cit.text:
            continue
        token = f"{token_prefix}{idx}_{old_cit.citation_id}\ufffd"
        for s_idx, current_text in enumerate(working):
            if old_cit.text in current_text:
                working[s_idx] = current_text.replace(old_cit.text, token, 1)
                tokens.append((token, new_by_id[old_cit.citation_id]))
                break

    for token, new_text in tokens:
        for s_idx, current_text in enumerate(working):
            if token in current_text:
                working[s_idx] = current_text.replace(token, new_text, 1)
                break

    return working


def replace_citations_in_text(
    text: str,
    old_citations: Sequence[Citation],
    new_citations: Sequence[Citation],
) -> str:
    """Dwuprzebiegowa podmiana tekstów cytowań w pojedynczym łańcuchu tekstu."""
    return replace_citations_two_pass([text], old_citations, new_citations)[0]


def serialize_citations(citations: Sequence[Citation]) -> str:
    """Serializuje listę `Citation` do JSON-a zapisywanego w tagu `AZP_CITATIONS`."""
    payload = [
        {
            "citation_id": c.citation_id,
            "keys": list(c.keys),
            "text": c.text,
        }
        for c in citations
    ]
    return json.dumps(payload, ensure_ascii=False)


def deserialize_citations(raw: str) -> tuple[Citation, ...]:
    """Odtwarza krotkę `Citation` z JSON-a w tagu `AZP_CITATIONS`."""
    if not raw or not raw.strip():
        return ()
    data = json.loads(raw)
    if not isinstance(data, list):
        return ()
    result: list[Citation] = []
    for item in data:
        if isinstance(item, dict) and item.get("keys"):
            result.append(
                Citation(
                    citation_id=str(item["citation_id"]),
                    keys=tuple(str(k) for k in item["keys"]),
                    text=str(item.get("text", "")),
                )
            )
    return tuple(result)


def serialize_footer_entries(entries: Sequence[FooterEntry]) -> str:
    """Serializuje listę `FooterEntry` do JSON-a zapisywanego w tagu `AZP_FOOTER_ENTRIES`."""
    payload = [
        {
            "key": e.key,
            "prefix": e.prefix,
            "spans": [
                {
                    "text": s.text,
                    "italic": bool(s.italic),
                    "bold": bool(s.bold),
                }
                for s in e.spans
            ],
        }
        for e in entries
    ]
    return json.dumps(payload, ensure_ascii=False)


def deserialize_footer_entries(raw: str) -> tuple[FooterEntry, ...]:
    """Odtwarza krotkę `FooterEntry` z JSON-a w tagu `AZP_FOOTER_ENTRIES`."""
    if not raw or not raw.strip():
        return ()
    data = json.loads(raw)
    if not isinstance(data, list):
        return ()
    result: list[FooterEntry] = []
    for item in data:
        if not isinstance(item, dict) or "key" not in item:
            continue
        raw_spans = item.get("spans", [])
        spans = tuple(
            TextSpan(
                text=str(sp.get("text", "")),
                italic=bool(sp.get("italic", False)),
                bold=bool(sp.get("bold", False)),
            )
            for sp in raw_spans
            if isinstance(sp, dict)
        )
        result.append(
            FooterEntry(
                key=str(item["key"]),
                prefix=str(item.get("prefix", "")),
                spans=spans,
            )
        )
    return tuple(result)


def serialize_settings(settings: Settings) -> str:
    """Serializuje `Settings` do JSON-a zapisywanego w tagu `AZP_SETTINGS`."""
    return json.dumps(settings.to_dict(), ensure_ascii=False)


def deserialize_settings(raw: str) -> Settings | None:
    """Odtwarza `Settings` z JSON-a w tagu `AZP_SETTINGS` (lub `None`, jeśli pusty)."""
    if not raw or not raw.strip():
        return None
    data = json.loads(raw)
    if not isinstance(data, dict):
        return None
    return Settings.from_dict(data)


def _get_tag(com_obj: Any, key: str) -> str:
    """Bezpiecznie odczytuje wartość tagu z obiektu COM PowerPointa."""
    try:
        val = com_obj.Tags(key)
        return str(val) if val else ""
    except Exception:
        return ""


def _set_tag(com_obj: Any, key: str, value: str) -> None:
    """Bezpiecznie ustawia (lub usuwa, gdy `value` puste) tag w obiekcie COM PowerPointa."""
    try:
        com_obj.Tags.Delete(key)
    except Exception:
        pass
    if value:
        com_obj.Tags.Add(key, value)


def _shape_has_text_frame(shape: Any) -> bool:
    """Zwraca `True`, jeśli kształt COM posiada dostępny `TextFrame.TextRange`."""
    try:
        if not shape.HasTextFrame:
            return False
        _ = shape.TextFrame.TextRange
        return True
    except Exception:
        return False


class PowerPointSlideDoc:
    """Implementacja protokołu `SlideDoc` komunikująca się z PowerPointem przez COM."""

    def __init__(
        self,
        app: Any | None = None,
        presentation: Any | None = None,
        active_slide_index: int | None = None,
    ) -> None:
        self._app = app
        self._presentation = presentation
        self._active_slide_index = active_slide_index
        self._cursor_target: _CursorTarget | None = None
        self._last_insert_cursor: tuple[int, tuple[int | None, str | None], int] | None = None

    def set_active_slide(self, slide_index: int | None) -> None:
        """Pomocnicza metoda ustawiająca aktywny slajd w trybie bezokienkowym (`WithWindow=0`)."""
        self._active_slide_index = slide_index

    def set_cursor_target(
        self,
        slide_index: int | None,
        shape_index_or_name: int | str | None = None,
        char_pos: int | None = None,
    ) -> None:
        """Ustawia lub czyści (`slide_index=None`) programowy cel kursora dla trybu `WithWindow=0`."""
        self._last_insert_cursor = None
        if slide_index is None:
            self._cursor_target = None
            app = self._get_app()
            if app is not None:
                try:
                    app.ActiveWindow.Selection.Unselect()
                except Exception:
                    pass
            return

        self._active_slide_index = int(slide_index)
        self._cursor_target = _CursorTarget(
            slide_index=int(slide_index),
            shape_index_or_name=shape_index_or_name,
            char_pos=char_pos,
        )

    def clear_cursor_target(self) -> None:
        """Czyści programowy cel kursora (odpowiednik `set_cursor_target(None)`)."""
        self.set_cursor_target(None)

    def _get_app(self) -> Any | None:
        if self._app is not None:
            return self._app
        if self._presentation is not None:
            try:
                return self._presentation.Application
            except Exception:
                return None
        try:
            import win32com.client

            return win32com.client.GetActiveObject("PowerPoint.Application")
        except Exception:
            return None

    def _get_presentation(self) -> Any | None:
        if self._presentation is not None:
            try:
                _ = self._presentation.Slides.Count
                return self._presentation
            except Exception:
                return None

        app = self._get_app()
        if app is None:
            return None

        try:
            if int(app.Presentations.Count) <= 0:
                return None
        except Exception:
            return None

        try:
            return app.ActivePresentation
        except Exception:
            try:
                return app.Presentations(1)
            except Exception:
                return None

    def _get_slide_com(self, slide_index: int) -> Any:
        pres = self._get_presentation()
        if pres is None:
            raise NoActiveSlideError("Brak otwartej prezentacji w PowerPoincie.")
        count = int(pres.Slides.Count)
        if not (1 <= slide_index <= count):
            raise IndexError(f"Indeks slajdu poza zakresem 1..{count}: {slide_index}")
        return pres.Slides(slide_index)

    def _find_footer_shape(self, slide: Any) -> Any | None:
        count = int(slide.Shapes.Count)
        for i in range(1, count + 1):
            shape = slide.Shapes(i)
            if _get_tag(shape, TAG_ROLE) == TAG_ROLE_FOOTER:
                return shape
        return None

    def _list_body_text_shapes(self, slide: Any) -> list[Any]:
        shapes: list[Any] = []
        count = int(slide.Shapes.Count)
        for i in range(1, count + 1):
            shape = slide.Shapes(i)
            if _get_tag(shape, TAG_ROLE) == TAG_ROLE_FOOTER:
                continue
            if _shape_has_text_frame(shape):
                shapes.append(shape)
        return shapes

    def _resolve_cursor_context(self) -> tuple[int, Any, Any, str, int] | None:
        """Zwraca `(slide_index, slide, shape, shape_text, char_pos)` lub `None`, gdy brak kursora."""
        pres = self._get_presentation()
        if pres is None:
            return None
        slide_count = int(pres.Slides.Count)
        if slide_count <= 0:
            return None

        if self._cursor_target is not None:
            t_slide_idx = self._cursor_target.slide_index
            if not (1 <= t_slide_idx <= slide_count):
                return None
            slide = pres.Slides(t_slide_idx)
            target_shape: Any | None = None
            spec = self._cursor_target.shape_index_or_name

            if spec is None:
                body_shapes = self._list_body_text_shapes(slide)
                if body_shapes:
                    target_shape = body_shapes[0]
            elif isinstance(spec, int):
                if 1 <= spec <= int(slide.Shapes.Count):
                    cand = slide.Shapes(spec)
                    if (
                        _get_tag(cand, TAG_ROLE) != TAG_ROLE_FOOTER
                        and _shape_has_text_frame(cand)
                    ):
                        target_shape = cand
            elif isinstance(spec, str):
                try:
                    cand = slide.Shapes(spec)
                    if (
                        _get_tag(cand, TAG_ROLE) != TAG_ROLE_FOOTER
                        and _shape_has_text_frame(cand)
                    ):
                        target_shape = cand
                except Exception:
                    target_shape = None

            if target_shape is None:
                return None

            shape_text = str(target_shape.TextFrame.TextRange.Text)
            if self._cursor_target.char_pos is None:
                char_pos = len(shape_text)
            else:
                char_pos = max(0, min(int(self._cursor_target.char_pos), len(shape_text)))
            return t_slide_idx, slide, target_shape, shape_text, char_pos

        # Tryb z oknem PowerPointa: odczyt z ActiveWindow.Selection
        app = self._get_app()
        if app is None:
            return None
        try:
            win = app.ActiveWindow
            sel = win.Selection
            if int(sel.Type) != PP_SELECTION_TEXT:
                return None
            shape = sel.ShapeRange(1)
            if _get_tag(shape, TAG_ROLE) == TAG_ROLE_FOOTER or not _shape_has_text_frame(shape):
                return None
            try:
                slide_idx = int(sel.SlideRange(1).SlideIndex)
            except Exception:
                slide_idx = int(win.View.Slide.SlideIndex)
            if not (1 <= slide_idx <= slide_count):
                return None
            slide = pres.Slides(slide_idx)
            shape_text = str(shape.TextFrame.TextRange.Text)
            try:
                start_1based = int(sel.TextRange.Start)
                char_pos = max(0, min(start_1based - 1, len(shape_text)))
            except Exception:
                char_pos = len(shape_text)
            return slide_idx, slide, shape, shape_text, char_pos
        except Exception:
            return None

    def has_open_presentation(self) -> bool:
        """Zwraca `True`, jeśli w PowerPoincie jest otwarta prezentacja."""
        return self._get_presentation() is not None

    def get_presentation_path(self) -> Path | None:
        """Zwraca ścieżkę do pliku `.pptx` lub `None`, jeśli plik nie był zapisany."""
        pres = self._get_presentation()
        if pres is None:
            return None
        try:
            pres_dir = str(pres.Path).strip()
            if not pres_dir:
                return None
            full_name = str(pres.FullName).strip()
            return Path(full_name) if full_name else None
        except Exception:
            return None

    def get_presentation_key(self) -> str | None:
        """Zwraca identyfikator aktywnej prezentacji (`FullName` lub `Name`) albo `None` (D19)."""
        pres = self._get_presentation()
        if pres is None:
            return None
        for attr in ("FullName", "Name"):
            try:
                val = str(getattr(pres, attr)).strip()
                if val:
                    return val
            except Exception:
                continue
        path = self.get_presentation_path()
        return str(path) if path is not None else None

    def focus_presentation_window(self) -> bool:
        """Przywraca fokus do okna PowerPointa po wstawieniu cytowania (D20)."""
        if self._cursor_target is not None:
            return False
        app = self._get_app()
        if app is not None:
            try:
                if int(app.Visible) == MSO_FALSE:
                    return False
            except Exception:
                pass
        pres = self._get_presentation()
        pres_name = ""
        if pres is not None:
            for attr in ("Name", "FullName"):
                try:
                    val = str(getattr(pres, attr)).strip()
                    if val:
                        pres_name = Path(val).name or val
                        break
                except Exception:
                    continue

        if app is not None:
            try:
                app.ActiveWindow.Activate()
            except Exception:
                pass

        try:
            import win32con
            import win32gui

            ppt_frame_candidates: list[tuple[int, str]] = []
            title_candidates: list[tuple[int, str]] = []

            def _enum_cb(hwnd: int, _extra: Any) -> None:
                if not win32gui.IsWindowVisible(hwnd):
                    return
                cls_name = win32gui.GetClassName(hwnd)
                title = win32gui.GetWindowText(hwnd)
                if cls_name == "TkTopLevel" or "asystent zotero dla powerpointa" in title.casefold():
                    return
                if cls_name == "PPTFrameClass":
                    ppt_frame_candidates.append((hwnd, title))
                elif "powerpoint" in title.casefold():
                    title_candidates.append((hwnd, title))

            win32gui.EnumWindows(_enum_cb, None)
            candidates = ppt_frame_candidates or title_candidates
            target_hwnd = choose_powerpoint_window(candidates, pres_name)
            if target_hwnd is not None:
                if win32gui.IsIconic(target_hwnd):
                    win32gui.ShowWindow(target_hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(target_hwnd)
                return True
        except Exception:
            pass
        return False

    def has_text_cursor(self) -> bool:
        """Zwraca `True`, jeśli kursor stoi w polu tekstowym na aktywnym slajdzie (D10)."""
        return self._resolve_cursor_context() is not None

    def get_active_slide_index(self) -> int:
        """Zwraca 1-based indeks aktywnego slajdu lub rzuca `NoActiveSlideError`."""
        pres = self._get_presentation()
        if pres is None:
            raise NoActiveSlideError("Brak otwartej prezentacji w PowerPoincie.")
        slide_count = int(pres.Slides.Count)
        if slide_count <= 0:
            raise NoActiveSlideError("Otwarta prezentacja nie zawiera żadnych slajdów.")

        if self._cursor_target is not None:
            idx = self._cursor_target.slide_index
            if 1 <= idx <= slide_count:
                return idx
            raise NoActiveSlideError(f"Nieprawidłowy indeks aktywnego slajdu: {idx}")

        if self._active_slide_index is not None:
            if 1 <= self._active_slide_index <= slide_count:
                return self._active_slide_index
            raise NoActiveSlideError(
                f"Nieprawidłowy indeks aktywnego slajdu: {self._active_slide_index}"
            )

        app = self._get_app()
        if app is not None:
            try:
                win = app.ActiveWindow
                try:
                    idx = int(win.Selection.SlideRange(1).SlideIndex)
                except Exception:
                    idx = int(win.View.Slide.SlideIndex)
                if 1 <= idx <= slide_count:
                    return idx
            except Exception:
                pass

        raise NoActiveSlideError("Brak aktywnego slajdu w PowerPoincie.")

    @staticmethod
    def _shape_key(shape: Any) -> tuple[int | None, str | None]:
        shape_id: int | None = None
        shape_name: str | None = None
        try:
            if hasattr(shape, "Id"):
                shape_id = int(shape.Id)
        except Exception:
            pass
        try:
            if hasattr(shape, "Name"):
                shape_name = str(shape.Name)
        except Exception:
            pass
        return (shape_id, shape_name)

    @staticmethod
    def _same_shape(a: Any, b: Any) -> bool:
        if a is b:
            return True
        try:
            if hasattr(a, "Id") and hasattr(b, "Id"):
                return int(a.Id) == int(b.Id)
        except Exception:
            pass
        try:
            if hasattr(a, "Name") and hasattr(b, "Name"):
                return str(a.Name) == str(b.Name)
        except Exception:
            pass
        return False

    def _get_citations_by_shape(self, slide: Any) -> list[tuple[Any, list[Citation]]]:
        """Zwraca listę `(shape, citations_list)` dla wszystkich kształtów tekstowych slajdu.

        Uwzględnia tag `AZP_CITATIONS` na poszczególnych kształtach oraz fallback
        z tagu slajdu dla kształtów jeszcze nieotagowanych.
        """
        body_shapes = self._list_body_text_shapes(slide)
        if not body_shapes:
            return []

        shape_cits: list[list[Citation]] = [
            list(deserialize_citations(_get_tag(sh, TAG_CITATIONS)))
            for sh in body_shapes
        ]
        assigned_ids = {c.citation_id for cits in shape_cits for c in cits}
        slide_cits = deserialize_citations(_get_tag(slide, TAG_CITATIONS))
        unassigned = [c for c in slide_cits if c.citation_id not in assigned_ids]

        if unassigned:
            working_texts = [str(sh.TextFrame.TextRange.Text) for sh in body_shapes]
            for s_idx, cits in enumerate(shape_cits):
                for existing in cits:
                    if existing.text and existing.text in working_texts[s_idx]:
                        working_texts[s_idx] = working_texts[s_idx].replace(
                            existing.text, "\ufffd", 1
                        )
            for cit in unassigned:
                placed = False
                if cit.text:
                    for s_idx, txt in enumerate(working_texts):
                        if cit.text in txt:
                            working_texts[s_idx] = txt.replace(cit.text, "\ufffd", 1)
                            shape_cits[s_idx].append(cit)
                            placed = True
                            break
                if not placed:
                    shape_cits[0].append(cit)

        return list(zip(body_shapes, shape_cits))

    def _sync_slide_citations_from_shapes(
        self,
        slide: Any,
        by_shape: Sequence[tuple[Any, Sequence[Citation]]],
    ) -> list[Citation]:
        """Zapisuje `AZP_CITATIONS` na kształtach i ich konkatenację (w kolejności kształtów) na slajdzie."""
        all_citations: list[Citation] = []
        for sh, cits in by_shape:
            _set_tag(sh, TAG_CITATIONS, serialize_citations(cits))
            all_citations.extend(cits)
        _set_tag(slide, TAG_CITATIONS, serialize_citations(all_citations))
        return all_citations

    def get_adjacent_citation_at_cursor(self) -> Citation | None:
        """Zwraca istniejące cytowanie sąsiadujące z kursorem (D13) lub `None`."""
        ctx = self._resolve_cursor_context()
        if ctx is None:
            return None
        _, slide, shape, shape_text, char_pos = ctx
        by_shape = self._get_citations_by_shape(slide)
        shape_citations: Sequence[Citation] = ()
        for sh, cits in by_shape:
            if self._same_shape(sh, shape):
                shape_citations = cits
                break
        adj = find_adjacent_citation_in_text(shape_text, char_pos, shape_citations)
        return adj[1] if adj is not None else None

    def insert_citation_at_cursor(self, citation: Citation) -> None:
        """Wstawia cytowanie w miejscu kursora (lub scala z sąsiednim – D13) i zapisuje tagi.

        Gdy kursor nie stoi w polu tekstowym, rzuca `NoTextCursorError` i nie wprowadza
        żadnych zmian w prezentacji (D10).
        """
        if not self.has_open_presentation():
            raise NoActiveSlideError("Brak otwartej prezentacji w PowerPoincie.")

        ctx = self._resolve_cursor_context()
        if ctx is None:
            raise NoTextCursorError(
                "Ustaw kursor w polu tekstowym na slajdzie przed wstawieniem cytowania."
            )

        slide_idx, slide, shape, shape_text, char_pos = ctx
        by_shape = self._get_citations_by_shape(slide)
        target_idx = 0
        shape_citations: list[Citation] = []
        for idx, (sh, cits) in enumerate(by_shape):
            if self._same_shape(sh, shape):
                target_idx = idx
                shape_citations = cits
                break

        edit, updated_shape_citations, new_char_pos = plan_citation_insert(
            shape_text=shape_text,
            char_pos=char_pos,
            citation=citation,
            existing_citations=shape_citations,
        )

        _apply_text_edit_to_text_range(shape.TextFrame.TextRange, edit)
        if by_shape:
            by_shape[target_idx] = (by_shape[target_idx][0], updated_shape_citations)
            self._sync_slide_citations_from_shapes(slide, by_shape)
        else:
            _set_tag(shape, TAG_CITATIONS, serialize_citations(updated_shape_citations))
            _set_tag(slide, TAG_CITATIONS, serialize_citations(updated_shape_citations))

        self._last_insert_cursor = (slide_idx, self._shape_key(shape), new_char_pos)
        if self._cursor_target is not None:
            self._cursor_target = _CursorTarget(
                slide_index=slide_idx,
                shape_index_or_name=self._cursor_target.shape_index_or_name,
                char_pos=new_char_pos,
            )
        else:
            try:
                shape.TextFrame.TextRange.Characters(new_char_pos + 1, 0).Select()
            except Exception:
                pass

    def _update_native_footer_placeholder_visibility(
        self,
        slide: Any,
        hide_native_footer: bool,
    ) -> None:
        """Ukrywa lub przywraca natywny placeholder stopki PowerPointa (`ppPlaceholderFooter = 15`)."""
        count = int(slide.Shapes.Count)
        for i in range(1, count + 1):
            shape = slide.Shapes(i)
            if _get_tag(shape, TAG_ROLE) == TAG_ROLE_FOOTER:
                continue
            try:
                if int(shape.PlaceholderFormat.Type) == PP_PLACEHOLDER_FOOTER:
                    if hide_native_footer:
                        if int(shape.Visible) != MSO_FALSE:
                            shape.Visible = MSO_FALSE
                            _set_tag(shape, TAG_HIDDEN_NATIVE_FOOTER, "1")
                    elif _get_tag(shape, TAG_HIDDEN_NATIVE_FOOTER) == "1":
                        shape.Visible = MSO_TRUE
                        _set_tag(shape, TAG_HIDDEN_NATIVE_FOOTER, "")
            except Exception:
                continue

    def set_slide_footer(
        self,
        slide_index: int,
        entries: Sequence[FooterEntry],
        settings: Settings,
    ) -> None:
        """Tworzy, aktualizuje lub usuwa (gdy `entries` puste) pole źródeł na slajdzie."""
        pres = self._get_presentation()
        if pres is None:
            raise NoActiveSlideError("Brak otwartej prezentacji w PowerPoincie.")
        slide = self._get_slide_com(slide_index)
        existing_shape = self._find_footer_shape(slide)

        if not entries:
            if existing_shape is not None:
                existing_shape.Delete()
            self._update_native_footer_placeholder_visibility(slide, hide_native_footer=False)
            return

        self._update_native_footer_placeholder_visibility(
            slide,
            hide_native_footer=(settings.footer_position == "replace_footer"),
        )

        slide_width = float(pres.PageSetup.SlideWidth)
        slide_height = float(pres.PageSetup.SlideHeight)
        geom = compute_footer_geometry(
            slide_width=slide_width,
            slide_height=slide_height,
            position=settings.footer_position,
            height=DEFAULT_FOOTER_HEIGHT_PT,
        )

        if existing_shape is None:
            shape = slide.Shapes.AddTextbox(
                MSO_TEXT_ORIENTATION_HORIZONTAL,
                geom.left,
                geom.top,
                geom.width,
                geom.height,
            )
            _set_tag(shape, TAG_ROLE, TAG_ROLE_FOOTER)
        else:
            shape = existing_shape
            shape.Left = geom.left
            shape.Width = geom.width
            shape.Top = geom.top

        full_text, ranges = build_footer_text_and_ranges(
            entries=entries,
            footer_header=settings.footer_header,
        )

        tf = shape.TextFrame
        tf.WordWrap = MSO_TRUE
        tf.MarginLeft = 4.0
        tf.MarginRight = 4.0
        tf.MarginTop = 2.0
        tf.MarginBottom = 2.0

        tr = tf.TextRange
        tr.Text = full_text
        tr.Font.Name = settings.font_name
        tr.Font.Size = settings.font_size_pt
        tr.Font.Italic = MSO_FALSE
        tr.Font.Bold = MSO_FALSE

        for char_range in ranges:
            if char_range.length > 0:
                sub = tr.Characters(char_range.start, char_range.length)
                sub.Font.Italic = MSO_TRUE if char_range.italic else MSO_FALSE
                sub.Font.Bold = MSO_TRUE if char_range.bold else MSO_FALSE

        shape.Left = geom.left
        shape.Width = geom.width
        tf.AutoSize = PP_AUTOSIZE_SHAPE_TO_FIT_TEXT
        actual_height = float(shape.Height) if float(shape.Height) > 0 else DEFAULT_FOOTER_HEIGHT_PT
        adjusted_geom = compute_footer_geometry(
            slide_width=slide_width,
            slide_height=slide_height,
            position=settings.footer_position,
            height=actual_height,
        )
        shape.Top = adjusted_geom.top

        _set_tag(shape, TAG_FOOTER_ENTRIES, serialize_footer_entries(entries))

    def read_slides(self) -> list[SlideRecord]:
        """Odczytuje wszystkie slajdy w kolejności od 1 do N wraz z ich cytowaniami i stopkami."""
        pres = self._get_presentation()
        if pres is None:
            return []

        records: list[SlideRecord] = []
        slide_count = int(pres.Slides.Count)
        for idx in range(1, slide_count + 1):
            slide = pres.Slides(idx)
            slide_id = int(slide.SlideID)
            by_shape = self._get_citations_by_shape(slide)
            if by_shape:
                citations = tuple(c for _, cits in by_shape for c in cits)
            else:
                citations = deserialize_citations(_get_tag(slide, TAG_CITATIONS))
            footer_shape = self._find_footer_shape(slide)
            if footer_shape is not None:
                footer_entries = deserialize_footer_entries(
                    _get_tag(footer_shape, TAG_FOOTER_ENTRIES)
                )
            else:
                footer_entries = ()

            records.append(
                SlideRecord(
                    slide_index=idx,
                    slide_id=slide_id,
                    citations=citations,
                    footer_entries=footer_entries,
                )
            )
        return records

    def update_slide_citations(
        self,
        slide_index: int,
        citations: Sequence[Citation],
    ) -> None:
        """Aktualizuje widoczne teksty cytowań (dwuprzebiegowo) oraz tagi na wskazanym slajdzie."""
        slide = self._get_slide_com(slide_index)
        by_shape = self._get_citations_by_shape(slide)
        new_by_id = {c.citation_id: c for c in citations}

        updated_by_shape: list[tuple[Any, list[Citation]]] = []
        for shape, old_shape_cits in by_shape:
            tr = shape.TextFrame.TextRange
            old_t = str(tr.Text)
            new_shape_cits = [
                new_by_id[c.citation_id]
                for c in old_shape_cits
                if c.citation_id in new_by_id
            ]
            edits = plan_citation_replacements(
                text=old_t,
                old_citations=old_shape_cits,
                new_citations=new_shape_cits,
            )
            for edit in edits:
                _apply_text_edit_to_text_range(tr, edit)
            if (
                edits
                and self._last_insert_cursor is not None
                and self._last_insert_cursor[0] == slide_index
                and self._last_insert_cursor[1] == self._shape_key(shape)
            ):
                old_char_pos = self._last_insert_cursor[2]
                delta = 0
                for edit in edits:
                    if edit.start + edit.length <= old_char_pos:
                        delta += len(edit.new_text) - edit.length
                    elif edit.start < old_char_pos < edit.start + edit.length:
                        delta += (edit.start + len(edit.new_text)) - old_char_pos
                char_pos = max(0, min(old_char_pos + delta, len(str(tr.Text))))
                self._last_insert_cursor = (slide_index, self._shape_key(shape), char_pos)
                if (
                    self._cursor_target is not None
                    and self._cursor_target.slide_index == slide_index
                ):
                    self._cursor_target = _CursorTarget(
                        slide_index=slide_index,
                        shape_index_or_name=self._cursor_target.shape_index_or_name,
                        char_pos=char_pos,
                    )
                else:
                    try:
                        tr.Characters(char_pos + 1, 0).Select()
                    except Exception:
                        pass
            updated_by_shape.append((shape, new_shape_cits))

        if updated_by_shape:
            for sh, cits in updated_by_shape:
                _set_tag(sh, TAG_CITATIONS, serialize_citations(cits))
        _set_tag(slide, TAG_CITATIONS, serialize_citations(citations))

    def load_settings(self) -> Settings | None:
        """Odczytuje ustawienia AZP zapisane w tagach prezentacji (`None` jeśli brak)."""
        pres = self._get_presentation()
        if pres is None:
            return None
        return deserialize_settings(_get_tag(pres, TAG_SETTINGS))

    def save_settings(self, settings: Settings) -> None:
        """Zapisuje ustawienia AZP w tagach aktywnej prezentacji."""
        pres = self._get_presentation()
        if pres is None:
            raise NoActiveSlideError("Brak otwartej prezentacji w PowerPoincie.")
        _set_tag(pres, TAG_SETTINGS, serialize_settings(settings))

    def create_backup(self) -> Path | None:
        """Tworzy kopię zapasową `<nazwa>.bak.<rozsz>` przez `SaveCopyAs` bez zapisu oryginału (D17)."""
        pres = self._get_presentation()
        if pres is None:
            return None
        path = self.get_presentation_path()
        if path is None:
            return None
        suffix = path.suffix or ".pptx"
        backup_path = path.with_suffix(f".bak{suffix}")
        try:
            if hasattr(pres, "SaveCopyAs"):
                pres.SaveCopyAs(str(backup_path))
            else:
                shutil.copy2(path, backup_path)
        except Exception as exc:
            raise BackupError(
                f"Nie udało się utworzyć kopii zapasowej ({backup_path.name}): {exc}"
            ) from exc
        return backup_path


__all__ = [
    "CharFormatRange",
    "DEFAULT_FOOTER_HEIGHT_PT",
    "FOOTER_HEADER_TEXT",
    "FooterGeometry",
    "PP_AUTOSIZE_SHAPE_TO_FIT_TEXT",
    "PP_PLACEHOLDER_FOOTER",
    "PP_SELECTION_TEXT",
    "PowerPointSlideDoc",
    "SlideDoc",
    "TAG_CITATIONS",
    "TAG_FOOTER_ENTRIES",
    "TAG_ROLE",
    "TAG_ROLE_FOOTER",
    "TAG_SETTINGS",
    "TextEdit",
    "apply_citation_at_cursor_in_text",
    "apply_text_edits",
    "build_footer_text_and_ranges",
    "choose_powerpoint_window",
    "compute_footer_geometry",
    "deserialize_citations",
    "deserialize_footer_entries",
    "deserialize_settings",
    "find_adjacent_citation_in_text",
    "merge_citation_texts",
    "merge_citations",
    "plan_citation_insert",
    "plan_citation_replacements",
    "replace_citations_in_text",
    "replace_citations_two_pass",
    "serialize_citations",
    "serialize_footer_entries",
    "serialize_settings",
]
