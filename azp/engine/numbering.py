"""Czyste funkcje numeracji cytowań, scalania i układania stopek (Etap B2.2).

Moduł nie wykonuje żadnych operacji wejścia-wyjścia ani wywołań COM/HTTP.
"""

from __future__ import annotations

from typing import Mapping, Sequence

from azp.contracts import (
    STYLE_APA,
    STYLE_IEEE_V4,
    Citation,
    FooterEntry,
    FormattedRef,
    Settings,
    SlideRecord,
    StyleInfo,
)


def is_numeric_style(style_id: str, styles: Sequence[StyleInfo] = ()) -> bool:
    """Określa, czy wskazany styl cytowania jest numeryczny (np. IEEE v4 vs APA 7)."""
    for style in styles:
        if style.style_id == style_id:
            return bool(style.is_numeric)
    if style_id == STYLE_IEEE_V4:
        return True
    if style_id == STYLE_APA:
        return False
    return "ieee" in style_id.casefold()


def assign_global_numbers(slides: Sequence[SlideRecord]) -> dict[str, int]:
    """Przypisuje numery `1..K` unikalnym kluczom w kolejności pierwszego wystąpienia w prezentacji (D4)."""
    number_map: dict[str, int] = {}
    for slide in slides:
        for citation in slide.citations:
            for key in citation.keys:
                if key not in number_map:
                    number_map[key] = len(number_map) + 1
    return number_map


def assign_slide_numbers(citations: Sequence[Citation]) -> dict[str, int]:
    """Przypisuje numery `1..M` unikalnym kluczom na pojedynczym slajdzie od `[1]` (D4)."""
    number_map: dict[str, int] = {}
    for citation in citations:
        for key in citation.keys:
            if key not in number_map:
                number_map[key] = len(number_map) + 1
    return number_map


def merge_citation_keys(
    existing_keys: Sequence[str],
    new_keys: Sequence[str],
) -> tuple[str, ...]:
    """Scala sekwencje kluczy cytowania z zachowaniem kolejności i bez duplikatów (D13)."""
    merged = tuple(dict.fromkeys((*existing_keys, *new_keys)))
    if not merged:
        raise ValueError("Połączona lista kluczy cytowania nie może być pusta.")
    return merged


def _strip_outer_parens(in_text: str) -> str:
    """Usuwa jedną parę zewnętrznych nawiasów okrągłych `(...)` z cytowania autor-rok."""
    stripped = in_text.strip()
    if len(stripped) >= 2 and stripped.startswith("(") and stripped.endswith(")"):
        depth = 0
        matches_outer = True
        for idx, ch in enumerate(stripped):
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0 and idx < len(stripped) - 1:
                    matches_outer = False
                    break
        if matches_outer and depth == 0:
            return stripped[1:-1].strip()
    return stripped


def format_citation_text(
    keys: Sequence[str],
    refs: Mapping[str, FormattedRef],
    is_numeric: bool,
    number_map: Mapping[str, int] | None = None,
) -> str:
    """Formatuje tekst cytowania w slajdzie (pojedynczego lub zbiorczego/scalonego – D2, D3, D13).

    - Dla stylu numerycznego (IEEE v4): zwraca `"[1]"` lub posortowane rosnąco `"[1, 2]"`.
    - Dla stylu autor-rok (APA 7): zwraca `"(Kowalski, 2023)"` lub
      `"(Kowalski, 2023; Nowak & Wiśniewski, 2024)"`.
    """
    unique_keys = tuple(dict.fromkeys(keys))
    if not unique_keys:
        raise ValueError("Lista kluczy cytowania nie może być pusta.")

    if is_numeric:
        if number_map is None:
            raise ValueError("Styl numeryczny wymaga mapy numeracji (number_map).")
        numbers = sorted({number_map[key] for key in unique_keys})
        return f"[{', '.join(str(n) for n in numbers)}]"

    parts = [_strip_outer_parens(refs[key].in_text) for key in unique_keys]
    return f"({'; '.join(parts)})"


def build_slide_footer_entries(
    citations: Sequence[Citation],
    refs: Mapping[str, FormattedRef],
    is_numeric: bool,
    number_map: Mapping[str, int] | None = None,
) -> tuple[FooterEntry, ...]:
    """Buduje wpisy pola źródeł (stopki) dla pojedynczego slajdu (D2, D3, T4, T8).

    - Styl numeryczny (IEEE v4): wpisy posortowane rosnąco po numerze `[n]` z prefiksem `"[n] "`.
    - Styl autor-rok (APA 7): wpisy posortowane alfabetycznie po `effective_sort_key` bez prefiksu.
    """
    unique_keys = list(
        dict.fromkeys(key for citation in citations for key in citation.keys)
    )
    if not unique_keys:
        return ()

    if is_numeric:
        if number_map is None:
            raise ValueError("Styl numeryczny wymaga mapy numeracji (number_map).")
        ordered_keys = sorted(unique_keys, key=lambda key: number_map[key])
        return tuple(
            FooterEntry(
                key=key,
                prefix=f"[{number_map[key]}] ",
                spans=refs[key].bib_body_spans,
            )
            for key in ordered_keys
        )

    ordered_keys = sorted(
        unique_keys,
        key=lambda key: (
            refs[key].effective_sort_key,
            refs[key].bib_plain.casefold(),
            key,
        ),
    )
    return tuple(
        FooterEntry(
            key=key,
            prefix="",
            spans=refs[key].bib_body_spans,
        )
        for key in ordered_keys
    )


def compute_presentation_plan(
    slides: Sequence[SlideRecord],
    refs: Mapping[str, FormattedRef],
    settings: Settings,
    styles: Sequence[StyleInfo] = (),
) -> list[SlideRecord]:
    """Wylicza docelowy stan cytowań i stopek dla wszystkich slajdów prezentacji."""
    numeric = is_numeric_style(settings.style_id, styles)
    global_map: dict[str, int] | None = (
        assign_global_numbers(slides)
        if numeric and settings.numbering_mode == "global"
        else None
    )

    planned_slides: list[SlideRecord] = []
    for slide in slides:
        if numeric:
            slide_map: Mapping[str, int] | None = (
                global_map
                if global_map is not None
                else assign_slide_numbers(slide.citations)
            )
        else:
            slide_map = None

        planned_citations = tuple(
            Citation(
                citation_id=cit.citation_id,
                keys=tuple(dict.fromkeys(cit.keys)),
                text=(
                    format_citation_text(
                        cit.keys,
                        refs,
                        is_numeric=numeric,
                        number_map=slide_map,
                    )
                    if cit.in_text
                    else ""
                ),
                in_text=cit.in_text,
            )
            for cit in slide.citations
        )
        planned_footer = build_slide_footer_entries(
            planned_citations,
            refs,
            is_numeric=numeric,
            number_map=slide_map,
        )
        planned_slides.append(
            SlideRecord(
                slide_index=slide.slide_index,
                slide_id=slide.slide_id,
                citations=planned_citations,
                footer_entries=planned_footer,
                has_bibliography=slide.has_bibliography,
            )
        )

    return planned_slides


def build_presentation_bibliography_entries(
    slides: Sequence[SlideRecord],
    refs: Mapping[str, FormattedRef],
    settings: Settings,
    styles: Sequence[StyleInfo] = (),
) -> tuple[FooterEntry, ...]:
    """Buduje pełny spis wykorzystanych źródeł (bibliografię) z całej prezentacji (D27).

    - Zbiera wszystkie unikalne klucze cytowań ze wszystkich slajdów w kolejności pierwszego wystąpienia.
    - Dla stylu numerycznego (IEEE v4): używa globalnej numeracji `1..K` i sortuje rosnąco po numerze.
    - Dla stylu autor-rok (APA 7): sortuje alfabetycznie bez prefiksu `[n]`.
    """
    unique_keys = list(
        dict.fromkeys(
            key
            for slide in slides
            for citation in slide.citations
            for key in citation.keys
        )
    )
    if not unique_keys:
        return ()

    numeric = is_numeric_style(settings.style_id, styles)
    if numeric:
        global_map = assign_global_numbers(slides)
        ordered_keys = sorted(unique_keys, key=lambda key: global_map[key])
        return tuple(
            FooterEntry(
                key=key,
                prefix=f"[{global_map[key]}] ",
                spans=refs[key].bib_body_spans,
            )
            for key in ordered_keys
        )

    ordered_keys = sorted(
        unique_keys,
        key=lambda key: (
            refs[key].effective_sort_key,
            refs[key].bib_plain.casefold(),
            key,
        ),
    )
    return tuple(
        FooterEntry(
            key=key,
            prefix="",
            spans=refs[key].bib_body_spans,
        )
        for key in ordered_keys
    )
