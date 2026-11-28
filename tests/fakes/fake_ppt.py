"""Atrapa `SlideDoc` do testowania GUI (B1.2) i silnika numeracji (B2.2) bez COM."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Sequence

from azp.contracts import (
    Citation,
    FooterEntry,
    NoActiveSlideError,
    NoTextCursorError,
    Settings,
    SlideRecord,
)


@dataclass
class _MutableSlide:
    slide_id: int
    body_text: str = ""
    citations: list[Citation] = field(default_factory=list)
    footer_entries: list[FooterEntry] = field(default_factory=list)
    last_footer_settings: Settings | None = None
    bibliography_entries: list[FooterEntry] = field(default_factory=list)
    bibliography_columns: list[list[FooterEntry]] = field(default_factory=list)
    last_bibliography_settings: Settings | None = None
    is_bib_continuation: bool = False


def _adjust_slide_index_after_bib_change(
    idx: int,
    base_slide_index: int,
    removed_count: int,
    added_count: int,
) -> int:
    if idx > base_slide_index + removed_count:
        return idx + (added_count - removed_count)
    if base_slide_index < idx <= base_slide_index + removed_count:
        return min(idx, base_slide_index + added_count)
    return idx


def _split_entries_into_columns(
    entries: Sequence[FooterEntry],
    columns: int,
) -> list[list[FooterEntry]]:
    cols = max(1, int(columns))
    if not entries:
        return []
    total = len(entries)
    actual_cols = min(cols, total)
    base, rem = divmod(total, actual_cols)
    result: list[list[FooterEntry]] = []
    idx = 0
    for c in range(actual_cols):
        count = base + (1 if c < rem else 0)
        result.append(list(entries[idx : idx + count]))
        idx += count
    return result


class FakeSlideDoc:
    """Atrapa dokumentu PowerPoint implementująca protokół `SlideDoc`."""

    def __init__(
        self,
        slide_count: int = 3,
        active_slide_index: int | None = 1,
        has_cursor: bool = True,
        presentation_path: Path | None = Path("C:/fake/prezentacja_testowa.pptx"),
    ) -> None:
        self._slides: list[_MutableSlide] = [
            _MutableSlide(slide_id=100 + i, body_text=f"Tekst slajdu {i}")
            for i in range(1, slide_count + 1)
        ]
        self.active_slide_index: int | None = active_slide_index
        self.has_cursor: bool = has_cursor
        self.presentation_path: Path | None = presentation_path
        self._settings: Settings | None = None
        self.backups_created: list[Path] = []

    def has_open_presentation(self) -> bool:
        return len(self._slides) > 0 and self.active_slide_index is not None

    def get_presentation_path(self) -> Path | None:
        return self.presentation_path

    def has_text_cursor(self) -> bool:
        return self.has_open_presentation() and self.has_cursor

    def get_active_slide_index(self) -> int:
        if not self.has_open_presentation() or self.active_slide_index is None:
            raise NoActiveSlideError("Brak otwartej prezentacji lub aktywnego slajdu.")
        if not (1 <= self.active_slide_index <= len(self._slides)):
            raise NoActiveSlideError(
                f"Nieprawidłowy indeks aktywnego slajdu: {self.active_slide_index}"
            )
        return self.active_slide_index

    def _get_slide(self, slide_index: int) -> _MutableSlide:
        if not (1 <= slide_index <= len(self._slides)):
            raise IndexError(f"Slajd poza zakresem: {slide_index}")
        return self._slides[slide_index - 1]

    def get_slide_body_text(self, slide_index: int) -> str:
        """Pomocnicza metoda testowa zwracająca tekst główny slajdu."""
        return self._get_slide(slide_index).body_text

    def get_slide_footer_settings(self, slide_index: int) -> Settings | None:
        """Pomocnicza metoda testowa zwracająca ostatnie ustawienia użyte do stopki."""
        return self._get_slide(slide_index).last_footer_settings

    def get_slide_bibliography_entries(self, slide_index: int) -> list[FooterEntry]:
        """Pomocnicza metoda testowa zwracająca wpisy bibliografii na slajdzie."""
        return list(self._get_slide(slide_index).bibliography_entries)

    def get_slide_bibliography_keys(self, slide_index: int) -> tuple[str, ...]:
        """Zwraca klucze wszystkich pozycji bibliografii ze slajdu (wraz z jego slajdami kontynuacji)."""
        slide = self._get_slide(slide_index)
        while slide_index > 1 and slide.is_bib_continuation:
            slide_index -= 1
            slide = self._get_slide(slide_index)
        keys: list[str] = [e.key for e in slide.bibliography_entries if e.key]
        next_idx = slide_index + 1
        while next_idx <= len(self._slides) and self._slides[next_idx - 1].is_bib_continuation:
            keys.extend(e.key for e in self._slides[next_idx - 1].bibliography_entries if e.key)
            next_idx += 1
        return tuple(dict.fromkeys(keys))

    def get_slide_bibliography_columns(self, slide_index: int) -> list[list[FooterEntry]]:
        """Pomocnicza metoda testowa zwracająca podział bibliografii na kolumny na slajdzie."""
        return [list(col) for col in self._get_slide(slide_index).bibliography_columns]

    def get_slide_bibliography_settings(self, slide_index: int) -> Settings | None:
        """Pomocnicza metoda testowa zwracająca ostatnie ustawienia użyte do bibliografii."""
        return self._get_slide(slide_index).last_bibliography_settings

    def reorder_slides(self, new_order_1based: Sequence[int]) -> None:
        """Pomocnicza metoda testowa symulująca przestawienie kolejności slajdów."""
        if sorted(new_order_1based) != list(range(1, len(self._slides) + 1)):
            raise ValueError("Nowa kolejność musi być permutacją indeksów 1..N.")
        self._slides = [self._slides[i - 1] for i in new_order_1based]

    def insert_citation_at_cursor(self, citation: Citation) -> None:
        slide_idx = self.get_active_slide_index()
        if not self.has_cursor:
            raise NoTextCursorError(
                "Ustaw kursor w polu tekstowym na slajdzie przed wstawieniem cytowania."
            )
        slide = self._get_slide(slide_idx)
        sep = " " if slide.body_text and not slide.body_text.endswith(" ") else ""
        slide.body_text = f"{slide.body_text}{sep}{citation.text}"
        if citation.in_text:
            insert_pos = next(
                (i for i, c in enumerate(slide.citations) if not c.in_text),
                len(slide.citations),
            )
            slide.citations.insert(insert_pos, citation)
        else:
            slide.citations.append(citation)

    def add_footer_citation(self, citation: Citation) -> None:
        slide_idx = self.get_active_slide_index()
        slide = self._get_slide(slide_idx)
        footer_cit = Citation(
            citation_id=citation.citation_id,
            keys=citation.keys,
            text="",
            in_text=False,
        )
        for idx, existing in enumerate(slide.citations):
            if not existing.in_text:
                merged_keys = tuple(dict.fromkeys((*existing.keys, *footer_cit.keys)))
                slide.citations[idx] = Citation(
                    citation_id=existing.citation_id,
                    keys=merged_keys,
                    text="",
                    in_text=False,
                )
                return
        slide.citations.append(footer_cit)

    def set_slide_footer(
        self,
        slide_index: int,
        entries: Sequence[FooterEntry],
        settings: Settings,
    ) -> None:
        slide = self._get_slide(slide_index)
        slide.footer_entries = list(entries)
        slide.last_footer_settings = settings if entries else None

    def _remove_bib_continuations_after(self, slide_index: int) -> int:
        removed_count = 0
        while (
            slide_index < len(self._slides)
            and self._slides[slide_index].is_bib_continuation
        ):
            self._slides.pop(slide_index)
            removed_count += 1
        return removed_count

    def set_slide_bibliography(
        self,
        slide_index: int,
        entries: Sequence[FooterEntry],
        settings: Settings,
    ) -> int:
        slide = self._get_slide(slide_index)
        while slide_index > 1 and slide.is_bib_continuation:
            slide_index -= 1
            slide = self._get_slide(slide_index)
        slide.is_bib_continuation = False

        removed_count = self._remove_bib_continuations_after(slide_index)

        if not entries:
            slide.bibliography_entries = []
            slide.bibliography_columns = []
            slide.last_bibliography_settings = None
            if self.active_slide_index is not None:
                self.active_slide_index = _adjust_slide_index_after_bib_change(
                    self.active_slide_index,
                    base_slide_index=slide_index,
                    removed_count=removed_count,
                    added_count=0,
                )
            return 0

        if settings.bib_fit_mode == "split_slides":
            # Przybliżona pojemność na slajd w zależności od rozmiaru czcionki i liczby kolumn
            per_col_cap = max(1, int(240 // max(12, int(settings.bib_font_size_pt * 2.2))))
            per_slide_cap = max(1, per_col_cap * settings.bib_columns)
        else:
            per_slide_cap = len(entries)

        chunks = [
            list(entries[i : i + per_slide_cap])
            for i in range(0, len(entries), per_slide_cap)
        ]
        first_chunk = chunks[0]
        slide.bibliography_entries = first_chunk
        slide.bibliography_columns = _split_entries_into_columns(
            first_chunk, settings.bib_columns
        )
        slide.last_bibliography_settings = settings

        next_id = max((s.slide_id for s in self._slides), default=100) + 1
        for offset, chunk in enumerate(chunks[1:], start=1):
            cont_slide = _MutableSlide(
                slide_id=next_id + offset,
                body_text="",
                bibliography_entries=chunk,
                bibliography_columns=_split_entries_into_columns(
                    chunk, settings.bib_columns
                ),
                last_bibliography_settings=settings,
                is_bib_continuation=True,
            )
            self._slides.insert(slide_index - 1 + offset, cont_slide)

        added_count = max(0, len(chunks) - 1)
        if self.active_slide_index is not None:
            self.active_slide_index = _adjust_slide_index_after_bib_change(
                self.active_slide_index,
                base_slide_index=slide_index,
                removed_count=removed_count,
                added_count=added_count,
            )

        return len(chunks)

    def read_slides(self) -> list[SlideRecord]:
        return [
            SlideRecord(
                slide_index=idx,
                slide_id=s.slide_id,
                citations=tuple(s.citations),
                footer_entries=tuple(s.footer_entries),
                has_bibliography=bool(s.bibliography_entries) and not s.is_bib_continuation,
            )
            for idx, s in enumerate(self._slides, start=1)
        ]

    def update_slide_citations(
        self,
        slide_index: int,
        citations: Sequence[Citation],
    ) -> None:
        slide = self._get_slide(slide_index)
        new_by_id = {c.citation_id: c.text for c in citations if c.in_text}
        # Przebieg 1: zamiana starych tekstów w kolejności występowania na unikalne znaczniki
        text = slide.body_text
        for old_cit in slide.citations:
            if (
                old_cit.in_text
                and old_cit.citation_id in new_by_id
                and old_cit.text
                and old_cit.text in text
            ):
                token = f"\x00CIT:{old_cit.citation_id}\x00"
                text = text.replace(old_cit.text, token, 1)
        # Przebieg 2: podstawienie nowych tekstów w miejsce znaczników
        for cid, new_text in new_by_id.items():
            token = f"\x00CIT:{cid}\x00"
            text = text.replace(token, new_text, 1)
        slide.body_text = text
        slide.citations = list(citations)

    def load_settings(self) -> Settings | None:
        return self._settings

    def save_settings(self, settings: Settings) -> None:
        self._settings = settings

    def create_backup(self) -> Path | None:
        if self.presentation_path is None:
            return None
        backup_path = self.presentation_path.with_suffix(".bak.pptx")
        self.backups_created.append(backup_path)
        return backup_path
