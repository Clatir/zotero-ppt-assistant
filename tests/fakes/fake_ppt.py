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
        slide.citations.append(citation)

    def set_slide_footer(
        self,
        slide_index: int,
        entries: Sequence[FooterEntry],
        settings: Settings,
    ) -> None:
        slide = self._get_slide(slide_index)
        slide.footer_entries = list(entries)
        slide.last_footer_settings = settings if entries else None

    def read_slides(self) -> list[SlideRecord]:
        return [
            SlideRecord(
                slide_index=idx,
                slide_id=s.slide_id,
                citations=tuple(s.citations),
                footer_entries=tuple(s.footer_entries),
            )
            for idx, s in enumerate(self._slides, start=1)
        ]

    def update_slide_citations(
        self,
        slide_index: int,
        citations: Sequence[Citation],
    ) -> None:
        slide = self._get_slide(slide_index)
        new_by_id = {c.citation_id: c.text for c in citations}
        # Przebieg 1: zamiana starych tekstów w kolejności występowania na unikalne znaczniki
        text = slide.body_text
        for old_cit in slide.citations:
            if old_cit.citation_id in new_by_id and old_cit.text and old_cit.text in text:
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
