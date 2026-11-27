"""Wspólne typy danych, wyjątki i interfejsy (kontrakty) dla modułów AZP.

Stanowi granicę między torem B1 (`azp/zotero`, `azp/gui`) a torem B2 (`azp/ppt`, `azp/engine`).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal, Protocol, Sequence, runtime_checkable

STYLE_IEEE_V4 = "http://www.zotero.org/styles/ieee-praca-inzynierska-access-date-v4"
STYLE_APA = "http://www.zotero.org/styles/apa"

NumberingMode = Literal["global", "per_slide"]
FooterPosition = Literal["bottom_full", "bottom_left", "replace_footer"]

VALID_NUMBERING_MODES: tuple[NumberingMode, ...] = ("global", "per_slide")
VALID_FOOTER_POSITIONS: tuple[FooterPosition, ...] = (
    "bottom_full",
    "bottom_left",
    "replace_footer",
)


class AzpError(Exception):
    """Bazowy wyjątek aplikacyjny AZP z komunikatem dla użytkownika po polsku."""


class ZoteroOfflineError(AzpError):
    """Zotero jest wyłączone albo lokalne API nie odpowiada."""


class NoActiveSlideError(AzpError):
    """Brak otwartej prezentacji lub aktywnego slajdu w PowerPoincie."""


class NoTextCursorError(AzpError):
    """Kursor nie znajduje się w polu tekstowym na slajdzie (D10)."""


@dataclass(frozen=True)
class TextSpan:
    """Fragment tekstu z informacją o formatowaniu (T4: kursywa z CSL HTML)."""

    text: str
    italic: bool = False
    bold: bool = False


def spans_to_plain_text(spans: Sequence[TextSpan]) -> str:
    """Łączy sekwencję przebiegów `TextSpan` w czysty tekst."""
    return "".join(span.text for span in spans)


@dataclass(frozen=True)
class StyleInfo:
    """Opis stylu CSL dostępnego w Zotero (D2, D3)."""

    style_id: str
    title: str
    is_numeric: bool


@dataclass(frozen=True)
class ZoteroItem:
    """Pozycja bibliograficzna z biblioteki Zotero."""

    key: str
    title: str
    authors: tuple[str, ...] = ()
    year: str = ""
    item_type: str = "journalArticle"

    @property
    def authors_label(self) -> str:
        """Zwięzły opis autorów do wyświetlenia na liście wyników w okienku."""
        if not self.authors:
            return "(brak autora)"
        if len(self.authors) == 1:
            return self.authors[0]
        if len(self.authors) == 2:
            return f"{self.authors[0]}, {self.authors[1]}"
        return f"{self.authors[0]} i in."

    @property
    def display_label(self) -> str:
        """Jednowierszowy opis pozycji na liście wyszukiwania."""
        year_part = f" ({self.year})" if self.year else ""
        return f"{self.authors_label}{year_part} – {self.title}"


@dataclass(frozen=True)
class FormattedRef:
    """Sformatowane przez silnik CSL Zotero cytowanie i opis bibliograficzny (T2, T4).

    `bib_body_spans` przechowuje treść opisu bibliograficznego z zachowaną kursywą,
    bez prefiksu numeru `[1]` w stylach numerycznych, aby silnik numeracji mógł
    podstawić właściwy numer `[n]`.
    """

    key: str
    style_id: str
    in_text: str
    bib_body_spans: tuple[TextSpan, ...]
    sort_key: str = ""

    @property
    def bib_plain(self) -> str:
        """Czysty tekst opisu bibliograficznego (bez znaczników)."""
        return spans_to_plain_text(self.bib_body_spans)

    @property
    def effective_sort_key(self) -> str:
        """Klucz do sortowania alfabetycznego (np. w stylu APA)."""
        return (self.sort_key or self.bib_plain).casefold()


@dataclass(frozen=True)
class Citation:
    """Wystąpienie cytowania w tekście slajdu.

    Może obejmować jeden lub więcej kluczy Zotero (np. scalone `[1, 2]`
    albo `(Kowalski, 2023; Nowak, 2024)`).
    """

    citation_id: str
    keys: tuple[str, ...]
    text: str = ""

    def __post_init__(self) -> None:
        if not self.keys:
            raise ValueError("Citation.keys nie może być puste.")


@dataclass(frozen=True)
class FooterEntry:
    """Pojedynczy wpis w polu źródeł (stopce) pod slajdem."""

    key: str
    spans: tuple[TextSpan, ...]
    prefix: str = ""

    def full_spans(self) -> tuple[TextSpan, ...]:
        """Zwraca pełną listę `TextSpan` wraz z ewentualnym prefiksem (np. `'[1] '`)."""
        if self.prefix:
            return (TextSpan(self.prefix), *self.spans)
        return self.spans

    @property
    def plain_text(self) -> str:
        """Pełny czysty tekst wpisu stopki."""
        return spans_to_plain_text(self.full_spans())


@dataclass(frozen=True)
class SlideRecord:
    """Stan cytowań i stopki na pojedynczym slajdzie prezentacji."""

    slide_index: int
    slide_id: int
    citations: tuple[Citation, ...] = ()
    footer_entries: tuple[FooterEntry, ...] = ()


@dataclass(frozen=True)
class Settings:
    """Ustawienia prezentacji i okienka (D2–D5, D9)."""

    style_id: str = STYLE_IEEE_V4
    numbering_mode: NumberingMode = "global"
    font_name: str = "Calibri"
    font_size_pt: int = 10
    footer_position: FooterPosition = "bottom_full"
    footer_header: bool = True
    always_on_top: bool = True

    def __post_init__(self) -> None:
        if self.numbering_mode not in VALID_NUMBERING_MODES:
            raise ValueError(f"Nieprawidłowy tryb numeracji: {self.numbering_mode!r}")
        if self.footer_position not in VALID_FOOTER_POSITIONS:
            raise ValueError(f"Nieprawidłowe położenie stopki: {self.footer_position!r}")
        if self.font_size_pt < 6 or self.font_size_pt > 72:
            raise ValueError(f"Rozmiar czcionki poza zakresem 6..72 pt: {self.font_size_pt}")
        if not self.font_name.strip():
            raise ValueError("Nazwa czcionki nie może być pusta.")
        if not self.style_id.strip():
            raise ValueError("Identyfikator stylu nie może być pusty.")

    def to_dict(self) -> dict[str, Any]:
        """Serializuje ustawienia do słownika (do zapisu w Tags lub config.json)."""
        return {
            "style_id": self.style_id,
            "numbering_mode": self.numbering_mode,
            "font_name": self.font_name,
            "font_size_pt": self.font_size_pt,
            "footer_position": self.footer_position,
            "footer_header": self.footer_header,
            "always_on_top": self.always_on_top,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> Settings:
        """Odtwarza `Settings` ze słownika z zachowaniem wartości domyślnych."""
        defaults = cls()
        return cls(
            style_id=str(data.get("style_id", defaults.style_id)),
            numbering_mode=data.get("numbering_mode", defaults.numbering_mode),
            font_name=str(data.get("font_name", defaults.font_name)),
            font_size_pt=int(data.get("font_size_pt", defaults.font_size_pt)),
            footer_position=data.get("footer_position", defaults.footer_position),
            footer_header=bool(data.get("footer_header", defaults.footer_header)),
            always_on_top=bool(data.get("always_on_top", defaults.always_on_top)),
        )


@runtime_checkable
class ZoteroSource(Protocol):
    """Kontrakt dostawcy danych i formatowania CSL ze Zotero (Tor B1.1)."""

    def is_available(self) -> bool:
        """Zwraca `True`, jeśli lokalne API Zotero odpowiada."""

    def search_items(self, query: str = "", limit: int = 25) -> list[ZoteroItem]:
        """Wyszukuje pozycje w bibliotece po autorze, tytule lub roku.

        Rzuca `ZoteroOfflineError`, gdy Zotero nie odpowiada.
        """

    def get_formatted_ref(self, key: str, style_id: str) -> FormattedRef:
        """Zwraca sformatowane cytowanie i opis bibliograficzny dla klucza i stylu.

        Rzuca `KeyError`, gdy pozycja nie istnieje, lub `ZoteroOfflineError`.
        """

    def get_formatted_refs(
        self, keys: Sequence[str], style_id: str
    ) -> dict[str, FormattedRef]:
        """Zwraca słownik `{key: FormattedRef}` dla podanych kluczy i stylu."""

    def list_styles(self) -> list[StyleInfo]:
        """Zwraca listę dostępnych stylów cytowania (co najmniej IEEE v4 oraz APA 7)."""


@runtime_checkable
class SlideDoc(Protocol):
    """Kontrakt warstwy dokumentu PowerPoint (Tor B2.1)."""

    def has_open_presentation(self) -> bool:
        """Zwraca `True`, jeśli w PowerPoincie jest otwarta prezentacja."""

    def get_presentation_path(self) -> Path | None:
        """Zwraca ścieżkę do pliku `.pptx` lub `None`, jeśli plik nie był zapisany."""

    def has_text_cursor(self) -> bool:
        """Zwraca `True`, jeśli kursor stoi w polu tekstowym na aktywnym slajdzie (D10)."""

    def get_active_slide_index(self) -> int:
        """Zwraca 1-based indeks aktywnego slajdu lub rzuca `NoActiveSlideError`."""

    def insert_citation_at_cursor(self, citation: Citation) -> None:
        """Wstawia cytowanie w miejscu kursora na aktywnym slajdzie i zapisuje tagi.

        Gdy kursor nie stoi w polu tekstowym, rzuca `NoTextCursorError` i nie wprowadza
        żadnych zmian w prezentacji (D10).
        """

    def set_slide_footer(
        self,
        slide_index: int,
        entries: Sequence[FooterEntry],
        settings: Settings,
    ) -> None:
        """Tworzy, aktualizuje lub usuwa (gdy `entries` puste) pole źródeł na slajdzie."""

    def read_slides(self) -> list[SlideRecord]:
        """Odczytuje wszystkie slajdy w kolejności od 1 do N wraz z ich cytowaniami i stopkami."""

    def update_slide_citations(
        self,
        slide_index: int,
        citations: Sequence[Citation],
    ) -> None:
        """Aktualizuje widoczne teksty cytowań oraz tagi na wskazanym slajdzie."""

    def load_settings(self) -> Settings | None:
        """Odczytuje ustawienia AZP zapisane w tagach prezentacji (`None` jeśli brak)."""

    def save_settings(self, settings: Settings) -> None:
        """Zapisuje ustawienia AZP w tagach aktywnej prezentacji."""

    def create_backup(self) -> Path | None:
        """Tworzy kopię zapasową pliku prezentacji przed przebudową całej prezentacji."""
