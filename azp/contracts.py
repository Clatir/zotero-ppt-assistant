"""Wspólne typy danych, wyjątki i interfejsy (kontrakty) dla modułów AZP.

Stanowi granicę między torem B1 (`azp/zotero`, `azp/gui`) a torem B2 (`azp/ppt`, `azp/engine`).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal, Protocol, Sequence, runtime_checkable

STYLE_IEEE_V4 = "http://www.zotero.org/styles/ieee-praca-inzynierska-access-date-v4"
STYLE_IEEE_FALLBACK = "http://www.zotero.org/styles/ieee"
STYLE_APA = "http://www.zotero.org/styles/apa"

NumberingMode = Literal["global", "per_slide"]
FooterPosition = Literal["bottom_full", "bottom_left", "replace_footer"]
BibFitMode = Literal["shrink_to_fit", "fixed_font", "split_slides"]

VALID_NUMBERING_MODES: tuple[NumberingMode, ...] = ("global", "per_slide")
VALID_FOOTER_POSITIONS: tuple[FooterPosition, ...] = (
    "bottom_full",
    "bottom_left",
    "replace_footer",
)
VALID_BIB_FIT_MODES: tuple[BibFitMode, ...] = (
    "shrink_to_fit",
    "fixed_font",
    "split_slides",
)
VALID_BIB_COLUMNS: tuple[int, ...] = (1, 2, 3, 4)


class AzpError(Exception):
    """Bazowy wyjątek aplikacyjny AZP z komunikatem dla użytkownika po polsku."""


class ZoteroOfflineError(AzpError):
    """Zotero jest wyłączone albo lokalne API nie odpowiada."""


class NoActiveSlideError(AzpError):
    """Brak otwartej prezentacji lub aktywnego slajdu w PowerPoincie."""


class NoTextCursorError(AzpError):
    """Kursor nie znajduje się w polu tekstowym na slajdzie (D10)."""


MSG_SAVE_BEFORE_REBUILD = "Zapisz prezentację, żeby można było zrobić kopię zapasową."


class BackupError(AzpError):
    """Nie można zrobić kopii zapasowej przed przebudową – przebudowa nie jest wykonywana (D17, D18)."""


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
class ZoteroCollection:
    """Kolekcja (lub podkolekcja) w bibliotece Zotero (D25)."""

    key: str
    name: str
    parent_key: str = ""


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
    """Wystąpienie cytowania na slajdzie.

    Może obejmować jeden lub więcej kluczy Zotero (np. scalone `[1, 2]`
    albo `(Kowalski, 2023; Nowak, 2024)`). Gdy `in_text=False`, cytowanie zostało
    dodane bez kursora tekstowego (lub na slajdzie ze zdjęciami) i widnieje wyłącznie
    w polu źródeł (stopce) na dole slajdu oraz w bibliografii (D26).
    """

    citation_id: str
    keys: tuple[str, ...]
    text: str = ""
    in_text: bool = True

    def __post_init__(self) -> None:
        if not self.keys:
            raise ValueError("Citation.keys nie może być puste.")


@dataclass(frozen=True)
class FooterEntry:
    """Pojedynczy wpis w polu źródeł (stopce) pod slajdem lub w bibliografii."""

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
    """Stan cytowań, stopki i bibliografii na pojedynczym slajdzie prezentacji."""

    slide_index: int
    slide_id: int
    citations: tuple[Citation, ...] = ()
    footer_entries: tuple[FooterEntry, ...] = ()
    has_bibliography: bool = False


@dataclass(frozen=True)
class Settings:
    """Ustawienia prezentacji i okienka (D2–D5, D9, D21, D27)."""

    style_id: str = STYLE_IEEE_V4
    numbering_mode: NumberingMode = "global"
    font_name: str = "Calibri"
    font_size_pt: int = 10
    footer_position: FooterPosition = "bottom_full"
    footer_header: bool = True
    always_on_top: bool = True
    dark_mode: bool = False
    bib_font_name: str = "Calibri"
    bib_font_size_pt: int = 12
    bib_fit_mode: BibFitMode = "shrink_to_fit"
    bib_columns: int = 1

    def __post_init__(self) -> None:
        if self.numbering_mode not in VALID_NUMBERING_MODES:
            raise ValueError(f"Nieprawidłowy tryb numeracji: {self.numbering_mode!r}")
        if self.footer_position not in VALID_FOOTER_POSITIONS:
            raise ValueError(f"Nieprawidłowe położenie stopki: {self.footer_position!r}")
        if (
            not isinstance(self.font_size_pt, int)
            or isinstance(self.font_size_pt, bool)
            or self.font_size_pt < 6
            or self.font_size_pt > 72
        ):
            raise ValueError(f"Rozmiar czcionki poza zakresem 6..72 pt: {self.font_size_pt!r}")
        if not isinstance(self.font_name, str) or not self.font_name.strip():
            raise ValueError("Nazwa czcionki nie może być pusta.")
        if not isinstance(self.style_id, str) or not self.style_id.strip():
            raise ValueError("Identyfikator stylu nie może być pusty.")
        if not isinstance(self.footer_header, bool):
            raise ValueError(f"Nieprawidłowa wartość footer_header: {self.footer_header!r}")
        if not isinstance(self.always_on_top, bool):
            raise ValueError(f"Nieprawidłowa wartość always_on_top: {self.always_on_top!r}")
        if not isinstance(self.dark_mode, bool):
            raise ValueError(f"Nieprawidłowa wartość dark_mode: {self.dark_mode!r}")
        if not isinstance(self.bib_font_name, str) or not self.bib_font_name.strip():
            raise ValueError("Nazwa czcionki bibliografii nie może być pusta.")
        if (
            not isinstance(self.bib_font_size_pt, int)
            or isinstance(self.bib_font_size_pt, bool)
            or self.bib_font_size_pt < 6
            or self.bib_font_size_pt > 72
        ):
            raise ValueError(
                f"Rozmiar czcionki bibliografii poza zakresem 6..72 pt: {self.bib_font_size_pt!r}"
            )
        if self.bib_fit_mode not in VALID_BIB_FIT_MODES:
            raise ValueError(
                f"Nieprawidłowy tryb dopasowania bibliografii: {self.bib_fit_mode!r}"
            )
        if (
            not isinstance(self.bib_columns, int)
            or isinstance(self.bib_columns, bool)
            or self.bib_columns not in VALID_BIB_COLUMNS
        ):
            raise ValueError(
                f"Liczba kolumn bibliografii poza zakresem {VALID_BIB_COLUMNS}: {self.bib_columns!r}"
            )

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
            "dark_mode": self.dark_mode,
            "bib_font_name": self.bib_font_name,
            "bib_font_size_pt": self.bib_font_size_pt,
            "bib_fit_mode": self.bib_fit_mode,
            "bib_columns": self.bib_columns,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> Settings:
        """Odtwarza `Settings` ze słownika z zachowaniem wartości domyślnych."""
        defaults = cls()
        return cls(
            style_id=data.get("style_id", defaults.style_id),
            numbering_mode=data.get("numbering_mode", defaults.numbering_mode),
            font_name=data.get("font_name", defaults.font_name),
            font_size_pt=data.get("font_size_pt", defaults.font_size_pt),
            footer_position=data.get("footer_position", defaults.footer_position),
            footer_header=data.get("footer_header", defaults.footer_header),
            always_on_top=data.get("always_on_top", defaults.always_on_top),
            dark_mode=data.get("dark_mode", defaults.dark_mode),
            bib_font_name=data.get("bib_font_name", defaults.bib_font_name),
            bib_font_size_pt=data.get("bib_font_size_pt", defaults.bib_font_size_pt),
            bib_fit_mode=data.get("bib_fit_mode", defaults.bib_fit_mode),
            bib_columns=data.get("bib_columns", defaults.bib_columns),
        )


@runtime_checkable
class ZoteroSource(Protocol):
    """Kontrakt dostawcy danych i formatowania CSL ze Zotero (Tor B1.1)."""

    def is_available(self) -> bool:
        """Zwraca `True`, jeśli lokalne API Zotero odpowiada."""

    def search_items(
        self,
        query: str = "",
        limit: int = 25,
        collection_key: str | None = None,
    ) -> list[ZoteroItem]:
        """Wyszukuje pozycje w bibliotece (lub we wskazanej kolekcji) po autorze, tytule lub roku.

        Rzuca `ZoteroOfflineError`, gdy Zotero nie odpowiada.
        """

    def list_collections(self) -> list[ZoteroCollection]:
        """Zwraca listę kolekcji i podkolekcji w bibliotece Zotero (D25)."""

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
        """Zwraca `True`, jeśli kursor stoi w polu tekstowym na aktywnym slajdzie."""

    def get_active_slide_index(self) -> int:
        """Zwraca 1-based indeks aktywnego slajdu lub rzuca `NoActiveSlideError`."""

    def insert_citation_at_cursor(self, citation: Citation) -> None:
        """Wstawia cytowanie w miejscu kursora na aktywnym slajdzie i zapisuje tagi.

        Gdy kursor nie stoi w polu tekstowym, rzuca `NoTextCursorError`.
        """

    def add_footer_citation(self, citation: Citation) -> None:
        """Dodaje przypis (cytowanie stopkowe `in_text=False`) na aktywnym slajdzie bez kursora (D26)."""

    def set_slide_footer(
        self,
        slide_index: int,
        entries: Sequence[FooterEntry],
        settings: Settings,
    ) -> None:
        """Tworzy, aktualizuje lub usuwa (gdy `entries` puste) pole źródeł na slajdzie."""

    def set_slide_bibliography(
        self,
        slide_index: int,
        entries: Sequence[FooterEntry],
        settings: Settings,
    ) -> int:
        """Wstawia lub aktualizuje spis wykorzystanych źródeł (bez nagłówka) na slajdzie (D27).

        Zwraca liczbę slajdów zajętych przez bibliografię (1 lub więcej w trybie `split_slides`).
        """

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
