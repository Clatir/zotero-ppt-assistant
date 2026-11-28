"""Pływające okienko tkinter do cytowania pozycji ze Zotero w PowerPoincie (etap B1.2)."""

from __future__ import annotations

import ctypes
import tkinter as tk
from dataclasses import replace
from pathlib import Path
from tkinter import font as tkfont
from tkinter import ttk
from typing import Callable, Sequence

from azp.config import load_config, save_config
from azp.contracts import (
    MSG_SAVE_BEFORE_REBUILD,
    STYLE_APA,
    STYLE_IEEE_V4,
    VALID_BIB_COLUMNS,
    AzpError,
    BackupError,
    BibFitMode,
    FooterPosition,
    NoActiveSlideError,
    NoTextCursorError,
    NumberingMode,
    Settings,
    SlideDoc,
    StyleInfo,
    ZoteroCollection,
    ZoteroItem,
    ZoteroOfflineError,
    ZoteroSource,
)
from azp.engine import CitationEngine

CiteCallback = Callable[[tuple[str, ...], Settings], str | None]
RebuildCallback = Callable[[Settings], str | None]
BibliographyCallback = Callable[[Settings], str | None]

NUMBERING_LABELS: dict[NumberingMode, str] = {
    "global": "Globalna (cała prezentacja)",
    "per_slide": "Od [1] na każdym slajdzie",
}
LABEL_TO_NUMBERING: dict[str, NumberingMode] = {
    label: mode for mode, label in NUMBERING_LABELS.items()
}

FOOTER_POSITION_LABELS: dict[FooterPosition, str] = {
    "bottom_full": "Dół slajdu – cała szerokość (nad stopką)",
    "bottom_left": "Lewy dolny róg – węższe pole",
    "replace_footer": "W miejscu stopki PowerPointa",
}
LABEL_TO_FOOTER_POSITION: dict[str, FooterPosition] = {
    label: pos for pos, label in FOOTER_POSITION_LABELS.items()
}

FOOTER_HEADER_LABELS: dict[bool, str] = {
    True: "Z nagłówkiem „Źródła:”",
    False: "Bez nagłówka",
}
LABEL_TO_FOOTER_HEADER: dict[str, bool] = {
    label: enabled for enabled, label in FOOTER_HEADER_LABELS.items()
}

BIB_FIT_LABELS: dict[BibFitMode, str] = {
    "shrink_to_fit": "Zmniejsz czcionkę do slajdu (Auto-fit)",
    "fixed_font": "Stała czcionka (dopasuj ramkę do tekstu)",
    "split_slides": "Podziel automatycznie na kolejne slajdy",
}
LABEL_TO_BIB_FIT: dict[str, BibFitMode] = {
    label: mode for mode, label in BIB_FIT_LABELS.items()
}

ALL_COLLECTIONS_LABEL = "Wszystkie pozycje (cała biblioteka)"

DEFAULT_FONT_SIZES: tuple[int, ...] = (8, 9, 10, 11, 12, 14, 16)
DEFAULT_BIB_FONT_SIZES: tuple[int, ...] = (*DEFAULT_FONT_SIZES, 18, 20, 24)

DARK_THEME_COLORS: dict[str, str] = {
    "bg": "#1e1e2e",
    "field_bg": "#2b2b3c",
    "fg": "#e6e6e6",
    "select_bg": "#3a6ea5",
    "select_fg": "#ffffff",
    "button_bg": "#313244",
    "button_active_bg": "#45475a",
    "border": "#45475a",
    "status_bg": "#181825",
}

LIGHT_THEME_COLORS: dict[str, str] = {
    "bg": "#f0f0f0",
    "field_bg": "#ffffff",
    "fg": "#111111",
    "select_bg": "#0078d7",
    "select_fg": "#ffffff",
    "button_bg": "#e1e1e1",
    "button_active_bg": "#d0d0d0",
    "border": "#cccccc",
    "status_bg": "#e8e8e8",
}

_FALLBACK_STYLES: tuple[StyleInfo, ...] = (
    StyleInfo(
        style_id=STYLE_IEEE_V4,
        title="IEEE Praca Inzynierska Access Date v4",
        is_numeric=True,
    ),
    StyleInfo(
        style_id=STYLE_APA,
        title="APA Style 7th edition",
        is_numeric=False,
    ),
)


def collect_system_fonts(root: tk.Misc) -> list[str]:
    """Zwraca posortowaną listę unikalnych czcionek systemowych z Calibri i Arial."""
    raw_families = tkfont.families(root)
    fonts = {name.strip() for name in raw_families if name and name.strip()}
    fonts.update({"Calibri", "Arial"})
    return sorted(fonts)


def _format_zotero_offline_message(exc: ZoteroOfflineError) -> str:
    msg = str(exc).strip()
    if "Zotero nie odpowiada" in msg:
        return msg
    if msg:
        return f"Zotero nie odpowiada: {msg}"
    return "Zotero nie odpowiada. Uruchom Zotero 10 i upewnij się, że lokalne API jest włączone."


def _describe_unexpected_error(exc: BaseException) -> str:
    args = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 3:
        desc = args[2][2]
        if isinstance(desc, str) and desc.strip():
            return desc.strip()
    if len(args) >= 2 and isinstance(args[1], str) and args[1].strip():
        return args[1].strip()
    msg = str(exc).strip()
    return msg if msg else type(exc).__name__


class AzpController:
    """Kontroler logiki okienka AZP oddzielony od warstwy widgetów tkinter."""

    def __init__(
        self,
        zotero: ZoteroSource,
        slide_doc: SlideDoc,
        settings: Settings | None = None,
        on_cite: CiteCallback | None = None,
        on_rebuild: RebuildCallback | None = None,
        on_bibliography: BibliographyCallback | None = None,
        config_path: Path | None = None,
    ) -> None:
        self.zotero = zotero
        self.slide_doc = slide_doc
        self.on_cite = on_cite
        self.on_rebuild = on_rebuild
        self.on_bibliography = on_bibliography
        self.config_path = config_path
        self._engine = CitationEngine(zotero=zotero, slide_doc=slide_doc)
        self._builtin_on_cite = self._engine.make_cite_callback()
        self._builtin_on_rebuild = self._engine.make_rebuild_callback()
        self._builtin_on_bibliography = self._engine.make_bibliography_callback()

        if settings is not None:
            self.settings = settings
        else:
            base_settings = load_config(config_path) if config_path is not None else Settings()
            try:
                loaded = (
                    slide_doc.load_settings()
                    if slide_doc.has_open_presentation()
                    else None
                )
            except Exception:
                loaded = None
            if loaded is not None:
                if config_path is not None:
                    self.settings = replace(
                        loaded,
                        always_on_top=base_settings.always_on_top,
                        dark_mode=base_settings.dark_mode,
                    )
                else:
                    self.settings = loaded
            else:
                self.settings = base_settings

        self.results: list[ZoteroItem] = []
        self.last_query: str = ""
        self.selected_collection_key: str | None = None
        self.status_message: str = "Gotowy."
        self._styles: list[StyleInfo] = self._load_styles()
        self._collections: list[ZoteroCollection] = self._load_collections()
        self._active_presentation_key: str | None = self._current_presentation_key()

    def _current_presentation_key(self) -> str | None:
        try:
            if not self.slide_doc.has_open_presentation():
                return None
        except Exception:
            return None
        get_key = getattr(self.slide_doc, "get_presentation_key", None)
        if callable(get_key):
            try:
                key = get_key()
            except Exception:
                return None
            if key is not None and str(key).strip():
                return str(key).strip()
        try:
            path = self.slide_doc.get_presentation_path()
        except Exception:
            return None
        return str(path) if path is not None else "__open_presentation__"

    def sync_with_active_presentation(self) -> bool:
        """Sprawdza, czy użytkownik przełączył aktywną prezentację i wczytuje jej ustawienia (D19)."""
        current_key = self._current_presentation_key()
        if current_key is None:
            self._active_presentation_key = None
            return False
        if current_key == self._active_presentation_key:
            return False

        self._active_presentation_key = current_key
        try:
            loaded = self.slide_doc.load_settings()
        except Exception:
            loaded = None

        current_topmost = self.settings.always_on_top
        current_dark_mode = self.settings.dark_mode
        if loaded is not None:
            self.settings = replace(
                loaded,
                always_on_top=current_topmost,
                dark_mode=current_dark_mode,
            )

        label = Path(current_key).name or current_key
        if label == "__open_presentation__":
            self.status_message = "Przełączono aktywną prezentację."
        else:
            self.status_message = f"Aktywna prezentacja: {label}."
        return True

    def _load_styles(self) -> list[StyleInfo]:
        try:
            styles = self.zotero.list_styles()
        except ZoteroOfflineError as exc:
            self.status_message = _format_zotero_offline_message(exc)
            return list(_FALLBACK_STYLES)
        return list(styles) if styles else list(_FALLBACK_STYLES)

    def get_styles(self) -> list[StyleInfo]:
        """Zwraca listę stylów cytowania dostępnych w Zotero."""
        return list(self._styles)

    def _load_collections(self) -> list[ZoteroCollection]:
        list_cols = getattr(self.zotero, "list_collections", None)
        if not callable(list_cols):
            return []
        try:
            collections = list_cols()
        except ZoteroOfflineError as exc:
            self.status_message = _format_zotero_offline_message(exc)
            return []
        return list(collections) if collections else []

    def get_collections(self) -> list[ZoteroCollection]:
        """Zwraca listę kolekcji i podkolekcji w bibliotece Zotero (D25)."""
        return list(self._collections)

    def refresh_collections(self) -> list[ZoteroCollection]:
        """Odświeża listę kolekcji z Zotero (D25)."""
        self._collections = self._load_collections()
        return list(self._collections)

    def set_collection(self, collection_key: str | None) -> list[ZoteroItem]:
        """Ustawia aktywny filtr kolekcji Zotero i ponawia wyszukiwanie (D25)."""
        clean_key = (
            collection_key.strip()
            if collection_key and collection_key.strip()
            else None
        )
        self.selected_collection_key = clean_key
        return self.search(self.last_query)

    def search(self, query: str = "", limit: int = 25) -> list[ZoteroItem]:
        """Wyszukuje pozycje w Zotero i aktualizuje listę wyników oraz pasek stanu."""
        self.last_query = query
        try:
            if self.selected_collection_key is not None:
                try:
                    items = self.zotero.search_items(
                        query=query,
                        limit=limit,
                        collection_key=self.selected_collection_key,
                    )
                except TypeError:
                    items = self.zotero.search_items(query=query, limit=limit)
            else:
                items = self.zotero.search_items(query=query, limit=limit)
        except ZoteroOfflineError as exc:
            self.results = []
            self.status_message = _format_zotero_offline_message(exc)
            return []

        if not self._collections:
            self.refresh_collections()

        self.results = list(items)
        if not self.results:
            self.status_message = "Brak wyników w Zotero dla podanego zapytania."
        else:
            self.status_message = f"Znaleziono pozycji: {len(self.results)}."
        return list(self.results)

    def cite_keys(self, keys: Sequence[str]) -> bool:
        """Wstawia cytowanie (w miejscu kursora lub w stopce na dole slajdu – D26)."""
        self.sync_with_active_presentation()
        selected_keys = tuple(dict.fromkeys(keys))
        if not selected_keys:
            self.status_message = "Wybierz pozycję do zacytowania."
            return False

        try:
            if not self.slide_doc.has_open_presentation():
                self.status_message = "Brak otwartej prezentacji w PowerPoincie."
                return False

            self.slide_doc.get_active_slide_index()
        except NoActiveSlideError:
            self.status_message = "Brak otwartej prezentacji lub aktywnego slajdu."
            return False
        except AzpError as exc:
            self.status_message = str(exc)
            return False
        except Exception as exc:
            self.status_message = f"PowerPoint zgłosił błąd: {_describe_unexpected_error(exc)}"
            return False

        callback = self.on_cite if self.on_cite is not None else self._builtin_on_cite
        try:
            msg = callback(selected_keys, self.settings)
        except ZoteroOfflineError as exc:
            self.status_message = _format_zotero_offline_message(exc)
            return False
        except NoTextCursorError:
            self.status_message = "Ustaw kursor w polu tekstowym na slajdzie."
            return False
        except NoActiveSlideError:
            self.status_message = "Brak otwartej prezentacji lub aktywnego slajdu."
            return False
        except AzpError as exc:
            self.status_message = str(exc)
            return False
        except Exception as exc:
            self.status_message = f"PowerPoint zgłosił błąd: {_describe_unexpected_error(exc)}"
            return False

        self.status_message = msg or "Wstawiono cytowanie."
        return True

    def insert_bibliography(self) -> bool:
        """Wstawia lub aktualizuje spis wykorzystanych źródeł na aktywnym slajdzie (D27)."""
        self.sync_with_active_presentation()
        try:
            if not self.slide_doc.has_open_presentation():
                self.status_message = "Brak otwartej prezentacji w PowerPoincie."
                return False

            self.slide_doc.get_active_slide_index()
        except NoActiveSlideError:
            self.status_message = "Brak otwartej prezentacji lub aktywnego slajdu."
            return False
        except AzpError as exc:
            self.status_message = str(exc)
            return False
        except Exception as exc:
            self.status_message = f"PowerPoint zgłosił błąd: {_describe_unexpected_error(exc)}"
            return False

        callback = (
            self.on_bibliography
            if self.on_bibliography is not None
            else self._builtin_on_bibliography
        )
        try:
            msg = callback(self.settings)
        except ZoteroOfflineError as exc:
            self.status_message = _format_zotero_offline_message(exc)
            return False
        except NoActiveSlideError:
            self.status_message = "Brak otwartej prezentacji lub aktywnego slajdu."
            return False
        except BackupError as exc:
            self.status_message = str(exc) or MSG_SAVE_BEFORE_REBUILD
            return False
        except AzpError as exc:
            self.status_message = str(exc)
            return False
        except Exception as exc:
            self.status_message = f"PowerPoint zgłosił błąd: {_describe_unexpected_error(exc)}"
            return False

        self.status_message = msg or "Wstawiono bibliografię."
        return True

    def refresh_presentation(self) -> bool:
        """Tworzy kopię zapasową i przebudowuje cytowania, stopki oraz bibliografię."""
        self.sync_with_active_presentation()
        try:
            if not self.slide_doc.has_open_presentation():
                self.status_message = "Brak otwartej prezentacji w PowerPoincie."
                return False

            if self.slide_doc.get_presentation_path() is None:
                self.status_message = MSG_SAVE_BEFORE_REBUILD
                return False
        except AzpError as exc:
            self.status_message = str(exc)
            return False
        except Exception as exc:
            self.status_message = f"PowerPoint zgłosił błąd: {_describe_unexpected_error(exc)}"
            return False

        callback = (
            self.on_rebuild if self.on_rebuild is not None else self._builtin_on_rebuild
        )
        try:
            msg = callback(self.settings)
        except ZoteroOfflineError as exc:
            self.status_message = _format_zotero_offline_message(exc)
            return False
        except BackupError as exc:
            self.status_message = str(exc) or MSG_SAVE_BEFORE_REBUILD
            return False
        except AzpError as exc:
            self.status_message = str(exc)
            return False
        except Exception as exc:
            self.status_message = f"PowerPoint zgłosił błąd: {_describe_unexpected_error(exc)}"
            return False

        self.status_message = msg or "Przebudowano prezentację."
        return True

    def _save_config_safe(self) -> None:
        if self.config_path is None:
            return
        try:
            save_config(self.settings, self.config_path)
        except OSError:
            pass

    def update_settings(self, rebuild: bool = True, **changes: object) -> bool:
        """Aktualizuje ustawienia i przy zmianie opcji formatowania przebudowuje prezentację (D14)."""
        self.sync_with_active_presentation()
        previous_settings = self.settings
        new_settings = replace(previous_settings, **changes)  # type: ignore[arg-type]
        if new_settings == previous_settings:
            return True

        self.settings = new_settings
        if not rebuild:
            try:
                if self.slide_doc.has_open_presentation():
                    self.slide_doc.save_settings(self.settings)
            except Exception:
                pass
            self._save_config_safe()
            return True

        try:
            has_pres = self.slide_doc.has_open_presentation()
        except Exception:
            has_pres = False

        if not has_pres:
            self._save_config_safe()
            self.status_message = "Zapisano ustawienie (brak otwartej prezentacji)."
            return False

        if not self.refresh_presentation():
            # Wycofaj zmianę, jeśli przebudowa nie powiodła się (np. Zotero offline)
            self.settings = previous_settings
            return False
        self._save_config_safe()
        return True


class AzpWindow:
    """Pływające okienko tkinter dla Asystenta Zotero dla PowerPointa."""

    def __init__(
        self,
        zotero: ZoteroSource,
        slide_doc: SlideDoc,
        settings: Settings | None = None,
        master: tk.Misc | None = None,
        on_cite: CiteCallback | None = None,
        on_rebuild: RebuildCallback | None = None,
        on_bibliography: BibliographyCallback | None = None,
        initial_search: bool = True,
        poll_interval_ms: int | None = 500,
        config_path: Path | None = None,
    ) -> None:
        self.controller = AzpController(
            zotero=zotero,
            slide_doc=slide_doc,
            settings=settings,
            on_cite=on_cite,
            on_rebuild=on_rebuild,
            on_bibliography=on_bibliography,
            config_path=config_path,
        )

        self.root: tk.Tk | tk.Toplevel
        if master is None:
            self.root = tk.Tk()
        else:
            self.root = tk.Toplevel(master)

        self.root.title("Asystent Zotero dla PowerPointa")
        self.root.geometry("430x760")
        self.root.minsize(380, 640)

        self._style_title_to_id: dict[str, str] = {}
        self._style_id_to_title: dict[str, str] = {}
        self._loaded_collections: tuple[ZoteroCollection, ...] = ()
        self._collection_label_to_key: dict[str, str] = {}
        self._collection_key_to_label: dict[str, str] = {}
        self._collection_name_to_key: dict[str, str] = {}
        self._poll_interval_ms: int | None = poll_interval_ms
        self._poll_after_id: str | None = None
        self._ttk_style = ttk.Style(self.root)
        self._default_ttk_theme = self._ttk_style.theme_use()
        self._applied_dark_mode: bool | None = None
        self._build_variables()
        self._build_ui()
        self._sync_widgets_from_settings()
        self._apply_topmost(self.controller.settings.always_on_top)
        self._apply_theme(self.controller.settings.dark_mode)

        if initial_search:
            self.perform_search("")

        self._schedule_presentation_poll()

    def _build_variables(self) -> None:
        self.collection_var = tk.StringVar(
            master=self.root, value=ALL_COLLECTIONS_LABEL
        )
        self.query_var = tk.StringVar(master=self.root, value="")
        self.style_var = tk.StringVar(master=self.root, value="")
        self.numbering_var = tk.StringVar(master=self.root, value="")
        self.font_name_var = tk.StringVar(master=self.root, value="")
        self.font_size_var = tk.StringVar(master=self.root, value="")
        self.footer_position_var = tk.StringVar(master=self.root, value="")
        self.footer_header_var = tk.StringVar(master=self.root, value="")
        self.bib_font_name_var = tk.StringVar(master=self.root, value="")
        self.bib_font_size_var = tk.StringVar(master=self.root, value="")
        self.bib_fit_var = tk.StringVar(master=self.root, value="")
        self.bib_columns_var = tk.StringVar(master=self.root, value="")
        self.always_on_top_var = tk.BooleanVar(
            master=self.root, value=self.controller.settings.always_on_top
        )
        self.dark_mode_var = tk.BooleanVar(
            master=self.root, value=self.controller.settings.dark_mode
        )
        self.status_var = tk.StringVar(
            master=self.root, value=self.controller.status_message
        )

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=10, style="Azp.TFrame")
        main.pack(fill=tk.BOTH, expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)
        self._main_frame = main

        # 1. Wyszukiwarka i wybór kolekcji (D25)
        search_frame = ttk.LabelFrame(
            main, text="Wyszukaj w Zotero", padding=8, style="Azp.TLabelframe"
        )
        search_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        search_frame.columnconfigure(1, weight=1)
        self._search_frame = search_frame

        collection_values = self._rebuild_collection_mappings()

        collection_lbl = ttk.Label(
            search_frame, text="Kolekcja:", style="Azp.TLabel"
        )
        collection_lbl.grid(row=0, column=0, sticky="w", padx=(0, 6), pady=(0, 6))

        self.collection_combo = ttk.Combobox(
            search_frame,
            textvariable=self.collection_var,
            values=collection_values,
            state="readonly",
            style="Azp.TCombobox",
        )
        self.collection_combo.grid(
            row=0, column=1, columnspan=2, sticky="ew", pady=(0, 6)
        )
        self.collection_combo.bind(
            "<<ComboboxSelected>>", self._on_collection_selected
        )

        self.search_entry = ttk.Entry(
            search_frame, textvariable=self.query_var, style="Azp.TEntry"
        )
        self.search_entry.grid(
            row=1, column=0, columnspan=2, sticky="ew", padx=(0, 6)
        )
        self.search_entry.bind("<Return>", self._on_search_event)
        self.search_entry.bind("<Control-Return>", self._on_cite_shortcut)
        self.search_entry.bind("<Down>", self._on_search_down_arrow)

        self.search_button = ttk.Button(
            search_frame,
            text="Szukaj",
            command=self.perform_search,
            style="Azp.TButton",
        )
        self.search_button.grid(row=1, column=2)

        # 2. Lista wyników (D13: wybór pojedynczy i wielokrotny Ctrl/Shift)
        results_frame = ttk.LabelFrame(
            main,
            text="Wyniki (Ctrl / Shift – wybór wielu)",
            padding=8,
            style="Azp.TLabelframe",
        )
        results_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 8))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        self._results_frame = results_frame

        self.results_listbox = tk.Listbox(
            results_frame,
            selectmode=tk.EXTENDED,
            exportselection=False,
            activestyle="dotbox",
            height=9,
        )
        self.results_listbox.grid(row=0, column=0, sticky="nsew")
        self.results_listbox.bind("<Return>", self._on_cite_shortcut)
        self.results_listbox.bind("<Control-Return>", self._on_cite_shortcut)
        self.results_listbox.bind("<Double-Button-1>", self._on_cite_shortcut)

        scrollbar = ttk.Scrollbar(
            results_frame, orient=tk.VERTICAL, command=self.results_listbox.yview
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.results_listbox.configure(yscrollcommand=scrollbar.set)

        # 3. Główne akcje (D26, D27)
        actions_frame = ttk.Frame(main, style="Azp.TFrame")
        actions_frame.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        actions_frame.columnconfigure(0, weight=1)
        actions_frame.columnconfigure(1, weight=1)
        self._actions_frame = actions_frame

        self.cite_button = ttk.Button(
            actions_frame,
            text="Cytuj",
            command=self.cite_selected,
            style="Azp.TButton",
        )
        self.cite_button.grid(row=0, column=0, sticky="ew", padx=(0, 4))

        self.refresh_button = ttk.Button(
            actions_frame,
            text="Odśwież / Przenumeruj",
            command=self.refresh_presentation,
            style="Azp.TButton",
        )
        self.refresh_button.grid(row=0, column=1, sticky="ew", padx=(4, 0))

        self.bibliography_button = ttk.Button(
            actions_frame,
            text="Wstaw bibliografię",
            command=self.insert_bibliography,
            style="Azp.TButton",
        )
        self.bibliography_button.grid(
            row=1, column=0, columnspan=2, sticky="ew", pady=(4, 0)
        )

        # 4. Ustawienia (D2–D5, D9, D14, D27)
        settings_frame = ttk.LabelFrame(
            main,
            text="Ustawienia cytowań i pola źródeł",
            padding=8,
            style="Azp.TLabelframe",
        )
        settings_frame.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        settings_frame.columnconfigure(1, weight=1)
        self._settings_frame = settings_frame

        styles = self.controller.get_styles()
        self._style_title_to_id = {s.title: s.style_id for s in styles}
        self._style_id_to_title = {s.style_id: s.title for s in styles}
        style_titles = [s.title for s in styles]
        if self.controller.settings.style_id not in self._style_id_to_title:
            fallback_title = self.controller.settings.style_id
            self._style_title_to_id[fallback_title] = fallback_title
            self._style_id_to_title[fallback_title] = fallback_title
            style_titles.append(fallback_title)

        self.style_combo = self._add_setting_row(
            settings_frame,
            row=0,
            label="Styl cytowania:",
            variable=self.style_var,
            values=style_titles,
            handler=self._on_style_selected,
        )

        self.numbering_combo = self._add_setting_row(
            settings_frame,
            row=1,
            label="Numeracja:",
            variable=self.numbering_var,
            values=list(NUMBERING_LABELS.values()),
            handler=self._on_numbering_selected,
        )

        fonts = collect_system_fonts(self.root)
        if self.controller.settings.font_name not in fonts:
            fonts = sorted({*fonts, self.controller.settings.font_name})
        self.font_combo = self._add_setting_row(
            settings_frame,
            row=2,
            label="Czcionka stopki:",
            variable=self.font_name_var,
            values=fonts,
            handler=self._on_font_selected,
        )

        font_sizes = sorted({*DEFAULT_FONT_SIZES, self.controller.settings.font_size_pt})
        self.font_size_combo = self._add_setting_row(
            settings_frame,
            row=3,
            label="Rozmiar (pt):",
            variable=self.font_size_var,
            values=[str(size) for size in font_sizes],
            handler=self._on_font_size_selected,
        )

        self.footer_position_combo = self._add_setting_row(
            settings_frame,
            row=4,
            label="Położenie źródeł:",
            variable=self.footer_position_var,
            values=list(FOOTER_POSITION_LABELS.values()),
            handler=self._on_footer_position_selected,
        )

        self.footer_header_combo = self._add_setting_row(
            settings_frame,
            row=5,
            label="Nagłówek źródeł:",
            variable=self.footer_header_var,
            values=list(FOOTER_HEADER_LABELS.values()),
            handler=self._on_footer_header_selected,
        )

        bib_fonts = sorted({*fonts, self.controller.settings.bib_font_name})
        self.bib_font_combo = self._add_setting_row(
            settings_frame,
            row=6,
            label="Czcionka bibliografii:",
            variable=self.bib_font_name_var,
            values=bib_fonts,
            handler=self._on_bib_font_selected,
        )

        bib_font_sizes = sorted(
            {*DEFAULT_BIB_FONT_SIZES, self.controller.settings.bib_font_size_pt}
        )
        self.bib_font_size_combo = self._add_setting_row(
            settings_frame,
            row=7,
            label="Rozmiar bibliografii (pt):",
            variable=self.bib_font_size_var,
            values=[str(size) for size in bib_font_sizes],
            handler=self._on_bib_font_size_selected,
        )

        self.bib_fit_combo = self._add_setting_row(
            settings_frame,
            row=8,
            label="Dopasowanie bibliografii:",
            variable=self.bib_fit_var,
            values=list(BIB_FIT_LABELS.values()),
            handler=self._on_bib_fit_selected,
        )

        self.bib_columns_combo = self._add_setting_row(
            settings_frame,
            row=9,
            label="Kolumny bibliografii:",
            variable=self.bib_columns_var,
            values=[str(col) for col in VALID_BIB_COLUMNS],
            handler=self._on_bib_columns_selected,
        )

        # 5. Zawsze na wierzchu + Tryb ciemny
        options_frame = ttk.Frame(main, style="Azp.TFrame")
        options_frame.grid(row=4, column=0, sticky="ew", pady=(0, 8))
        self._options_frame = options_frame

        self.always_on_top_check = ttk.Checkbutton(
            options_frame,
            text="Zawsze na wierzchu",
            variable=self.always_on_top_var,
            command=self._on_always_on_top_toggled,
            style="Azp.TCheckbutton",
        )
        self.always_on_top_check.pack(side=tk.LEFT, anchor="w")

        self.dark_mode_check = ttk.Checkbutton(
            options_frame,
            text="Tryb ciemny",
            variable=self.dark_mode_var,
            command=self._on_dark_mode_toggled,
            style="Azp.TCheckbutton",
        )
        self.dark_mode_check.pack(side=tk.RIGHT, anchor="e")

        # 6. Pasek stanu
        self.status_label = ttk.Label(
            main,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor="w",
            padding=(6, 4),
            wraplength=390,
            style="AzpStatus.TLabel",
        )
        self.status_label.grid(row=5, column=0, sticky="ew")
        self._bind_shortcuts()

    def _bind_shortcuts(self) -> None:
        """Rejestruje skróty klawiszowe wewnątrz okienka (D15)."""
        self.root.bind("<Control-Return>", self._on_cite_shortcut)
        self.root.bind("<F5>", self._on_refresh_shortcut)
        self.root.bind("<Control-r>", self._on_refresh_shortcut)
        self.root.bind("<Control-R>", self._on_refresh_shortcut)
        self.root.bind("<Control-f>", self._on_focus_search_shortcut)
        self.root.bind("<Control-F>", self._on_focus_search_shortcut)

    def _add_setting_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        values: Sequence[str],
        handler: Callable[[tk.Event[ttk.Combobox]], None],
    ) -> ttk.Combobox:
        lbl = ttk.Label(parent, text=label, style="Azp.TLabel")
        lbl.grid(row=row, column=0, sticky="w", padx=(0, 8), pady=3)
        combo = ttk.Combobox(
            parent,
            textvariable=variable,
            values=list(values),
            state="readonly",
            style="Azp.TCombobox",
        )
        combo.grid(row=row, column=1, sticky="ew", pady=3)
        combo.bind("<<ComboboxSelected>>", handler)
        return combo

    def _sync_widgets_from_settings(self) -> None:
        s = self.controller.settings
        if s.style_id not in self._style_id_to_title:
            self._style_title_to_id[s.style_id] = s.style_id
            self._style_id_to_title[s.style_id] = s.style_id
            style_values = list(self.style_combo.cget("values"))
            if s.style_id not in style_values:
                self.style_combo.configure(values=[*style_values, s.style_id])
        self.style_var.set(self._style_id_to_title[s.style_id])
        self.numbering_var.set(NUMBERING_LABELS[s.numbering_mode])
        font_values = list(self.font_combo.cget("values"))
        if s.font_name not in font_values:
            self.font_combo.configure(values=sorted({*font_values, s.font_name}))
        self.font_name_var.set(s.font_name)
        size_values = [int(v) for v in self.font_size_combo.cget("values")]
        if s.font_size_pt not in size_values:
            self.font_size_combo.configure(
                values=[str(v) for v in sorted({*size_values, s.font_size_pt})]
            )
        self.font_size_var.set(str(s.font_size_pt))
        self.footer_position_var.set(FOOTER_POSITION_LABELS[s.footer_position])
        self.footer_header_var.set(FOOTER_HEADER_LABELS[s.footer_header])

        bib_font_values = list(self.bib_font_combo.cget("values"))
        if s.bib_font_name not in bib_font_values:
            self.bib_font_combo.configure(
                values=sorted({*bib_font_values, s.bib_font_name})
            )
        self.bib_font_name_var.set(s.bib_font_name)
        bib_size_values = [int(v) for v in self.bib_font_size_combo.cget("values")]
        if s.bib_font_size_pt not in bib_size_values:
            self.bib_font_size_combo.configure(
                values=[str(v) for v in sorted({*bib_size_values, s.bib_font_size_pt})]
            )
        self.bib_font_size_var.set(str(s.bib_font_size_pt))
        self.bib_fit_var.set(BIB_FIT_LABELS[s.bib_fit_mode])
        self.bib_columns_var.set(str(s.bib_columns))

        self.always_on_top_var.set(s.always_on_top)
        self.dark_mode_var.set(s.dark_mode)
        if self._applied_dark_mode != s.dark_mode:
            self._apply_theme(s.dark_mode)

    def _schedule_presentation_poll(self) -> None:
        if self._poll_interval_ms is not None and self._poll_interval_ms > 0:
            try:
                self._poll_after_id = self.root.after(
                    self._poll_interval_ms, self._on_presentation_poll
                )
            except tk.TclError:
                self._poll_after_id = None

    def _on_presentation_poll(self) -> None:
        self._poll_after_id = None
        try:
            self.check_active_presentation()
        except Exception:
            pass
        self._schedule_presentation_poll()

    def check_active_presentation(self) -> bool:
        """Sprawdza aktywną prezentację i synchronizuje kontrolki okienka po przełączeniu (D19)."""
        changed = self.controller.sync_with_active_presentation()
        if changed:
            self._sync_widgets_from_settings()
            self.status_var.set(self.controller.status_message)
        return changed

    def _apply_topmost(self, enabled: bool) -> None:
        self.root.wm_attributes("-topmost", bool(enabled))

    def _apply_theme(self, dark: bool) -> None:
        """Stosuje motyw ciemny lub jasny dla wszystkich kontrolek okienka."""
        self._applied_dark_mode = bool(dark)
        palette = DARK_THEME_COLORS if dark else LIGHT_THEME_COLORS

        try:
            if dark and "clam" in self._ttk_style.theme_names():
                self._ttk_style.theme_use("clam")
            elif not dark and self._default_ttk_theme:
                self._ttk_style.theme_use(self._default_ttk_theme)
        except tk.TclError:
            pass

        self.root.configure(bg=palette["bg"])
        self.results_listbox.configure(
            bg=palette["field_bg"],
            fg=palette["fg"],
            selectbackground=palette["select_bg"],
            selectforeground=palette["select_fg"],
            highlightbackground=palette["border"],
            highlightcolor=palette["select_bg"],
        )

        self._ttk_style.configure("Azp.TFrame", background=palette["bg"])
        self._ttk_style.configure(
            "Azp.TLabelframe",
            background=palette["bg"],
            foreground=palette["fg"],
            bordercolor=palette["border"],
        )
        self._ttk_style.configure(
            "Azp.TLabelframe.Label",
            background=palette["bg"],
            foreground=palette["fg"],
        )
        self._ttk_style.configure(
            "Azp.TLabel",
            background=palette["bg"],
            foreground=palette["fg"],
        )
        self._ttk_style.configure(
            "AzpStatus.TLabel",
            background=palette["status_bg"],
            foreground=palette["fg"],
        )
        self._ttk_style.configure(
            "Azp.TCheckbutton",
            background=palette["bg"],
            foreground=palette["fg"],
        )
        self._ttk_style.map(
            "Azp.TCheckbutton",
            background=[("active", palette["bg"])],
            foreground=[("active", palette["fg"])],
        )
        self._ttk_style.configure(
            "Azp.TButton",
            background=palette["button_bg"],
            foreground=palette["fg"],
            bordercolor=palette["border"],
        )
        self._ttk_style.map(
            "Azp.TButton",
            background=[("active", palette["button_active_bg"])],
            foreground=[("active", palette["fg"])],
        )
        self._ttk_style.configure(
            "Azp.TEntry",
            fieldbackground=palette["field_bg"],
            foreground=palette["fg"],
            insertcolor=palette["fg"],
            bordercolor=palette["border"],
        )
        self._ttk_style.configure(
            "Azp.TCombobox",
            fieldbackground=palette["field_bg"],
            background=palette["button_bg"],
            foreground=palette["fg"],
            arrowcolor=palette["fg"],
            bordercolor=palette["border"],
        )
        self._ttk_style.map(
            "Azp.TCombobox",
            fieldbackground=[("readonly", palette["field_bg"])],
            selectbackground=[("readonly", palette["field_bg"])],
            selectforeground=[("readonly", palette["fg"])],
            foreground=[("readonly", palette["fg"])],
            background=[("readonly", palette["button_bg"])],
        )

        try:
            self.root.option_add("*TCombobox*Listbox.background", palette["field_bg"])
            self.root.option_add("*TCombobox*Listbox.foreground", palette["fg"])
            self.root.option_add(
                "*TCombobox*Listbox.selectBackground", palette["select_bg"]
            )
            self.root.option_add(
                "*TCombobox*Listbox.selectForeground", palette["select_fg"]
            )
        except tk.TclError:
            pass

        # Pasek tytułowy Windows 10/11 (DWMWA_USE_IMMERSIVE_DARK_MODE = 20)
        try:
            hwnd = int(self.root.winfo_id())
            value = ctypes.c_int(1 if dark else 0)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 20, ctypes.byref(value), ctypes.sizeof(value)
            )
        except Exception:
            pass

    def _on_collection_selected(self, _event: tk.Event[ttk.Combobox]) -> None:
        self.set_collection(self.collection_var.get())

    def _on_search_event(self, _event: tk.Event[ttk.Entry]) -> None:
        self.perform_search()

    def _on_search_down_arrow(self, _event: tk.Event[ttk.Entry]) -> str:
        if self.controller.results:
            self.results_listbox.focus_set()
            if not self.results_listbox.curselection():
                self.results_listbox.selection_set(0)
                self.results_listbox.activate(0)
        return "break"

    def _on_cite_shortcut(self, _event: tk.Event[tk.Misc]) -> str:
        self.cite_selected()
        return "break"

    def _on_refresh_shortcut(self, _event: tk.Event[tk.Misc]) -> str:
        self.refresh_presentation()
        return "break"

    def _on_focus_search_shortcut(self, _event: tk.Event[tk.Misc]) -> str:
        self.focus_search_entry()
        return "break"

    def _on_style_selected(self, _event: tk.Event[ttk.Combobox]) -> None:
        title = self.style_var.get()
        style_id = self._style_title_to_id[title]
        self._apply_setting_change(style_id=style_id)

    def _on_numbering_selected(self, _event: tk.Event[ttk.Combobox]) -> None:
        mode = LABEL_TO_NUMBERING[self.numbering_var.get()]
        self._apply_setting_change(numbering_mode=mode)

    def _on_font_selected(self, _event: tk.Event[ttk.Combobox]) -> None:
        font_name = self.font_name_var.get()
        self._apply_setting_change(font_name=font_name)

    def _on_font_size_selected(self, _event: tk.Event[ttk.Combobox]) -> None:
        size_pt = int(self.font_size_var.get())
        self._apply_setting_change(font_size_pt=size_pt)

    def _on_footer_position_selected(self, _event: tk.Event[ttk.Combobox]) -> None:
        position = LABEL_TO_FOOTER_POSITION[self.footer_position_var.get()]
        self._apply_setting_change(footer_position=position)

    def _on_footer_header_selected(self, _event: tk.Event[ttk.Combobox]) -> None:
        enabled = LABEL_TO_FOOTER_HEADER[self.footer_header_var.get()]
        self._apply_setting_change(footer_header=enabled)

    def _on_bib_font_selected(self, _event: tk.Event[ttk.Combobox]) -> None:
        font_name = self.bib_font_name_var.get()
        self._apply_setting_change(bib_font_name=font_name)

    def _on_bib_font_size_selected(self, _event: tk.Event[ttk.Combobox]) -> None:
        size_pt = int(self.bib_font_size_var.get())
        self._apply_setting_change(bib_font_size_pt=size_pt)

    def _on_bib_fit_selected(self, _event: tk.Event[ttk.Combobox]) -> None:
        mode = LABEL_TO_BIB_FIT[self.bib_fit_var.get()]
        self._apply_setting_change(bib_fit_mode=mode)

    def _on_bib_columns_selected(self, _event: tk.Event[ttk.Combobox]) -> None:
        columns = int(self.bib_columns_var.get())
        self._apply_setting_change(bib_columns=columns)

    def _on_always_on_top_toggled(self) -> None:
        enabled = bool(self.always_on_top_var.get())
        self._apply_topmost(enabled)
        self.controller.update_settings(rebuild=False, always_on_top=enabled)
        self.status_var.set(
            "Okienko zawsze na wierzchu: włączone."
            if enabled
            else "Okienko zawsze na wierzchu: wyłączone."
        )

    def _on_dark_mode_toggled(self) -> None:
        enabled = bool(self.dark_mode_var.get())
        self._apply_theme(enabled)
        self.controller.update_settings(rebuild=False, dark_mode=enabled)
        self.status_var.set(
            "Tryb ciemny: włączony."
            if enabled
            else "Tryb ciemny: wyłączony."
        )

    def _apply_setting_change(self, **changes: object) -> bool:
        ok = self.controller.update_settings(rebuild=True, **changes)
        self._sync_widgets_from_settings()
        self.status_var.set(self.controller.status_message)
        return ok

    def _rebuild_collection_mappings(self) -> list[str]:
        """Buduje unikalne etykiety kolekcji dla Comboboksa oraz słowniki mapowań (D25, H1)."""
        collections = self.controller.get_collections()
        self._loaded_collections = tuple(collections)
        self._collection_label_to_key = {}
        self._collection_key_to_label = {}
        self._collection_name_to_key = {}

        name_counts: dict[str, int] = {}
        for col in collections:
            name_counts[col.name] = name_counts.get(col.name, 0) + 1

        labels: list[str] = [ALL_COLLECTIONS_LABEL]
        for col in collections:
            if name_counts[col.name] > 1 or col.name == ALL_COLLECTIONS_LABEL:
                label = f"{col.name} ({col.key})"
            else:
                label = col.name
                self._collection_name_to_key[col.name] = col.key
            self._collection_label_to_key[label] = col.key
            self._collection_key_to_label[col.key] = label
            labels.append(label)
        return labels

    def _refresh_collection_combo_if_changed(self) -> None:
        current = tuple(self.controller.get_collections())
        if current != self._loaded_collections:
            values = self._rebuild_collection_mappings()
            self.collection_combo.configure(values=values)
            sel_key = self.controller.selected_collection_key
            if sel_key and sel_key in self._collection_key_to_label:
                self.collection_var.set(self._collection_key_to_label[sel_key])
            elif sel_key is None:
                self.collection_var.set(ALL_COLLECTIONS_LABEL)

    # --- Publiczne API okienka (dla UI, skrótów oraz testów) ---

    def set_collection(self, key_or_name: str | None) -> list[ZoteroItem]:
        """Ustawia filtr kolekcji Zotero (po kluczu, etykiecie lub nazwie) i odświeża listę wyników (D25)."""
        cleaned = (key_or_name or "").strip()
        if not cleaned or cleaned == ALL_COLLECTIONS_LABEL:
            collection_key: str | None = None
            self.collection_var.set(ALL_COLLECTIONS_LABEL)
        elif cleaned in self._collection_label_to_key:
            collection_key = self._collection_label_to_key[cleaned]
            self.collection_var.set(cleaned)
        elif cleaned in self._collection_key_to_label:
            collection_key = cleaned
            self.collection_var.set(self._collection_key_to_label[cleaned])
        elif cleaned in self._collection_name_to_key:
            collection_key = self._collection_name_to_key[cleaned]
            self.collection_var.set(self._collection_key_to_label[collection_key])
        else:
            collection_key = cleaned
            self.collection_var.set(cleaned)

        items = self.controller.set_collection(collection_key)
        self._refresh_collection_combo_if_changed()
        self.results_listbox.delete(0, tk.END)
        for item in items:
            self.results_listbox.insert(tk.END, item.display_label)
        self.status_var.set(self.controller.status_message)
        return items

    def get_selected_collection_key(self) -> str | None:
        """Zwraca klucz aktualnie wybranej kolekcji lub `None` dla całej biblioteki (D25)."""
        return self.controller.selected_collection_key

    def focus_search_entry(self) -> None:
        """Ustawia fokus na polu wyszukiwania i zaznacza cały wpisany tekst (D15)."""
        self.search_entry.focus_set()
        self.search_entry.selection_range(0, tk.END)
        self.search_entry.icursor(tk.END)

    def set_search_query(self, query: str) -> None:
        """Ustawia tekst zapytania w polu wyszukiwania."""
        self.query_var.set(query)

    def perform_search(self, query: str | None = None) -> list[ZoteroItem]:
        """Wykonuje wyszukiwanie w Zotero i wypełnia listę wyników."""
        if query is not None:
            self.query_var.set(query)
        items = self.controller.search(self.query_var.get())
        self._refresh_collection_combo_if_changed()
        self.results_listbox.delete(0, tk.END)
        for item in items:
            self.results_listbox.insert(tk.END, item.display_label)
        self.status_var.set(self.controller.status_message)
        return items

    def get_result_items(self) -> list[ZoteroItem]:
        """Zwraca aktualnie wyświetlone pozycje z Zotero."""
        return list(self.controller.results)

    def select_result_indices(self, indices: Sequence[int]) -> None:
        """Programowo zaznacza pozycje na liście wyników według indeksów 0-based."""
        self.results_listbox.selection_clear(0, tk.END)
        count = len(self.controller.results)
        for idx in indices:
            if 0 <= idx < count:
                self.results_listbox.selection_set(idx)

    def select_result_keys(self, keys: Sequence[str]) -> None:
        """Programowo zaznacza pozycje na liście wyników według kluczy Zotero."""
        wanted = set(keys)
        indices = [
            idx
            for idx, item in enumerate(self.controller.results)
            if item.key in wanted
        ]
        self.select_result_indices(indices)

    def get_selected_indices(self) -> tuple[int, ...]:
        """Zwraca indeksy zaznaczonych pozycji na liście wyników."""
        return tuple(int(i) for i in self.results_listbox.curselection())

    def get_selected_keys(self) -> tuple[str, ...]:
        """Zwraca klucze Zotero zaznaczonych pozycji na liście wyników."""
        items = self.controller.results
        return tuple(
            items[idx].key for idx in self.get_selected_indices() if 0 <= idx < len(items)
        )

    def cite_selected(self) -> bool:
        """Wstawia cytowanie dla zaznaczonych pozycji na liście wyników."""
        self.check_active_presentation()
        ok = self.controller.cite_keys(self.get_selected_keys())
        self._sync_widgets_from_settings()
        self.status_var.set(self.controller.status_message)
        return ok

    def insert_bibliography(self) -> bool:
        """Wstawia lub aktualizuje spis wykorzystanych źródeł na aktywnym slajdzie (D27)."""
        self.check_active_presentation()
        ok = self.controller.insert_bibliography()
        self._sync_widgets_from_settings()
        self.status_var.set(self.controller.status_message)
        return ok

    def refresh_presentation(self) -> bool:
        """Tworzy kopię zapasową i przebudowuje całą prezentację."""
        self.check_active_presentation()
        ok = self.controller.refresh_presentation()
        self._sync_widgets_from_settings()
        self.status_var.set(self.controller.status_message)
        return ok

    def set_style(self, style_id: str) -> bool:
        """Zmienia styl cytowania i natychmiast przebudowuje prezentację (D14)."""
        if style_id not in self._style_id_to_title:
            raise ValueError(f"Nieznany identyfikator stylu: {style_id!r}")
        self.style_var.set(self._style_id_to_title[style_id])
        return self._apply_setting_change(style_id=style_id)

    def set_numbering_mode(self, mode: NumberingMode) -> bool:
        """Zmienia tryb numeracji i natychmiast przebudowuje prezentację (D14)."""
        if mode not in NUMBERING_LABELS:
            raise ValueError(f"Nieprawidłowy tryb numeracji: {mode!r}")
        self.numbering_var.set(NUMBERING_LABELS[mode])
        return self._apply_setting_change(numbering_mode=mode)

    def set_font_name(self, font_name: str) -> bool:
        """Zmienia czcionkę stopki i natychmiast przebudowuje prezentację (D14)."""
        if not font_name.strip():
            raise ValueError("Nazwa czcionki nie może być pusta.")
        values = list(self.font_combo.cget("values"))
        if font_name not in values:
            self.font_combo.configure(values=sorted({*values, font_name}))
        self.font_name_var.set(font_name)
        return self._apply_setting_change(font_name=font_name)

    def set_font_size(self, size_pt: int) -> bool:
        """Zmienia rozmiar czcionki stopki i natychmiast przebudowuje prezentację (D14)."""
        if size_pt < 6 or size_pt > 72:
            raise ValueError(f"Rozmiar czcionki poza zakresem 6..72 pt: {size_pt}")
        values = [int(v) for v in self.font_size_combo.cget("values")]
        if size_pt not in values:
            self.font_size_combo.configure(
                values=[str(s) for s in sorted({*values, size_pt})]
            )
        self.font_size_var.set(str(size_pt))
        return self._apply_setting_change(font_size_pt=size_pt)

    def set_footer_position(self, position: FooterPosition) -> bool:
        """Zmienia położenie pola źródeł i natychmiast przebudowuje prezentację (D9, D14)."""
        if position not in FOOTER_POSITION_LABELS:
            raise ValueError(f"Nieprawidłowe położenie stopki: {position!r}")
        self.footer_position_var.set(FOOTER_POSITION_LABELS[position])
        return self._apply_setting_change(footer_position=position)

    def set_footer_header(self, enabled: bool) -> bool:
        """Zmienia obecność nagłówka „Źródła:” i natychmiast przebudowuje prezentację (D9, D14)."""
        self.footer_header_var.set(FOOTER_HEADER_LABELS[bool(enabled)])
        return self._apply_setting_change(footer_header=bool(enabled))

    def set_bib_font_name(self, font_name: str) -> bool:
        """Zmienia czcionkę bibliografii i natychmiast przebudowuje prezentację (D14, D27)."""
        if not font_name.strip():
            raise ValueError("Nazwa czcionki bibliografii nie może być pusta.")
        values = list(self.bib_font_combo.cget("values"))
        if font_name not in values:
            self.bib_font_combo.configure(values=sorted({*values, font_name}))
        self.bib_font_name_var.set(font_name)
        return self._apply_setting_change(bib_font_name=font_name)

    def set_bib_font_size(self, size_pt: int) -> bool:
        """Zmienia rozmiar czcionki bibliografii i natychmiast przebudowuje prezentację (D14, D27)."""
        if size_pt < 6 or size_pt > 72:
            raise ValueError(
                f"Rozmiar czcionki bibliografii poza zakresem 6..72 pt: {size_pt}"
            )
        values = [int(v) for v in self.bib_font_size_combo.cget("values")]
        if size_pt not in values:
            self.bib_font_size_combo.configure(
                values=[str(s) for s in sorted({*values, size_pt})]
            )
        self.bib_font_size_var.set(str(size_pt))
        return self._apply_setting_change(bib_font_size_pt=size_pt)

    def set_bib_fit_mode(self, mode: BibFitMode) -> bool:
        """Zmienia tryb dopasowania bibliografii i natychmiast przebudowuje prezentację (D14, D27)."""
        if mode not in BIB_FIT_LABELS:
            raise ValueError(
                f"Nieprawidłowy tryb dopasowania bibliografii: {mode!r}"
            )
        self.bib_fit_var.set(BIB_FIT_LABELS[mode])
        return self._apply_setting_change(bib_fit_mode=mode)

    def set_bib_columns(self, columns: int) -> bool:
        """Zmienia liczbę kolumn bibliografii i natychmiast przebudowuje prezentację (D14, D27)."""
        if (
            not isinstance(columns, int)
            or isinstance(columns, bool)
            or columns not in VALID_BIB_COLUMNS
        ):
            raise ValueError(
                f"Liczba kolumn bibliografii poza zakresem {VALID_BIB_COLUMNS}: {columns!r}"
            )
        self.bib_columns_var.set(str(columns))
        return self._apply_setting_change(bib_columns=columns)

    def set_always_on_top(self, enabled: bool) -> None:
        """Włącza lub wyłącza tryb „Zawsze na wierzchu”."""
        self.always_on_top_var.set(bool(enabled))
        self._on_always_on_top_toggled()

    def set_dark_mode(self, enabled: bool) -> None:
        """Włącza lub wyłącza tryb ciemny okienka."""
        self.dark_mode_var.set(bool(enabled))
        self._on_dark_mode_toggled()

    def is_dark_mode(self) -> bool:
        """Zwraca `True`, jeśli w okienku aktywny jest tryb ciemny."""
        return bool(self.dark_mode_var.get())

    def get_status_message(self) -> str:
        """Zwraca bieżący komunikat wyświetlany na pasku stanu."""
        return self.status_var.get()

    def update_idletasks(self) -> None:
        """Odświeża oczekujące zadania tkinter bez blokującej pętli zdarzeń."""
        self.root.update_idletasks()

    def run(self) -> None:
        """Uruchamia główną pętlę zdarzeń tkinter."""
        self.root.mainloop()

    def destroy(self) -> None:
        """Bezpiecznie zamyka okienko."""
        if self._poll_after_id is not None:
            try:
                self.root.after_cancel(self._poll_after_id)
            except tk.TclError:
                pass
            self._poll_after_id = None
        try:
            self.root.destroy()
        except tk.TclError:
            pass
