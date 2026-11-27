"""Pływające okienko tkinter do cytowania pozycji ze Zotero w PowerPoincie (etap B1.2)."""

from __future__ import annotations

import tkinter as tk
import uuid
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
    AzpError,
    BackupError,
    Citation,
    FooterEntry,
    FooterPosition,
    FormattedRef,
    NoActiveSlideError,
    NoTextCursorError,
    NumberingMode,
    Settings,
    SlideDoc,
    SlideRecord,
    StyleInfo,
    ZoteroItem,
    ZoteroOfflineError,
    ZoteroSource,
)

CiteCallback = Callable[[tuple[str, ...], Settings], str | None]
RebuildCallback = Callable[[Settings], str | None]

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

DEFAULT_FONT_SIZES: tuple[int, ...] = (8, 9, 10, 11, 12, 14, 16)

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


def _assign_numeric_map(
    citations: Sequence[Citation],
    extra_keys: Sequence[str] = (),
) -> dict[str, int]:
    """Przypisuje kolejne numery 1..N unikalnym kluczom w kolejności pierwszego wystąpienia."""
    num_map: dict[str, int] = {}
    for cit in citations:
        for key in cit.keys:
            if key not in num_map:
                num_map[key] = len(num_map) + 1
    for key in extra_keys:
        if key not in num_map:
            num_map[key] = len(num_map) + 1
    return num_map


def _strip_outer_parens(text: str) -> str:
    stripped = text.strip()
    if stripped.startswith("(") and stripped.endswith(")") and len(stripped) >= 2:
        return stripped[1:-1].strip()
    return stripped


def _format_in_text_citation(
    keys: Sequence[str],
    refs: dict[str, FormattedRef],
    is_numeric: bool,
    num_map: dict[str, int],
) -> str:
    unique_keys = list(dict.fromkeys(keys))
    if is_numeric:
        numbers = sorted({num_map[k] for k in unique_keys})
        return f"[{', '.join(str(n) for n in numbers)}]"
    parts = [_strip_outer_parens(refs[k].in_text) for k in unique_keys]
    return f"({'; '.join(parts)})"


def _build_footer_entries(
    slide_citations: Sequence[Citation],
    refs: dict[str, FormattedRef],
    is_numeric: bool,
    num_map: dict[str, int],
) -> list[FooterEntry]:
    unique_keys: list[str] = []
    seen: set[str] = set()
    for cit in slide_citations:
        for key in cit.keys:
            if key not in seen:
                seen.add(key)
                unique_keys.append(key)

    if not unique_keys:
        return []

    if is_numeric:
        ordered_keys = sorted(unique_keys, key=lambda k: num_map[k])
        return [
            FooterEntry(
                key=key,
                spans=refs[key].bib_body_spans,
                prefix=f"[{num_map[key]}] ",
            )
            for key in ordered_keys
        ]

    ordered_keys = sorted(unique_keys, key=lambda k: refs[k].effective_sort_key)
    return [
        FooterEntry(
            key=key,
            spans=refs[key].bib_body_spans,
            prefix="",
        )
        for key in ordered_keys
    ]


class AzpController:
    """Kontroler logiki okienka AZP oddzielony od warstwy widgetów tkinter."""

    def __init__(
        self,
        zotero: ZoteroSource,
        slide_doc: SlideDoc,
        settings: Settings | None = None,
        on_cite: CiteCallback | None = None,
        on_rebuild: RebuildCallback | None = None,
        config_path: Path | None = None,
    ) -> None:
        self.zotero = zotero
        self.slide_doc = slide_doc
        self.on_cite = on_cite
        self.on_rebuild = on_rebuild
        self.config_path = config_path

        if settings is not None:
            self.settings = settings
        else:
            base_settings = load_config(config_path) if config_path is not None else Settings()
            loaded = slide_doc.load_settings() if slide_doc.has_open_presentation() else None
            if loaded is not None:
                if config_path is not None:
                    self.settings = replace(loaded, always_on_top=base_settings.always_on_top)
                else:
                    self.settings = loaded
            else:
                self.settings = base_settings

        self.results: list[ZoteroItem] = []
        self.last_query: str = ""
        self.status_message: str = "Gotowy."
        self._styles: list[StyleInfo] = self._load_styles()
        self._active_presentation_key: str | None = self._current_presentation_key()

    def _current_presentation_key(self) -> str | None:
        if not self.slide_doc.has_open_presentation():
            return None
        get_key = getattr(self.slide_doc, "get_presentation_key", None)
        if callable(get_key):
            try:
                key = get_key()
            except AzpError:
                return None
            if key is not None and str(key).strip():
                return str(key).strip()
        try:
            path = self.slide_doc.get_presentation_path()
        except AzpError:
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
        except AzpError:
            loaded = None

        current_topmost = self.settings.always_on_top
        if loaded is not None:
            self.settings = replace(loaded, always_on_top=current_topmost)

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

    def _fetch_formatted_refs(
        self, keys: Sequence[str]
    ) -> dict[str, FormattedRef] | None:
        if not keys:
            return {}
        try:
            return self.zotero.get_formatted_refs(keys, self.settings.style_id)
        except ZoteroOfflineError as exc:
            self.status_message = _format_zotero_offline_message(exc)
            return None
        except KeyError as exc:
            detail = exc.args[0] if exc.args else str(exc)
            self.status_message = f"Nie znaleziono opisu pozycji w Zotero: {detail}"
            return None
        except AzpError as exc:
            self.status_message = str(exc)
            return None

    def get_styles(self) -> list[StyleInfo]:
        """Zwraca listę stylów cytowania dostępnych w Zotero."""
        return list(self._styles)

    def _is_numeric_style(self, style_id: str) -> bool:
        for style in self._styles:
            if style.style_id == style_id:
                return style.is_numeric
        return style_id == STYLE_IEEE_V4

    def search(self, query: str = "", limit: int = 25) -> list[ZoteroItem]:
        """Wyszukuje pozycje w Zotero i aktualizuje listę wyników oraz pasek stanu."""
        self.last_query = query
        try:
            items = self.zotero.search_items(query=query, limit=limit)
        except ZoteroOfflineError as exc:
            self.results = []
            self.status_message = _format_zotero_offline_message(exc)
            return []

        self.results = list(items)
        if not self.results:
            self.status_message = "Brak wyników w Zotero dla podanego zapytania."
        else:
            self.status_message = f"Znaleziono pozycji: {len(self.results)}."
        return list(self.results)

    def cite_keys(self, keys: Sequence[str]) -> bool:
        """Wstawia cytowanie dla wybranych kluczy na aktywnym slajdzie."""
        self.sync_with_active_presentation()
        selected_keys = tuple(dict.fromkeys(keys))
        if not selected_keys:
            self.status_message = "Wybierz pozycję do zacytowania."
            return False

        if not self.slide_doc.has_open_presentation():
            self.status_message = "Brak otwartej prezentacji w PowerPoincie."
            return False

        # D10: zanim cokolwiek zmienimy w prezentacji, wymagamy kursora w polu tekstowym
        if not self.slide_doc.has_text_cursor():
            self.status_message = "Ustaw kursor w polu tekstowym na slajdzie."
            return False

        if self.on_cite is not None:
            try:
                msg = self.on_cite(selected_keys, self.settings)
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

            self.status_message = msg or "Wstawiono cytowanie."
            return True

        return self._cite_builtin(selected_keys)

    def _cite_builtin(self, selected_keys: tuple[str, ...]) -> bool:
        try:
            active_idx = self.slide_doc.get_active_slide_index()
            slides = self.slide_doc.read_slides()
        except NoActiveSlideError:
            self.status_message = "Brak otwartej prezentacji lub aktywnego slajdu."
            return False
        except AzpError as exc:
            self.status_message = str(exc)
            return False

        active_slide = next((s for s in slides if s.slide_index == active_idx), None)
        if active_slide is None:
            self.status_message = "Nie znaleziono aktywnego slajdu w prezentacji."
            return False

        is_numeric = self._is_numeric_style(self.settings.style_id)
        if is_numeric and self.settings.numbering_mode == "global":
            existing_citations = [c for s in slides for c in s.citations]
        else:
            existing_citations = list(active_slide.citations)

        num_map = _assign_numeric_map(existing_citations, selected_keys)

        keys_needed: list[str] = list(
            dict.fromkeys(
                [k for cit in active_slide.citations for k in cit.keys] + list(selected_keys)
            )
        )
        refs = self._fetch_formatted_refs(keys_needed)
        if refs is None:
            return False

        citation_text = _format_in_text_citation(selected_keys, refs, is_numeric, num_map)
        new_citation = Citation(
            citation_id=f"cit-{uuid.uuid4().hex[:12]}",
            keys=selected_keys,
            text=citation_text,
        )

        try:
            self.slide_doc.insert_citation_at_cursor(new_citation)
            updated_slide_citations = [*active_slide.citations, new_citation]
            footer_entries = _build_footer_entries(
                updated_slide_citations, refs, is_numeric, num_map
            )
            self.slide_doc.set_slide_footer(active_idx, footer_entries, self.settings)
            self.slide_doc.save_settings(self.settings)
        except NoTextCursorError:
            self.status_message = "Ustaw kursor w polu tekstowym na slajdzie."
            return False
        except NoActiveSlideError:
            self.status_message = "Brak otwartej prezentacji lub aktywnego slajdu."
            return False
        except AzpError as exc:
            self.status_message = str(exc)
            return False

        self.status_message = f"Wstawiono cytowanie {citation_text} na slajdzie {active_idx}."
        return True

    def refresh_presentation(self) -> bool:
        """Tworzy kopię zapasową i przebudowuje cytowania oraz stopki na wszystkich slajdach."""
        self.sync_with_active_presentation()
        if not self.slide_doc.has_open_presentation():
            self.status_message = "Brak otwartej prezentacji w PowerPoincie."
            return False

        if self.slide_doc.get_presentation_path() is None:
            self.status_message = MSG_SAVE_BEFORE_REBUILD
            return False

        if self.on_rebuild is not None:
            try:
                msg = self.on_rebuild(self.settings)
            except ZoteroOfflineError as exc:
                self.status_message = _format_zotero_offline_message(exc)
                return False
            except BackupError as exc:
                self.status_message = str(exc) or MSG_SAVE_BEFORE_REBUILD
                return False
            except AzpError as exc:
                self.status_message = str(exc)
                return False

            self.status_message = msg or "Przebudowano prezentację."
            return True

        return self._rebuild_builtin()

    def _rebuild_builtin(self) -> bool:
        try:
            slides: list[SlideRecord] = self.slide_doc.read_slides()
        except AzpError as exc:
            self.status_message = str(exc)
            return False

        all_keys = list(
            dict.fromkeys(k for s in slides for cit in s.citations for k in cit.keys)
        )

        refs = self._fetch_formatted_refs(all_keys)
        if refs is None:
            return False

        try:
            backup_path = self.slide_doc.create_backup()
            if backup_path is None:
                self.status_message = MSG_SAVE_BEFORE_REBUILD
                return False
            is_numeric = self._is_numeric_style(self.settings.style_id)

            global_num_map: dict[str, int] = {}
            if is_numeric and self.settings.numbering_mode == "global":
                global_num_map = _assign_numeric_map(
                    [cit for s in slides for cit in s.citations]
                )

            for slide in slides:
                if is_numeric:
                    num_map = (
                        global_num_map
                        if self.settings.numbering_mode == "global"
                        else _assign_numeric_map(slide.citations)
                    )
                else:
                    num_map = {}

                updated_citations = [
                    Citation(
                        citation_id=cit.citation_id,
                        keys=cit.keys,
                        text=_format_in_text_citation(cit.keys, refs, is_numeric, num_map),
                    )
                    for cit in slide.citations
                ]
                self.slide_doc.update_slide_citations(slide.slide_index, updated_citations)

                footer_entries = _build_footer_entries(
                    updated_citations, refs, is_numeric, num_map
                )
                self.slide_doc.set_slide_footer(slide.slide_index, footer_entries, self.settings)

            self.slide_doc.save_settings(self.settings)
        except AzpError as exc:
            self.status_message = str(exc)
            return False

        if backup_path is not None:
            self.status_message = (
                f"Przebudowano prezentację (kopia zapasowa: {backup_path.name})."
            )
        else:
            self.status_message = "Przebudowano prezentację."
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
            if self.slide_doc.has_open_presentation():
                self.slide_doc.save_settings(self.settings)
            self._save_config_safe()
            return True

        if not self.slide_doc.has_open_presentation():
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
            config_path=config_path,
        )

        self.root: tk.Tk | tk.Toplevel
        if master is None:
            self.root = tk.Tk()
        else:
            self.root = tk.Toplevel(master)

        self.root.title("Asystent Zotero dla PowerPointa")
        self.root.geometry("430x640")
        self.root.minsize(380, 540)

        self._style_title_to_id: dict[str, str] = {}
        self._style_id_to_title: dict[str, str] = {}
        self._poll_interval_ms: int | None = poll_interval_ms
        self._poll_after_id: str | None = None
        self._build_variables()
        self._build_ui()
        self._sync_widgets_from_settings()
        self._apply_topmost(self.controller.settings.always_on_top)

        if initial_search:
            self.perform_search("")

        self._schedule_presentation_poll()

    def _build_variables(self) -> None:
        self.query_var = tk.StringVar(master=self.root, value="")
        self.style_var = tk.StringVar(master=self.root, value="")
        self.numbering_var = tk.StringVar(master=self.root, value="")
        self.font_name_var = tk.StringVar(master=self.root, value="")
        self.font_size_var = tk.StringVar(master=self.root, value="")
        self.footer_position_var = tk.StringVar(master=self.root, value="")
        self.footer_header_var = tk.StringVar(master=self.root, value="")
        self.always_on_top_var = tk.BooleanVar(
            master=self.root, value=self.controller.settings.always_on_top
        )
        self.status_var = tk.StringVar(
            master=self.root, value=self.controller.status_message
        )

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        # 1. Wyszukiwarka
        search_frame = ttk.LabelFrame(main, text="Wyszukaj w Zotero", padding=8)
        search_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        search_frame.columnconfigure(0, weight=1)

        self.search_entry = ttk.Entry(search_frame, textvariable=self.query_var)
        self.search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.search_entry.bind("<Return>", self._on_search_event)
        self.search_entry.bind("<Control-Return>", self._on_cite_shortcut)

        self.search_button = ttk.Button(
            search_frame, text="Szukaj", command=self.perform_search
        )
        self.search_button.grid(row=0, column=1)

        # 2. Lista wyników (D13: wybór pojedynczy i wielokrotny Ctrl/Shift)
        results_frame = ttk.LabelFrame(
            main, text="Wyniki (Ctrl / Shift – wybór wielu)", padding=8
        )
        results_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 8))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)

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

        scrollbar = ttk.Scrollbar(
            results_frame, orient=tk.VERTICAL, command=self.results_listbox.yview
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.results_listbox.configure(yscrollcommand=scrollbar.set)

        # 3. Główne akcje
        actions_frame = ttk.Frame(main)
        actions_frame.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        actions_frame.columnconfigure(0, weight=1)
        actions_frame.columnconfigure(1, weight=1)

        self.cite_button = ttk.Button(
            actions_frame, text="Cytuj", command=self.cite_selected
        )
        self.cite_button.grid(row=0, column=0, sticky="ew", padx=(0, 4))

        self.refresh_button = ttk.Button(
            actions_frame, text="Odśwież / Przenumeruj", command=self.refresh_presentation
        )
        self.refresh_button.grid(row=0, column=1, sticky="ew", padx=(4, 0))

        # 4. Ustawienia (D2–D5, D9, D14)
        settings_frame = ttk.LabelFrame(
            main, text="Ustawienia cytowań i pola źródeł", padding=8
        )
        settings_frame.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        settings_frame.columnconfigure(1, weight=1)

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

        # 5. Zawsze na wierzchu
        options_frame = ttk.Frame(main)
        options_frame.grid(row=4, column=0, sticky="ew", pady=(0, 8))

        self.always_on_top_check = ttk.Checkbutton(
            options_frame,
            text="Zawsze na wierzchu",
            variable=self.always_on_top_var,
            command=self._on_always_on_top_toggled,
        )
        self.always_on_top_check.pack(anchor="w")

        # 6. Pasek stanu
        self.status_label = ttk.Label(
            main,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor="w",
            padding=(6, 4),
            wraplength=390,
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
        lbl = ttk.Label(parent, text=label)
        lbl.grid(row=row, column=0, sticky="w", padx=(0, 8), pady=3)
        combo = ttk.Combobox(
            parent,
            textvariable=variable,
            values=list(values),
            state="readonly",
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
        self.always_on_top_var.set(s.always_on_top)

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

    def _on_search_event(self, _event: tk.Event[ttk.Entry]) -> None:
        self.perform_search()

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

    def _on_always_on_top_toggled(self) -> None:
        enabled = bool(self.always_on_top_var.get())
        self._apply_topmost(enabled)
        self.controller.update_settings(rebuild=False, always_on_top=enabled)
        self.status_var.set(
            "Okienko zawsze na wierzchu: włączone."
            if enabled
            else "Okienko zawsze na wierzchu: wyłączone."
        )

    def _apply_setting_change(self, **changes: object) -> bool:
        ok = self.controller.update_settings(rebuild=True, **changes)
        self._sync_widgets_from_settings()
        self.status_var.set(self.controller.status_message)
        return ok

    # --- Publiczne API okienka (dla UI, skrótów oraz testów) ---

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

    def set_always_on_top(self, enabled: bool) -> None:
        """Włącza lub wyłącza tryb „Zawsze na wierzchu”."""
        self.always_on_top_var.set(bool(enabled))
        self._on_always_on_top_toggled()

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
