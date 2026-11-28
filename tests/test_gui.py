"""Testy okienka GUI (etap B1.2) z atrapami `FakeZoteroSource` i `FakeSlideDoc`."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import font as tkfont

import pytest

from azp.contracts import (
    STYLE_APA,
    STYLE_IEEE_V4,
    Citation,
    NoTextCursorError,
    Settings,
    ZoteroCollection,
    ZoteroItem,
)
from azp.gui import (
    ALL_COLLECTIONS_LABEL,
    BIB_FIT_LABELS,
    FOOTER_HEADER_LABELS,
    FOOTER_POSITION_LABELS,
    NUMBERING_LABELS,
    AzpController,
    AzpWindow,
)
from tests.fakes.fake_ppt import FakeSlideDoc
from tests.fakes.fake_zotero import FakeZoteroSource


def _make_window(
    tk_root: tk.Tk | None = None,
    zotero: FakeZoteroSource | None = None,
    doc: FakeSlideDoc | None = None,
    settings: Settings | None = None,
    initial_search: bool = True,
    on_cite=None,
    on_rebuild=None,
    on_bibliography=None,
) -> tuple[AzpWindow, FakeZoteroSource, FakeSlideDoc]:
    zot = zotero if zotero is not None else FakeZoteroSource()
    ppt = doc if doc is not None else FakeSlideDoc()
    win = AzpWindow(
        zotero=zot,
        slide_doc=ppt,
        settings=settings,
        master=tk_root,
        on_cite=on_cite,
        on_rebuild=on_rebuild,
        on_bibliography=on_bibliography,
        initial_search=initial_search,
    )
    win.update_idletasks()
    return win, zot, ppt


def test_search_displays_results_and_filters_on_query_and_return_key(tk_root: tk.Tk) -> None:
    win, _zot, _doc = _make_window(tk_root=tk_root)
    try:
        assert win.results_listbox.size() == 4
        assert "Kowalski, Adam (2023)" in win.results_listbox.get(0)

        win.set_search_query("Nowak")
        win.search_button.invoke()
        win.update_idletasks()
        assert win.results_listbox.size() == 1
        assert win.get_result_items()[0].key == "NOWAK2024"

        win.set_search_query("2013")
        win.root.update()
        win.search_entry.focus_force()
        win.root.update()
        win.search_entry.event_generate("<Return>")
        win.root.update()
        assert win.results_listbox.size() == 1
        assert win.get_result_items()[0].key == "DOVE2013"

        win.perform_search("")
        assert win.results_listbox.size() == 4
    finally:
        win.destroy()


def test_zotero_offline_shows_status_without_unhandled_exception(tk_root: tk.Tk) -> None:
    zot = FakeZoteroSource(offline=True)
    win, _zot, doc = _make_window(tk_root=tk_root, zotero=zot)
    try:
        assert "Zotero nie odpowiada" in win.get_status_message()
        assert win.results_listbox.size() == 0

        # Przejście w tryb online na chwilę, wybór i wstawienie pozycji, a potem awaria
        zot.offline = False
        win.perform_search("")
        assert win.results_listbox.size() == 4
        win.select_result_keys(["KOWAL2023"])
        assert win.cite_selected() is True

        zot.offline = True
        win.select_result_keys(["NOWAK2024"])
        assert win.cite_selected() is False
        assert "Zotero nie odpowiada" in win.get_status_message()
        assert len(doc.read_slides()[0].citations) == 1
        assert doc.backups_created == []

        # Zmiana stylu oraz odświeżenie przy wyłączonym Zotero nie rzucają wyjątku
        assert win.set_style(STYLE_APA) is False
        assert "Zotero nie odpowiada" in win.get_status_message()
        assert win.refresh_presentation() is False
        assert "Zotero nie odpowiada" in win.get_status_message()
    finally:
        win.destroy()


def test_single_selection_cite_inserts_citation_and_footer(tk_root: tk.Tk) -> None:
    win, _zot, doc = _make_window(tk_root=tk_root)
    try:
        win.select_result_keys(["KOWAL2023"])
        assert win.cite_selected() is True

        slides = doc.read_slides()
        assert "[1]" in doc.get_slide_body_text(1)
        assert len(slides[0].citations) == 1
        assert slides[0].citations[0].keys == ("KOWAL2023",)
        assert slides[0].citations[0].text == "[1]"

        assert len(slides[0].footer_entries) == 1
        assert slides[0].footer_entries[0].key == "KOWAL2023"
        assert slides[0].footer_entries[0].prefix == "[1] "
        assert "Kowalski" in slides[0].footer_entries[0].plain_text

        saved = doc.load_settings()
        assert saved is not None
        assert saved.style_id == STYLE_IEEE_V4

        # Na drugim slajdzie powtórzenie KOWAL2023 dostaje [1], a nowa pozycja [2]
        doc.active_slide_index = 2
        win.select_result_keys(["KOWAL2023"])
        assert win.cite_selected() is True
        win.select_result_keys(["NOWAK2024"])
        assert win.cite_selected() is True

        slides = doc.read_slides()
        assert [c.text for c in slides[1].citations] == ["[1]", "[2]"]
        assert [e.prefix for e in slides[1].footer_entries] == ["[1] ", "[2] "]
    finally:
        win.destroy()


def test_multi_selection_d13_ieee_and_apa(tk_root: tk.Tk) -> None:
    win, _zot, doc = _make_window(tk_root=tk_root)
    try:
        assert win.results_listbox.cget("selectmode") == tk.EXTENDED
        win.results_listbox.selection_clear(0, tk.END)
        win.results_listbox.selection_set(0, 1)  # KOWAL2023, NOWAK2024

        assert win.cite_selected() is True
        slides = doc.read_slides()
        assert slides[0].citations[0].keys == ("KOWAL2023", "NOWAK2024")
        assert slides[0].citations[0].text == "[1, 2]"
        assert "[1, 2]" in doc.get_slide_body_text(1)
        assert [e.key for e in slides[0].footer_entries] == ["KOWAL2023", "NOWAK2024"]
        assert [e.prefix for e in slides[0].footer_entries] == ["[1] ", "[2] "]
    finally:
        win.destroy()

    # Ten sam wielokrotny wybór w stylu APA
    doc_apa = FakeSlideDoc()
    win_apa, _zot, doc_apa = _make_window(
        tk_root=tk_root, doc=doc_apa, settings=Settings(style_id=STYLE_APA)
    )
    try:
        win_apa.select_result_indices([0, 1])
        assert win_apa.cite_selected() is True
        slides_apa = doc_apa.read_slides()
        assert (
            slides_apa[0].citations[0].text
            == "(Kowalski, 2023; Nowak & Wiśniewski, 2024)"
        )
        assert [e.prefix for e in slides_apa[0].footer_entries] == ["", ""]
        assert [e.key for e in slides_apa[0].footer_entries] == ["KOWAL2023", "NOWAK2024"]
    finally:
        win_apa.destroy()


def test_cite_blocked_without_cursor_d10_and_without_selection(tk_root: tk.Tk) -> None:
    doc = FakeSlideDoc(has_cursor=False)
    win, _zot, doc = _make_window(tk_root=tk_root, doc=doc)
    try:
        before_body = doc.get_slide_body_text(1)

        # Brak wyboru na liście
        assert win.cite_selected() is False
        assert "Wybierz pozycję" in win.get_status_message()
        assert doc.get_slide_body_text(1) == before_body

        # Wybór pozycji bez kursora w polu tekstowym (D26 zastępuje D10):
        # dodaje przypis do stopki na dole slajdu bez modyfikacji body_text
        win.select_result_keys(["KOWAL2023"])
        assert win.cite_selected() is True
        assert doc.get_slide_body_text(1) == before_body
        slides = doc.read_slides()
        assert len(slides[0].citations) == 1
        assert slides[0].citations[0].keys == ("KOWAL2023",)
        assert slides[0].citations[0].in_text is False
        assert len(slides[0].footer_entries) == 1
        assert slides[0].footer_entries[0].key == "KOWAL2023"
        assert slides[0].footer_entries[0].prefix == "[1] "
        assert doc.backups_created == []
    finally:
        win.destroy()


def test_cite_handles_unexpected_no_text_cursor_error_from_slide_doc(tk_root: tk.Tk) -> None:
    class RaisingSlideDoc(FakeSlideDoc):
        def insert_citation_at_cursor(self, citation: Citation) -> None:
            raise NoTextCursorError("Brak kursora")

    doc = RaisingSlideDoc(has_cursor=True)
    win, _zot, doc = _make_window(tk_root=tk_root, doc=doc)
    try:
        win.select_result_keys(["KOWAL2023"])
        assert win.cite_selected() is False
        assert "kursor" in win.get_status_message().casefold()
        assert doc.read_slides()[0].footer_entries == ()
        assert doc.load_settings() is None
    finally:
        win.destroy()


def test_cite_without_open_presentation(tk_root: tk.Tk) -> None:
    doc = FakeSlideDoc(slide_count=0, active_slide_index=None)
    win, _zot, _doc = _make_window(tk_root=tk_root, doc=doc)
    try:
        win.select_result_keys(["KOWAL2023"])
        assert win.cite_selected() is False
        assert "prezentacji" in win.get_status_message().casefold()
    finally:
        win.destroy()


def test_system_fonts_footer_positions_and_header_options(tk_root: tk.Tk) -> None:
    win, _zot, _doc = _make_window(tk_root=tk_root)
    try:
        font_values = list(win.font_combo.cget("values"))
        sys_fonts = {f.strip() for f in tkfont.families(win.root) if f and f.strip()}
        assert sys_fonts.issubset(set(font_values))
        assert "Calibri" in font_values
        assert "Arial" in font_values
        assert len(font_values) == len(set(font_values))
        assert font_values == sorted(font_values)

        pos_values = list(win.footer_position_combo.cget("values"))
        assert pos_values == [
            FOOTER_POSITION_LABELS["bottom_full"],
            FOOTER_POSITION_LABELS["bottom_left"],
            FOOTER_POSITION_LABELS["replace_footer"],
        ]

        header_values = list(win.footer_header_combo.cget("values"))
        assert header_values == [
            FOOTER_HEADER_LABELS[True],
            FOOTER_HEADER_LABELS[False],
        ]

        numbering_values = list(win.numbering_combo.cget("values"))
        assert numbering_values == [
            NUMBERING_LABELS["global"],
            NUMBERING_LABELS["per_slide"],
        ]
    finally:
        win.destroy()


def test_setting_changes_immediately_rebuild_presentation_with_backup_d14(
    tk_root: tk.Tk,
) -> None:
    win, _zot, doc = _make_window(tk_root=tk_root)
    try:
        # Slajd 1: KOWAL2023 -> [1]
        doc.active_slide_index = 1
        win.select_result_keys(["KOWAL2023"])
        assert win.cite_selected() is True

        # Slajd 2: NOWAK2024 i DOVE2013 -> [2], [3]
        doc.active_slide_index = 2
        win.select_result_keys(["NOWAK2024"])
        assert win.cite_selected() is True
        win.select_result_keys(["DOVE2013"])
        assert win.cite_selected() is True
        assert doc.backups_created == []

        # 1. Zmiana trybu numeracji global -> per_slide przez zdarzenie Combobox
        win.numbering_var.set(NUMBERING_LABELS["per_slide"])
        win.numbering_combo.event_generate("<<ComboboxSelected>>")
        win.update_idletasks()
        assert len(doc.backups_created) == 1
        slides = doc.read_slides()
        assert [c.text for c in slides[1].citations] == ["[1]", "[2]"]
        assert [e.prefix for e in slides[1].footer_entries] == ["[1] ", "[2] "]
        assert doc.load_settings().numbering_mode == "per_slide"

        # 2. Zmiana stylu IEEE v4 -> APA (sortowanie alfabetyczne: Dove przed Nowak)
        win.style_var.set("APA Style 7th edition")
        win.style_combo.event_generate("<<ComboboxSelected>>")
        win.update_idletasks()
        assert len(doc.backups_created) == 2
        slides = doc.read_slides()
        assert slides[0].citations[0].text == "(Kowalski, 2023)"
        assert [c.text for c in slides[1].citations] == [
            "(Nowak & Wiśniewski, 2024)",
            "(Dove et al., 2013)",
        ]
        assert [e.key for e in slides[1].footer_entries] == ["DOVE2013", "NOWAK2024"]
        assert [e.prefix for e in slides[1].footer_entries] == ["", ""]

        # 3. Zmiana czcionki -> Arial
        assert win.set_font_name("Arial") is True
        assert len(doc.backups_created) == 3
        assert doc.get_slide_footer_settings(1).font_name == "Arial"

        # 4. Zmiana rozmiaru czcionki -> 12
        assert win.set_font_size(12) is True
        assert len(doc.backups_created) == 4
        assert doc.get_slide_footer_settings(1).font_size_pt == 12

        # 5. Zmiana położenia pola źródeł -> bottom_left -> replace_footer
        assert win.set_footer_position("bottom_left") is True
        assert len(doc.backups_created) == 5
        assert doc.get_slide_footer_settings(1).footer_position == "bottom_left"

        win.footer_position_var.set(FOOTER_POSITION_LABELS["replace_footer"])
        win.footer_position_combo.event_generate("<<ComboboxSelected>>")
        win.update_idletasks()
        assert len(doc.backups_created) == 6
        assert doc.get_slide_footer_settings(1).footer_position == "replace_footer"

        # 6. Przełącznik nagłówka „Źródła:”
        assert win.set_footer_header(False) is True
        assert len(doc.backups_created) == 7
        assert doc.get_slide_footer_settings(1).footer_header is False
    finally:
        win.destroy()


def test_refresh_button_renumbers_after_slide_reorder(tk_root: tk.Tk) -> None:
    win, _zot, doc = _make_window(tk_root=tk_root)
    try:
        doc.active_slide_index = 1
        win.select_result_keys(["KOWAL2023"])
        assert win.cite_selected() is True

        doc.active_slide_index = 2
        win.select_result_keys(["NOWAK2024"])
        assert win.cite_selected() is True

        # Przestawiamy slajdy: dawny slajd 2 staje się slajdem 1
        doc.reorder_slides([2, 1, 3])
        win.refresh_button.invoke()
        win.update_idletasks()

        assert len(doc.backups_created) == 1
        slides = doc.read_slides()
        assert slides[0].citations[0].keys == ("NOWAK2024",)
        assert slides[0].citations[0].text == "[1]"
        assert slides[0].footer_entries[0].prefix == "[1] "

        assert slides[1].citations[0].keys == ("KOWAL2023",)
        assert slides[1].citations[0].text == "[2]"
        assert slides[1].footer_entries[0].prefix == "[2] "
    finally:
        win.destroy()


def test_always_on_top_toggle_updates_topmost_without_rebuild(tk_root: tk.Tk) -> None:
    win, _zot, doc = _make_window(tk_root=tk_root)
    try:
        assert bool(win.root.wm_attributes("-topmost")) is True
        assert win.controller.settings.always_on_top is True

        win.always_on_top_check.invoke()
        win.update_idletasks()
        assert bool(win.root.wm_attributes("-topmost")) is False
        assert win.controller.settings.always_on_top is False
        assert doc.load_settings() is not None
        assert doc.load_settings().always_on_top is False
        assert doc.backups_created == []

        win.set_always_on_top(True)
        assert bool(win.root.wm_attributes("-topmost")) is True
        assert win.controller.settings.always_on_top is True
    finally:
        win.destroy()


def test_external_callbacks_are_invoked_when_provided(tk_root: tk.Tk) -> None:
    cite_calls: list[tuple[tuple[str, ...], Settings]] = []
    rebuild_calls: list[Settings] = []

    def fake_on_cite(keys: tuple[str, ...], settings: Settings) -> str:
        cite_calls.append((keys, settings))
        return "Zewnętrzne cytowanie OK"

    def fake_on_rebuild(settings: Settings) -> str:
        rebuild_calls.append(settings)
        return "Zewnętrzna przebudowa OK"

    win, _zot, _doc = _make_window(
        tk_root=tk_root, on_cite=fake_on_cite, on_rebuild=fake_on_rebuild
    )
    try:
        win.select_result_keys(["KOWAL2023", "ZIEL2022"])
        assert win.cite_selected() is True
        assert len(cite_calls) == 1
        assert cite_calls[0][0] == ("KOWAL2023", "ZIEL2022")
        assert win.get_status_message() == "Zewnętrzne cytowanie OK"

        assert win.refresh_presentation() is True
        assert len(rebuild_calls) == 1
        assert win.get_status_message() == "Zewnętrzna przebudowa OK"
    finally:
        win.destroy()


def test_loads_existing_presentation_settings_on_startup(tk_root: tk.Tk) -> None:
    doc = FakeSlideDoc()
    saved = Settings(
        style_id=STYLE_APA,
        numbering_mode="per_slide",
        font_name="Arial",
        font_size_pt=12,
        footer_position="bottom_left",
        footer_header=False,
        always_on_top=False,
    )
    doc.save_settings(saved)

    win, _zot, _doc = _make_window(tk_root=tk_root, doc=doc)
    try:
        assert win.controller.settings == saved
        assert win.style_var.get() == "APA Style 7th edition"
        assert win.numbering_var.get() == NUMBERING_LABELS["per_slide"]
        assert win.font_name_var.get() == "Arial"
        assert win.font_size_var.get() == "12"
        assert win.footer_position_var.get() == FOOTER_POSITION_LABELS["bottom_left"]
        assert win.footer_header_var.get() == FOOTER_HEADER_LABELS[False]
        assert win.always_on_top_var.get() is False
        assert bool(win.root.wm_attributes("-topmost")) is False
    finally:
        win.destroy()


def test_controller_works_directly_without_window() -> None:
    zot = FakeZoteroSource()
    doc = FakeSlideDoc()
    ctrl = AzpController(zotero=zot, slide_doc=doc)

    items = ctrl.search("Zieliński")
    assert [i.key for i in items] == ["ZIEL2022"]
    assert ctrl.cite_keys(["ZIEL2022"]) is True
    assert doc.read_slides()[0].citations[0].text == "[1]"
    assert ctrl.refresh_presentation() is True
    assert len(doc.backups_created) == 1


def test_failed_rebuild_rolls_back_settings_and_combobox_h1(tk_root: tk.Tk) -> None:
    zot = FakeZoteroSource()
    win, zot, doc = _make_window(tk_root=tk_root, zotero=zot)
    try:
        win.select_result_keys(["KOWAL2023"])
        assert win.cite_selected() is True

        zot.offline = True
        assert win.set_style(STYLE_APA) is False
        assert "Zotero nie odpowiada" in win.get_status_message()
        assert win.controller.settings.style_id == STYLE_IEEE_V4
        assert win.style_var.get() == "IEEE Praca Inzynierska Access Date v4"

        zot.offline = False
        assert win.set_style(STYLE_APA) is True
        assert win.controller.settings.style_id == STYLE_APA
        assert win.style_var.get() == "APA Style 7th edition"
        assert doc.read_slides()[0].citations[0].text == "(Kowalski, 2023)"
    finally:
        win.destroy()


def test_missing_item_key_error_handled_gracefully_m1(tk_root: tk.Tk) -> None:
    missing_item = ZoteroItem(
        key="MISSING999",
        title="Pozycja bez opisu w Zotero",
        authors=("Anonim, Jan",),
        year="2025",
    )
    zot = FakeZoteroSource(items=[missing_item])
    doc = FakeSlideDoc()
    win, zot, doc = _make_window(tk_root=tk_root, zotero=zot, doc=doc)
    try:
        # 1. Brakujący klucz podczas cytowania
        win.select_result_keys(["MISSING999"])
        assert win.cite_selected() is False
        assert "MISSING999" in win.get_status_message() or "Nie znaleziono" in win.get_status_message()
        assert doc.read_slides()[0].citations == ()

        # 2. Brakujący klucz podczas przebudowy prezentacji
        doc.insert_citation_at_cursor(
            Citation(citation_id="c-missing", keys=("MISSING999",), text="[1]")
        )
        assert win.refresh_presentation() is False
        assert "MISSING999" in win.get_status_message() or "Nie znaleziono" in win.get_status_message()
    finally:
        win.destroy()


# ---------------------------------------------------------------------------
# C1: niezapisana prezentacja (D18) i śledzenie aktywnej prezentacji (D19)
# ---------------------------------------------------------------------------


class _MultiPresentationDoc:
    """Atrapa kilku otwartych prezentacji; aktywną wybiera `activate(name)` (D19)."""

    def __init__(self, docs: dict[str, FakeSlideDoc], active: str | None) -> None:
        self.docs = docs
        self.active = active

    def activate(self, name: str | None) -> None:
        self.active = name

    def get_presentation_key(self) -> str | None:
        return self.active

    def _doc(self) -> FakeSlideDoc:
        if self.active is None:
            return FakeSlideDoc(slide_count=0, active_slide_index=None)
        return self.docs[self.active]

    def __getattr__(self, name: str):  # type: ignore[no-untyped-def]
        return getattr(self._doc(), name)


def test_c1_unsaved_presentation_blocks_rebuild_and_rolls_back_d18(tk_root: tk.Tk) -> None:
    from azp.contracts import MSG_SAVE_BEFORE_REBUILD

    rebuild_calls: list[Settings] = []

    def on_rebuild(settings: Settings) -> str:
        rebuild_calls.append(settings)
        return "nie powinno się wywołać"

    doc = FakeSlideDoc(presentation_path=None)
    for callback in (None, on_rebuild):
        win, _zot, doc = _make_window(
            tk_root=tk_root, doc=FakeSlideDoc(presentation_path=None), on_rebuild=callback
        )
        try:
            win.select_result_keys(["KOWAL2023"])
            assert win.cite_selected() is True  # Cytuj działa normalnie
            text_after_cite = doc.get_slide_body_text(1)

            assert win.set_style(STYLE_APA) is False
            assert win.get_status_message() == MSG_SAVE_BEFORE_REBUILD
            assert win.controller.settings.style_id == STYLE_IEEE_V4
            assert win.style_var.get() == "IEEE Praca Inzynierska Access Date v4"

            assert win.refresh_presentation() is False
            assert win.get_status_message() == MSG_SAVE_BEFORE_REBUILD
            assert doc.get_slide_body_text(1) == text_after_cite
            assert doc.backups_created == []
        finally:
            win.destroy()
    assert rebuild_calls == []


def test_c1_window_follows_active_presentation_settings_d19(tk_root: tk.Tk) -> None:
    doc_a = FakeSlideDoc(presentation_path=Path("C:/fake/A.pptx"))
    saved_a = Settings(style_id=STYLE_IEEE_V4, font_name="Arial", font_size_pt=12)
    doc_a.save_settings(saved_a)
    doc_b = FakeSlideDoc(presentation_path=Path("C:/fake/B.pptx"))
    multi = _MultiPresentationDoc({"A": doc_a, "B": doc_b}, active="A")

    win = AzpWindow(
        zotero=FakeZoteroSource(),
        slide_doc=multi,  # type: ignore[arg-type]
        master=tk_root,
        poll_interval_ms=None,
    )
    try:
        assert win.controller.settings == saved_a
        assert win.font_size_var.get() == "12"

        # B nie ma ustawień AZP -> dostaje ustawienia widoczne w okienku
        multi.activate("B")
        win.check_active_presentation()
        assert win.controller.settings == saved_a
        assert win.set_style(STYLE_APA) is True
        assert doc_b.load_settings() is not None
        assert doc_b.load_settings().style_id == STYLE_APA
        assert doc_a.load_settings() == saved_a  # A nietknięta

        # Powrót do A: okienko wczytuje ustawienia A i odświeża listy
        multi.activate("A")
        win.check_active_presentation()
        assert win.controller.settings.style_id == STYLE_IEEE_V4
        assert win.style_var.get() == "IEEE Praca Inzynierska Access Date v4"
        assert win.font_name_var.get() == "Arial"
        assert "A" in win.get_status_message()

        # Przełączenie bez odpytywania: Cytuj sam sprawdza aktywną prezentację
        multi.activate("B")
        win.select_result_keys(["KOWAL2023"])
        assert win.cite_selected() is True
        assert doc_b.read_slides()[0].citations[0].text == "(Kowalski, 2023)"
        assert win.style_var.get() == "APA Style 7th edition"

        # „Zawsze na wierzchu” to opcja okienka, nie prezentacji – nie zmienia się przy przełączeniu
        win.set_always_on_top(False)
        multi.activate("A")
        win.check_active_presentation()
        assert win.controller.settings.always_on_top is False
        assert win.always_on_top_var.get() is False

        # Zamknięcie wszystkich prezentacji nie zmienia ustawień w okienku
        multi.activate(None)
        win.check_active_presentation()
        assert win.controller.settings.style_id == STYLE_IEEE_V4
    finally:
        win.destroy()


def test_c1_unknown_style_in_presentation_does_not_crash_window_d19(tk_root: tk.Tk) -> None:
    doc_a = FakeSlideDoc(presentation_path=Path("C:/fake/A.pptx"))
    doc_b = FakeSlideDoc(presentation_path=Path("C:/fake/B.pptx"))
    doc_b.save_settings(Settings(style_id="http://www.zotero.org/styles/nieznany"))
    multi = _MultiPresentationDoc({"A": doc_a, "B": doc_b}, active="A")
    win = AzpWindow(
        zotero=FakeZoteroSource(),
        slide_doc=multi,  # type: ignore[arg-type]
        master=tk_root,
        poll_interval_ms=None,
    )
    try:
        multi.activate("B")
        win.check_active_presentation()
        assert win.controller.settings.style_id == "http://www.zotero.org/styles/nieznany"
        assert win.style_var.get() == "http://www.zotero.org/styles/nieznany"
    finally:
        win.destroy()



def test_c1_m1_setting_change_or_refresh_right_after_presentation_switch_syncs_d19(
    tk_root: tk.Tk,
) -> None:
    """M1 (C1): zmiana ustawienia lub refresh bez uprzedniego timera synchronizuje wszystkie kontrolki."""
    doc_a = FakeSlideDoc(presentation_path=Path("C:/fake/A.pptx"))
    doc_a.save_settings(
        Settings(style_id=STYLE_IEEE_V4, font_name="Calibri", font_size_pt=10)
    )
    doc_b = FakeSlideDoc(presentation_path=Path("C:/fake/B.pptx"))
    doc_b.save_settings(
        Settings(style_id=STYLE_IEEE_V4, font_name="Arial", font_size_pt=14)
    )
    multi = _MultiPresentationDoc({"A": doc_a, "B": doc_b}, active="A")

    win = AzpWindow(
        zotero=FakeZoteroSource(),
        slide_doc=multi,  # type: ignore[arg-type]
        master=tk_root,
        poll_interval_ms=None,
    )
    try:
        # 1) Przełączenie na B i zmiana stylu przed wywołaniem check_active_presentation()
        multi.activate("B")
        assert win.set_style(STYLE_APA) is True
        win.check_active_presentation()
        assert win.font_name_var.get() == "Arial"
        assert win.font_size_var.get() == "14"

        # 2) Bezpośrednie controller.refresh_presentation() po przełączeniu na A wczytuje ustawienia A
        multi.activate("A")
        assert win.controller.refresh_presentation() is True
        assert win.controller.settings.style_id == STYLE_IEEE_V4
        assert win.controller.settings.font_name == "Calibri"
    finally:
        win.destroy()


def test_dark_mode_toggle_styles_window_and_persists_without_rebuild(
    tk_root: tk.Tk, tmp_path: Path
) -> None:
    """Przełącznik Tryb ciemny zmienia kolory okienka i zapisuje ustawienie bez przebudowy prezentacji."""
    from azp.config import load_config
    from azp.gui import DARK_THEME_COLORS, LIGHT_THEME_COLORS

    cfg_path = tmp_path / "config.json"
    doc = FakeSlideDoc()
    win = AzpWindow(
        zotero=FakeZoteroSource(),
        slide_doc=doc,
        master=tk_root,
        poll_interval_ms=None,
        config_path=cfg_path,
    )
    try:
        assert win.is_dark_mode() is False
        assert win.results_listbox.cget("bg") == LIGHT_THEME_COLORS["field_bg"]

        win.dark_mode_check.invoke()
        win.update_idletasks()
        assert win.is_dark_mode() is True
        assert win.controller.settings.dark_mode is True
        assert win.results_listbox.cget("bg") == DARK_THEME_COLORS["field_bg"]
        assert win.results_listbox.cget("fg") == DARK_THEME_COLORS["fg"]
        assert win.root.cget("bg") == DARK_THEME_COLORS["bg"]
        assert "Tryb ciemny: włączony" in win.get_status_message()
        assert doc.backups_created == []
        assert load_config(cfg_path).dark_mode is True

        win.set_dark_mode(False)
        win.update_idletasks()
        assert win.is_dark_mode() is False
        assert win.results_listbox.cget("bg") == LIGHT_THEME_COLORS["field_bg"]
        assert "Tryb ciemny: wyłączony" in win.get_status_message()
        assert load_config(cfg_path).dark_mode is False
    finally:
        win.destroy()


def test_double_click_cites_and_down_arrow_moves_focus_to_results(
    tk_root: tk.Tk,
) -> None:
    """Dwuklik na liście wyników wstawia cytowanie, a strzałka w dół z pola szukania wybiera pierwszy wynik."""
    win, _zot, doc = _make_window(tk_root=tk_root)
    try:
        win.root.update()
        win.search_entry.focus_force()
        win.root.update()
        win.search_entry.event_generate("<Down>")
        win.root.update()
        assert win.get_selected_indices() == (0,)

        assert win.results_listbox.bind("<Double-Button-1>")
        win.results_listbox.event_generate("<ButtonPress-1>", x=10, y=10)
        win.results_listbox.event_generate("<ButtonRelease-1>", x=10, y=10)
        win.results_listbox.event_generate("<ButtonPress-1>", x=10, y=10)
        win.results_listbox.event_generate("<ButtonRelease-1>", x=10, y=10)
        win.root.update()
        assert len(doc.read_slides()[0].citations) == 1
    finally:
        win.destroy()


# ---------------------------------------------------------------------------
# D25, D26, D27: Kolekcje Zotero, cytowanie bez kursora i slajd bibliografii
# ---------------------------------------------------------------------------


def test_collection_filter_combobox_and_search_within_collection_d25(
    tk_root: tk.Tk,
) -> None:
    win, _zot, _doc = _make_window(tk_root=tk_root)
    try:
        values = list(win.collection_combo.cget("values"))
        assert values == [
            ALL_COLLECTIONS_LABEL,
            "Sztuczna inteligencja",
            "Sztuczna inteligencja / Uczenie maszynowe",
            "Etyka i prawo",
        ]
        assert win.collection_var.get() == ALL_COLLECTIONS_LABEL
        assert win.get_selected_collection_key() is None
        assert len(win.get_result_items()) == 4

        # Wybór podkolekcji przez zdarzenie Combobox
        win.collection_var.set("Sztuczna inteligencja / Uczenie maszynowe")
        win.collection_combo.event_generate("<<ComboboxSelected>>")
        win.update_idletasks()
        assert win.get_selected_collection_key() == "COL_ML"
        assert [i.key for i in win.get_result_items()] == ["NOWAK2024", "ZIEL2022"]

        # Wyszukiwanie wewnątrz wybranej podkolekcji
        win.perform_search("Zieliński")
        assert [i.key for i in win.get_result_items()] == ["ZIEL2022"]
        win.perform_search("Dove")
        assert win.get_result_items() == []

        # Przełączenie na kolekcję Etyka i prawo po kluczu (z zachowaniem zapytania "Dove")
        win.set_collection("COL_ETH")
        assert win.get_selected_collection_key() == "COL_ETH"
        assert win.collection_var.get() == "Etyka i prawo"
        assert [i.key for i in win.get_result_items()] == ["DOVE2013"]

        # Powrót do całej biblioteki
        win.perform_search("")
        win.set_collection(ALL_COLLECTIONS_LABEL)
        assert win.get_selected_collection_key() is None
        assert len(win.get_result_items()) == 4
    finally:
        win.destroy()


def test_cite_without_text_cursor_inserts_footer_entry_and_renumbers_d26(
    tk_root: tk.Tk,
) -> None:
    doc = FakeSlideDoc(has_cursor=False)
    win, _zot, doc = _make_window(tk_root=tk_root, doc=doc)
    try:
        # Slajd 1: brak kursora (np. slajd ze zdjęciami) -> cytowanie trafia tylko do stopki
        doc.active_slide_index = 1
        body_1_before = doc.get_slide_body_text(1)
        win.select_result_keys(["KOWAL2023"])
        assert win.cite_selected() is True
        assert doc.get_slide_body_text(1) == body_1_before

        win.select_result_keys(["NOWAK2024"])
        assert win.cite_selected() is True
        assert doc.get_slide_body_text(1) == body_1_before

        slides = doc.read_slides()
        assert len(slides[0].citations) == 1
        assert slides[0].citations[0].in_text is False
        assert slides[0].citations[0].keys == ("KOWAL2023", "NOWAK2024")
        assert [e.prefix for e in slides[0].footer_entries] == ["[1] ", "[2] "]

        # Slajd 2: z kursorem -> nowa pozycja dostaje [3], bo [1] i [2] były w stopce slajdu 1
        doc.active_slide_index = 2
        doc.has_cursor = True
        win.select_result_keys(["DOVE2013"])
        assert win.cite_selected() is True
        slides = doc.read_slides()
        assert slides[1].citations[0].in_text is True
        assert slides[1].citations[0].text == "[3]"
        assert "[3]" in doc.get_slide_body_text(2)
    finally:
        win.destroy()


def test_insert_bibliography_button_and_bibliography_settings_d27(
    tk_root: tk.Tk,
) -> None:
    win, _zot, doc = _make_window(tk_root=tk_root)
    try:
        # Cytujemy na slajdach 1 i 2 (w tym jedno bez kursora)
        doc.active_slide_index = 1
        doc.has_cursor = True
        win.select_result_keys(["KOWAL2023", "NOWAK2024"])
        assert win.cite_selected() is True

        doc.active_slide_index = 2
        doc.has_cursor = False
        win.select_result_keys(["DOVE2013", "ZIEL2022"])
        assert win.cite_selected() is True

        # Slajd 3: kliknięcie przycisku "Wstaw bibliografię"
        doc.active_slide_index = 3
        win.bibliography_button.invoke()
        win.update_idletasks()

        bib_entries = doc.get_slide_bibliography_entries(3)
        assert [e.key for e in bib_entries] == [
            "KOWAL2023",
            "NOWAK2024",
            "DOVE2013",
            "ZIEL2022",
        ]
        assert [e.prefix for e in bib_entries] == ["[1] ", "[2] ", "[3] ", "[4] "]
        assert doc.read_slides()[2].has_bibliography is True

        # Zmiana liczby kolumn bibliografii na 2 (przebudowuje prezentację z kopią zapasową)
        win.bib_columns_var.set("2")
        win.bib_columns_combo.event_generate("<<ComboboxSelected>>")
        win.update_idletasks()
        assert win.controller.settings.bib_columns == 2
        cols = doc.get_slide_bibliography_columns(3)
        assert len(cols) == 2
        assert [len(col) for col in cols] == [2, 2]

        # Zmiana czcionki, rozmiaru i trybu dopasowania bibliografii
        assert win.set_bib_font_name("Arial") is True
        assert doc.get_slide_bibliography_settings(3).bib_font_name == "Arial"

        assert win.set_bib_font_size(14) is True
        assert doc.get_slide_bibliography_settings(3).bib_font_size_pt == 14

        win.bib_fit_var.set(BIB_FIT_LABELS["fixed_font"])
        win.bib_fit_combo.event_generate("<<ComboboxSelected>>")
        win.update_idletasks()
        assert win.controller.settings.bib_fit_mode == "fixed_font"
        assert doc.get_slide_bibliography_settings(3).bib_fit_mode == "fixed_font"

        assert win.set_bib_fit_mode("split_slides") is True
        assert doc.get_slide_bibliography_settings(3).bib_fit_mode == "split_slides"
    finally:
        win.destroy()


def test_duplicate_collection_names_are_disambiguated_by_key_h1(
    tk_root: tk.Tk,
) -> None:
    zot = FakeZoteroSource(
        collections=[
            ZoteroCollection(key="COL_1", name="Robocze"),
            ZoteroCollection(key="COL_2", name="Robocze"),
            ZoteroCollection(key="COL_ALL", name=ALL_COLLECTIONS_LABEL),
        ],
        collection_items={
            "COL_1": {"KOWAL2023"},
            "COL_2": {"NOWAK2024"},
            "COL_ALL": {"DOVE2013"},
        },
    )
    win, _zot, _doc = _make_window(tk_root=tk_root, zotero=zot)
    try:
        values = list(win.collection_combo.cget("values"))
        assert values == [
            ALL_COLLECTIONS_LABEL,
            "Robocze (COL_1)",
            "Robocze (COL_2)",
            f"{ALL_COLLECTIONS_LABEL} (COL_ALL)",
        ]

        # Wybór pierwszej z kolekcji o nazwie "Robocze"
        win.collection_var.set("Robocze (COL_1)")
        win.collection_combo.event_generate("<<ComboboxSelected>>")
        win.update_idletasks()
        assert win.get_selected_collection_key() == "COL_1"
        assert [i.key for i in win.get_result_items()] == ["KOWAL2023"]

        # Wybór drugiej z kolekcji o nazwie "Robocze"
        win.collection_var.set("Robocze (COL_2)")
        win.collection_combo.event_generate("<<ComboboxSelected>>")
        win.update_idletasks()
        assert win.get_selected_collection_key() == "COL_2"
        assert [i.key for i in win.get_result_items()] == ["NOWAK2024"]

        # Wybór kolekcji o nazwie kolidującej z ALL_COLLECTIONS_LABEL
        win.set_collection("COL_ALL")
        assert win.get_selected_collection_key() == "COL_ALL"
        assert win.collection_var.get() == f"{ALL_COLLECTIONS_LABEL} (COL_ALL)"
        assert [i.key for i in win.get_result_items()] == ["DOVE2013"]
    finally:
        win.destroy()


def test_collections_refresh_automatically_when_zotero_comes_online_m1(
    tk_root: tk.Tk,
) -> None:
    zot = FakeZoteroSource(offline=True)
    win, zot, _doc = _make_window(tk_root=tk_root, zotero=zot)
    try:
        assert list(win.collection_combo.cget("values")) == [ALL_COLLECTIONS_LABEL]

        # Włączenie Zotero w trakcie działania okienka
        zot.offline = False
        win.perform_search("")
        values = list(win.collection_combo.cget("values"))
        assert "Sztuczna inteligencja" in values
        assert "Sztuczna inteligencja / Uczenie maszynowe" in values
        assert "Etyka i prawo" in values

        win.set_collection("Etyka i prawo")
        assert win.get_selected_collection_key() == "COL_ETH"
        assert [i.key for i in win.get_result_items()] == ["KOWAL2023", "DOVE2013"]
    finally:
        win.destroy()


def test_insert_bibliography_errors_empty_offline_and_com_exception_m2(
    tk_root: tk.Tk,
) -> None:
    zot = FakeZoteroSource()
    doc = FakeSlideDoc()
    win, zot, doc = _make_window(tk_root=tk_root, zotero=zot, doc=doc)
    try:
        # 1. Prezentacja bez żadnych cytowań -> False + czytelny komunikat
        assert win.insert_bibliography() is False
        assert "Brak wykorzystanych źródeł" in win.get_status_message()

        # Dodajemy cytowanie, po czym wyłączamy Zotero
        win.select_result_keys(["KOWAL2023"])
        assert win.cite_selected() is True

        zot.offline = True
        assert win.insert_bibliography() is False
        assert "Zotero nie odpowiada" in win.get_status_message()
    finally:
        win.destroy()

    # 3. Wyjątek COM z warstwy slide_doc podczas insert_bibliography
    class FailingBibDoc(FakeSlideDoc):
        def set_slide_bibliography(self, slide_index, entries, settings) -> int:  # type: ignore[no-untyped-def]
            raise RuntimeError("COM failure in set_slide_bibliography")

    fail_doc = FailingBibDoc()
    win_fail, _zot, fail_doc = _make_window(tk_root=tk_root, doc=fail_doc)
    try:
        win_fail.select_result_keys(["KOWAL2023"])
        assert win_fail.cite_selected() is True
        assert win_fail.insert_bibliography() is False
        assert "PowerPoint zgłosił błąd" in win_fail.get_status_message()
    finally:
        win_fail.destroy()


def test_bibliography_settings_rollback_on_unsaved_d18_and_sync_on_switch_d19_m3(
    tk_root: tk.Tk,
) -> None:
    from azp.contracts import MSG_SAVE_BEFORE_REBUILD

    # 1. D18: Niezapisana prezentacja wycofuje zmiany 4 ustawień bibliografii
    unsaved_doc = FakeSlideDoc(presentation_path=None)
    win_unsaved, _zot, _doc = _make_window(tk_root=tk_root, doc=unsaved_doc)
    try:
        assert win_unsaved.set_bib_font_name("Arial") is False
        assert win_unsaved.get_status_message() == MSG_SAVE_BEFORE_REBUILD
        assert win_unsaved.controller.settings.bib_font_name == "Calibri"
        assert win_unsaved.bib_font_name_var.get() == "Calibri"

        assert win_unsaved.set_bib_font_size(18) is False
        assert win_unsaved.controller.settings.bib_font_size_pt == 12
        assert win_unsaved.bib_font_size_var.get() == "12"

        assert win_unsaved.set_bib_fit_mode("split_slides") is False
        assert win_unsaved.controller.settings.bib_fit_mode == "shrink_to_fit"
        assert win_unsaved.bib_fit_var.get() == BIB_FIT_LABELS["shrink_to_fit"]

        assert win_unsaved.set_bib_columns(3) is False
        assert win_unsaved.controller.settings.bib_columns == 1
        assert win_unsaved.bib_columns_var.get() == "1"
    finally:
        win_unsaved.destroy()

    # 2. D19: Przełączenie aktywnej prezentacji synchronizuje ustawienia bibliografii
    doc_a = FakeSlideDoc(presentation_path=Path("C:/fake/A.pptx"))
    doc_a.save_settings(
        Settings(
            bib_font_name="Calibri",
            bib_font_size_pt=12,
            bib_fit_mode="shrink_to_fit",
            bib_columns=1,
        )
    )
    doc_b = FakeSlideDoc(presentation_path=Path("C:/fake/B.pptx"))
    doc_b.save_settings(
        Settings(
            bib_font_name="Arial",
            bib_font_size_pt=16,
            bib_fit_mode="split_slides",
            bib_columns=3,
        )
    )
    multi = _MultiPresentationDoc({"A": doc_a, "B": doc_b}, active="A")
    win_multi = AzpWindow(
        zotero=FakeZoteroSource(),
        slide_doc=multi,  # type: ignore[arg-type]
        master=tk_root,
        poll_interval_ms=None,
    )
    try:
        assert win_multi.bib_font_name_var.get() == "Calibri"
        assert win_multi.bib_columns_var.get() == "1"

        multi.activate("B")
        win_multi.check_active_presentation()
        assert win_multi.bib_font_name_var.get() == "Arial"
        assert win_multi.bib_font_size_var.get() == "16"
        assert win_multi.bib_fit_var.get() == BIB_FIT_LABELS["split_slides"]
        assert win_multi.bib_columns_var.get() == "3"
    finally:
        win_multi.destroy()


# ---------------------------------------------------------------------------
# D3: Zwijana sekcja ustawień (D29) i wstawianie bibliografii z zaznaczonych pozycji / kolekcji (D28)
# ---------------------------------------------------------------------------


def test_collapsible_settings_panel_default_toggle_and_config_persistence_d29(
    tk_root: tk.Tk, tmp_path: Path
) -> None:
    from azp.config import load_config
    from azp.gui import HIDE_SETTINGS_LABEL, SHOW_SETTINGS_LABEL

    cfg_path = tmp_path / "config.json"
    doc = FakeSlideDoc()
    win = AzpWindow(
        zotero=FakeZoteroSource(),
        slide_doc=doc,
        master=tk_root,
        poll_interval_ms=None,
        config_path=cfg_path,
    )
    try:
        # Domyślnie sekcja ustawień jest zwinięta (ukryta)
        assert win.is_settings_expanded() is False
        assert win.controller.settings.settings_expanded is False
        assert win.toggle_settings_button.cget("text") == SHOW_SETTINGS_LABEL
        assert win._settings_frame.winfo_manager() == ""

        # Kliknięcie przycisku rozwija sekcję ustawień i zapisuje w config.json bez przebudowy prezentacji
        win.toggle_settings_button.invoke()
        win.update_idletasks()
        assert win.is_settings_expanded() is True
        assert win.controller.settings.settings_expanded is True
        assert win.toggle_settings_button.cget("text") == HIDE_SETTINGS_LABEL
        assert win._settings_frame.winfo_manager() == "grid"
        assert "Rozwinięto ustawienia" in win.get_status_message()
        assert doc.backups_created == []
        assert load_config(cfg_path).settings_expanded is True

        # Ponowne kliknięcie zwija sekcję ustawień
        win.toggle_settings_button.invoke()
        win.update_idletasks()
        assert win.is_settings_expanded() is False
        assert win.controller.settings.settings_expanded is False
        assert win.toggle_settings_button.cget("text") == SHOW_SETTINGS_LABEL
        assert win._settings_frame.winfo_manager() == ""
        assert "Ukryto ustawienia" in win.get_status_message()
        assert doc.backups_created == []
        assert load_config(cfg_path).settings_expanded is False

        # Programowe przełączenie set_settings_expanded(True)
        win.set_settings_expanded(True)
        assert win.is_settings_expanded() is True
        assert load_config(cfg_path).settings_expanded is True
    finally:
        win.destroy()

    # Nowe okienko odtwarza stan rozwinięcia z config.json, nawet gdy prezentacja ma zapisane settings_expanded=False
    doc_with_tags = FakeSlideDoc()
    doc_with_tags.save_settings(Settings(settings_expanded=False))
    win2 = AzpWindow(
        zotero=FakeZoteroSource(),
        slide_doc=doc_with_tags,
        master=tk_root,
        poll_interval_ms=None,
        config_path=cfg_path,
    )
    try:
        assert win2.is_settings_expanded() is True
        assert win2._settings_frame.winfo_manager() == "grid"
    finally:
        win2.destroy()


def test_insert_bibliography_from_selected_results_or_collection_on_clean_presentation_d28(
    tk_root: tk.Tk,
) -> None:
    # 1. Czysta prezentacja bez cytowań + zaznaczone pozycje na liście wyników -> wstawia bibliografię
    doc1 = FakeSlideDoc(slide_count=2, active_slide_index=2, has_cursor=False)
    win1, _zot1, doc1 = _make_window(tk_root=tk_root, doc=doc1)
    try:
        win1.select_result_keys(["KOWAL2023", "DOVE2013"])
        win1.bibliography_button.invoke()
        win1.update_idletasks()

        assert "Wstawiono bibliografię (2 źródeł)" in win1.get_status_message()
        entries = doc1.get_slide_bibliography_entries(2)
        assert [e.key for e in entries] == ["KOWAL2023", "DOVE2013"]
        assert [e.prefix for e in entries] == ["[1] ", "[2] "]

        # Zmiana liczby kolumn i stylu na APA zachowuje tę bibliografię
        assert win1.set_bib_columns(2) is True
        assert len(doc1.get_slide_bibliography_columns(2)) == 2
        assert win1.set_style(STYLE_APA) is True
        assert [e.key for e in doc1.get_slide_bibliography_entries(2)] == [
            "DOVE2013",
            "KOWAL2023",
        ]
    finally:
        win1.destroy()

    # 2. Czysta prezentacja bez cytowań i bez zaznaczenia, ale z wybraną kolekcją Zotero -> wstawia pozycje z kolekcji
    doc2 = FakeSlideDoc(slide_count=1, active_slide_index=1, has_cursor=False)
    win2, _zot2, doc2 = _make_window(tk_root=tk_root, doc=doc2)
    try:
        win2.set_collection("COL_ETH")
        assert win2.get_selected_keys() == ()
        assert win2.insert_bibliography() is True
        assert [e.key for e in doc2.get_slide_bibliography_entries(1)] == [
            "KOWAL2023",
            "DOVE2013",
        ]
    finally:
        win2.destroy()
