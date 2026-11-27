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
    ZoteroItem,
)
from azp.gui import (
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

        # Wybór pozycji, ale brak kursora w polu tekstowym (D10)
        win.select_result_keys(["KOWAL2023"])
        assert win.cite_selected() is False
        assert "kursor" in win.get_status_message().casefold()
        assert doc.get_slide_body_text(1) == before_body
        assert doc.read_slides()[0].citations == ()
        assert doc.read_slides()[0].footer_entries == ()
        assert doc.load_settings() is None
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
