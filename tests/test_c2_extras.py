"""Testy etapu C2: konfiguracja `config.json`, skróty klawiszowe (D15) i skrypt startowy (D16)."""

from __future__ import annotations

import json
import tkinter as tk
from pathlib import Path

import pytest

from azp.config import DEFAULT_CONFIG_PATH, load_config, save_config
from azp.contracts import STYLE_APA, STYLE_IEEE_V4, Settings
from azp.gui import AzpController, AzpWindow
from azp.main import build_window
from tests.fakes.fake_ppt import FakeSlideDoc
from tests.fakes.fake_zotero import FakeZoteroSource

REPO_ROOT = Path(__file__).resolve().parent.parent


# ---------------------------------------------------------------------------
# 1. load_config / save_config (`azp/config.py`)
# ---------------------------------------------------------------------------


def test_default_config_path_points_to_repo_root_and_is_gitignored() -> None:
    assert DEFAULT_CONFIG_PATH == REPO_ROOT / "config.json"
    gitignore_lines = (REPO_ROOT / ".gitignore").read_text(encoding="utf-8").splitlines()
    assert "config.json" in [line.strip() for line in gitignore_lines]


def test_save_and_load_config_roundtrip_all_fields(tmp_path: Path) -> None:
    cfg_path = tmp_path / "nested" / "config.json"
    custom = Settings(
        style_id=STYLE_APA,
        numbering_mode="per_slide",
        font_name="Arial",
        font_size_pt=14,
        footer_position="replace_footer",
        footer_header=False,
        always_on_top=False,
    )

    returned = save_config(custom, cfg_path)
    assert returned == cfg_path
    assert cfg_path.exists()

    raw = json.loads(cfg_path.read_text(encoding="utf-8"))
    assert raw == custom.to_dict()
    assert load_config(cfg_path) == custom


@pytest.mark.parametrize(
    "content",
    [
        None,  # brak pliku
        "",  # pusty plik
        "   \n\t ",  # same białe znaki
        "{niepoprawny json",  # uszkodzony JSON
        "[1, 2, 3]",  # poprawny JSON, ale nie obiekt
        json.dumps({"font_size_pt": 999}),  # rozmiar spoza zakresu 6..72
        json.dumps({"font_size_pt": "nie-liczba"}),  # zły typ
        json.dumps({"numbering_mode": "invalid_mode"}),  # nieznany tryb
        json.dumps({"footer_position": "top_center"}),  # nieznane położenie
        json.dumps({"font_name": "   "}),  # pusta czcionka
        json.dumps({"style_id": ""}),  # pusty styl
        json.dumps({"font_name": None}),  # M1: None jako str
        json.dumps({"font_name": ["Arial"]}),  # M1: lista jako str
        json.dumps({"style_id": None}),  # M1: None jako style_id
        json.dumps({"footer_header": "false", "font_size_pt": 12}),  # M1: str jako bool
        json.dumps({"always_on_top": "false", "font_size_pt": 12}),  # M1: str jako bool
        json.dumps({"font_size_pt": True}),  # M1: bool jako int
        json.dumps({"font_size_pt": 12.7}),  # M1: float jako int
    ],
)
def test_load_config_falls_back_to_defaults_on_missing_or_invalid_data(
    tmp_path: Path, content: str | None
) -> None:
    cfg_path = tmp_path / "config.json"
    if content is not None:
        cfg_path.write_text(content, encoding="utf-8")

    assert load_config(cfg_path) == Settings()


# ---------------------------------------------------------------------------
# 2. Integracja `config.json` z `AzpController`, `AzpWindow` i `build_window`
# ---------------------------------------------------------------------------


def test_controller_loads_and_saves_config_and_preserves_always_on_top(
    tmp_path: Path,
) -> None:
    cfg_path = tmp_path / "config.json"
    saved_cfg = Settings(
        style_id=STYLE_APA,
        numbering_mode="per_slide",
        font_name="Arial",
        font_size_pt=12,
        footer_position="bottom_left",
        footer_header=False,
        always_on_top=False,
    )
    save_config(saved_cfg, cfg_path)

    # 2a. Prezentacja bez własnych tagów -> wczytuje pełny config.json
    doc_empty = FakeSlideDoc()
    ctrl_empty = AzpController(
        zotero=FakeZoteroSource(),
        slide_doc=doc_empty,
        config_path=cfg_path,
    )
    assert ctrl_empty.settings == saved_cfg

    # 2b. Prezentacja z własnymi tagami -> zachowuje ustawienia prezentacji,
    #     ale `always_on_top` bierze z `config.json`
    doc_with_tags = FakeSlideDoc()
    doc_with_tags.save_settings(
        Settings(style_id=STYLE_IEEE_V4, font_name="Calibri", font_size_pt=11, always_on_top=True)
    )
    ctrl_tags = AzpController(
        zotero=FakeZoteroSource(),
        slide_doc=doc_with_tags,
        config_path=cfg_path,
    )
    assert ctrl_tags.settings.style_id == STYLE_IEEE_V4
    assert ctrl_tags.settings.font_size_pt == 11
    assert ctrl_tags.settings.always_on_top is False

    # 2c. Brak otwartej prezentacji: zmiana ustawienia zapisuje je jako domyślne do config.json
    doc_none = FakeSlideDoc(slide_count=0, active_slide_index=None)
    ctrl_none = AzpController(
        zotero=FakeZoteroSource(),
        slide_doc=doc_none,
        config_path=cfg_path,
    )
    assert ctrl_none.update_settings(rebuild=True, font_size_pt=16) is False
    assert load_config(cfg_path).font_size_pt == 16


def test_window_and_build_window_persist_config_and_ignore_os_error(
    tk_root: tk.Tk, tmp_path: Path
) -> None:
    cfg_path = tmp_path / "config.json"
    save_config(Settings(font_name="Arial", font_size_pt=12, always_on_top=False), cfg_path)

    doc = FakeSlideDoc()
    win = build_window(
        zotero=FakeZoteroSource(),
        slide_doc=doc,
        master=tk_root,
        poll_interval_ms=None,
        config_path=cfg_path,
    )
    try:
        assert win.controller.settings.font_name == "Arial"
        assert win.controller.settings.font_size_pt == 12
        assert win.controller.settings.always_on_top is False
        assert bool(win.root.wm_attributes("-topmost")) is False

        # Zmiana opcji formatowania (rebuild=True) zapisuje config.json
        assert win.set_style(STYLE_APA) is True
        reloaded = load_config(cfg_path)
        assert reloaded.style_id == STYLE_APA
        assert reloaded.font_name == "Arial"

        # Przełączenie `always_on_top` (rebuild=False) zapisuje config.json
        win.set_always_on_top(True)
        assert load_config(cfg_path).always_on_top is True
    finally:
        win.destroy()

    # Błąd zapisu (np. ścieżka jest katalogiem -> OSError) nie przerywa pracy okienka
    dir_as_cfg = tmp_path / "config_dir"
    dir_as_cfg.mkdir()
    win_err = AzpWindow(
        zotero=FakeZoteroSource(),
        slide_doc=FakeSlideDoc(),
        master=tk_root,
        poll_interval_ms=None,
        config_path=dir_as_cfg,
    )
    try:
        win_err.set_always_on_top(False)
        assert win_err.controller.settings.always_on_top is False
    finally:
        win_err.destroy()


def test_failed_rebuild_does_not_overwrite_config_json(
    tk_root: tk.Tk, tmp_path: Path
) -> None:
    cfg_path = tmp_path / "config.json"
    initial = Settings(style_id=STYLE_IEEE_V4, font_name="Calibri", font_size_pt=10)
    save_config(initial, cfg_path)

    # 1. D18: niezapisana prezentacja (presentation_path=None)
    unsaved_doc = FakeSlideDoc(presentation_path=None)
    win_unsaved = AzpWindow(
        zotero=FakeZoteroSource(),
        slide_doc=unsaved_doc,
        master=tk_root,
        poll_interval_ms=None,
        config_path=cfg_path,
    )
    try:
        assert win_unsaved.set_style(STYLE_APA) is False
        assert load_config(cfg_path) == initial
    finally:
        win_unsaved.destroy()

    # 2. ZoteroOfflineError podczas przebudowy
    zot = FakeZoteroSource()
    saved_doc = FakeSlideDoc()
    win_offline = AzpWindow(
        zotero=zot,
        slide_doc=saved_doc,
        master=tk_root,
        poll_interval_ms=None,
        config_path=cfg_path,
    )
    try:
        win_offline.select_result_keys(["KOWAL2023"])
        assert win_offline.cite_selected() is True

        zot.offline = True
        assert win_offline.set_style(STYLE_APA) is False
        assert load_config(cfg_path) == initial
    finally:
        win_offline.destroy()


# ---------------------------------------------------------------------------
# 3. Skróty klawiszowe wewnątrz okienka (D15)
# ---------------------------------------------------------------------------


def test_keyboard_shortcuts_d15_in_window(tk_root: tk.Tk) -> None:
    doc = FakeSlideDoc()
    win = AzpWindow(
        zotero=FakeZoteroSource(),
        slide_doc=doc,
        master=tk_root,
        poll_interval_ms=None,
    )
    try:
        win.root.update()
        win.root.focus_force()
        win.root.update()

        # Enter na liście wyników -> Cytuj
        win.select_result_keys(["KOWAL2023"])
        win.results_listbox.focus_set()
        win.root.update()
        win.results_listbox.event_generate("<Return>")
        win.root.update()
        assert len(doc.read_slides()[0].citations) == 1
        assert doc.read_slides()[0].citations[0].keys == ("KOWAL2023",)

        # Ctrl+Enter na liście wyników -> Cytuj dokładnie raz ("break" blokuje podwójne wywołanie)
        win.select_result_keys(["NOWAK2024"])
        win.results_listbox.focus_set()
        win.root.update()
        win.results_listbox.event_generate("<Control-Return>")
        win.root.update()
        assert len(doc.read_slides()[0].citations) == 2
        assert doc.read_slides()[0].citations[1].keys == ("NOWAK2024",)

        # Ctrl+Enter z innego widgetu w oknie (np. z pola wyszukiwania) -> Cytuj
        win.select_result_keys(["ZIEL2022"])
        win.search_entry.focus_set()
        win.root.update()
        win.search_entry.event_generate("<Control-Return>")
        win.root.update()
        assert len(doc.read_slides()[0].citations) == 3
        assert doc.read_slides()[0].citations[2].keys == ("ZIEL2022",)

        # F5 i Ctrl+R -> Odśwież / Przenumeruj
        assert doc.backups_created == []
        win.root.event_generate("<F5>")
        win.root.update()
        assert len(doc.backups_created) == 1

        win.root.event_generate("<Control-r>")
        win.root.update()
        assert len(doc.backups_created) == 2

        win.root.event_generate("<Control-R>")
        win.root.update()
        assert len(doc.backups_created) == 3

        # Ctrl+F / focus_search_entry -> fokus na search_entry i zaznaczenie całego tekstu
        win.set_search_query("Kowalski")
        win.results_listbox.focus_set()
        win.root.update()
        win.root.event_generate("<Control-f>")
        win.root.update()
        assert win.root.focus_get() == win.search_entry
        assert win.search_entry.selection_present()
        assert win.search_entry.selection_get() == "Kowalski"

        win.results_listbox.focus_set()
        win.search_entry.selection_clear()
        win.root.update()
        win.focus_search_entry()
        win.root.update()
        assert win.root.focus_get() == win.search_entry
        assert win.search_entry.selection_present()
        assert win.search_entry.selection_get() == "Kowalski"
    finally:
        win.destroy()


# ---------------------------------------------------------------------------
# 4. Skrypty startowe bez czarnej konsoli (D16)
# ---------------------------------------------------------------------------


def test_launcher_scripts_exist_and_are_valid_d16() -> None:
    pyw_path = REPO_ROOT / "start_azp.pyw"
    bat_path = REPO_ROOT / "start_azp.bat"

    assert pyw_path.is_file()
    assert bat_path.is_file()

    pyw_source = pyw_path.read_text(encoding="utf-8")
    compile(pyw_source, str(pyw_path), "exec")
    assert "sys.path" in pyw_source
    assert "from azp.main import main" in pyw_source
    assert "main()" in pyw_source

    bat_source = bat_path.read_text(encoding="utf-8")
    assert 'cd /d "%~dp0"' in bat_source
    assert 'start "" pythonw "%~dp0start_azp.pyw"' in bat_source
    assert "python -m azp" in bat_source



def test_h1_test_main_does_not_depend_on_or_touch_repo_config_json(
    tk_root: tk.Tk, monkeypatch: pytest.MonkeyPatch
) -> None:
    """H1 (C2): testy w test_main.py muszą przekazywać config_path=None, nie używać config.json z dysku."""
    import tests.test_main as tm

    config_io_calls: list[Path | None] = []
    monkeypatch.setattr(
        "azp.gui.app.load_config",
        lambda path=None: (config_io_calls.append(path), Settings())[1],
    )
    monkeypatch.setattr(
        "azp.gui.app.save_config",
        lambda settings, path=None: (config_io_calls.append(path), Path("dummy"))[1],
    )

    tm.test_build_window_wires_engine_and_returns_focus_after_cite_d20(tk_root)
    tm.test_build_window_turns_unexpected_powerpoint_errors_into_status(tk_root)
    assert config_io_calls == []
