"""Testy etapu C2: konfiguracja `config.json`, skróty klawiszowe (D15) i skrypt startowy (D16)."""

from __future__ import annotations

import io
import json
import subprocess
import time
import tkinter as tk
import zipfile
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
        json.dumps({"dark_mode": "true", "font_size_pt": 12}),  # str jako bool dla dark_mode
        json.dumps({"settings_expanded": "false", "font_size_pt": 12}),  # str jako bool dla settings_expanded
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


def test_install_pp_addin_bat_and_cli_flags(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from azp.main import main

    addin_bat = REPO_ROOT / "install_pp_addin.bat"
    assert addin_bat.is_file()
    assert "--install-addin" in addin_bat.read_text(encoding="utf-8")

    installed: list[bool] = []
    uninstalled: list[bool] = []
    monkeypatch.setattr(
        "azp.main.install_powerpoint_addin",
        lambda: (installed.append(True), tmp_path / "AsystentZotero.ppam")[1],
    )
    monkeypatch.setattr(
        "azp.main.uninstall_powerpoint_addin",
        lambda: uninstalled.append(True),
    )

    assert main(["--install-addin"]) == 0
    assert installed == [True]
    assert main(["--uninstall-addin"]) == 0
    assert uninstalled == [True]


# ---------------------------------------------------------------------------
# 5. Instalator i aktualizator (`azp/updater.py`, `--update`, .bat, README.md)
# ---------------------------------------------------------------------------


def _build_zip_bytes(entries: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in entries.items():
            zf.writestr(name, data)
    return buf.getvalue()


def test_extract_github_zip_bytes_extracts_preserves_config_skips_git_and_blocks_zip_slip(
    tmp_path: Path,
) -> None:
    from azp.updater import extract_github_zip_bytes

    existing_cfg = tmp_path / "config.json"
    existing_cfg.write_text('{"font_size_pt": 18}\n', encoding="utf-8")

    zip_bytes = _build_zip_bytes(
        {
            "zotero-ppt-assistant-main/": b"",
            "zotero-ppt-assistant-main/azp/main.py": b"# updated main\n",
            "zotero-ppt-assistant-main/README.md": b"# AZP\n",
            "zotero-ppt-assistant-main/config.json": b'{"font_size_pt": 10}\n',
            "zotero-ppt-assistant-main/.git/HEAD": b"ref: refs/heads/main\n",
        }
    )

    count = extract_github_zip_bytes(zip_bytes, tmp_path)
    assert count == 2
    assert (tmp_path / "azp" / "main.py").read_text(encoding="utf-8") == "# updated main\n"
    assert (tmp_path / "README.md").read_text(encoding="utf-8") == "# AZP\n"
    assert existing_cfg.read_text(encoding="utf-8") == '{"font_size_pt": 18}\n'
    assert not (tmp_path / ".git").exists()

    for bad_entry in (
        "../evil.txt",
        "zotero-ppt-assistant-main/../../evil.txt",
        "zotero-ppt-assistant-main/C:evil.txt",
        "C:evil.txt/foo.txt",
        "zotero-ppt-assistant-main/sub/..",
        "zotero-ppt-assistant-main/a/../b.py",
        "zotero-ppt-assistant-main/plik.txt:stream",
    ):
        bad_zip = _build_zip_bytes({bad_entry: b"boom"})
        with pytest.raises(ValueError):
            extract_github_zip_bytes(bad_zip, tmp_path)

    # Dwufazowa atomowość: błąd walidacji kolejnego wpisu nie zapisuje wcześniejszego pliku
    atomic_target = tmp_path / "atomic_dir"
    atomic_target.mkdir()
    partial_zip = _build_zip_bytes(
        {
            "zotero-ppt-assistant-main/should_not_exist.txt": b"hello",
            "zotero-ppt-assistant-main/a/../b.py": b"bad",
        }
    )
    with pytest.raises(ValueError):
        extract_github_zip_bytes(partial_zip, atomic_target)
    assert not (atomic_target / "should_not_exist.txt").exists()

    # Puste archiwum ZIP -> ValueError
    with pytest.raises(ValueError):
        extract_github_zip_bytes(_build_zip_bytes({}), tmp_path)


def test_update_from_github_zip_downloads_and_extracts(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from azp.updater import auto_update_if_needed, update_from_github_zip

    sha = "d" * 40
    payload = _build_zip_bytes(
        {
            "zotero-ppt-assistant-main/azp/new_mod.py": b"VALUE = 1\n",
        }
    )

    class _FakeResponse:
        def __init__(self, data: bytes) -> None:
            self._data = data

        def read(self) -> bytes:
            return self._data

        def __enter__(self) -> _FakeResponse:
            return self

        def __exit__(self, *_args: object) -> None:
            return None

    seen_urls: list[tuple[str, float]] = []

    def fake_urlopen(url: str, timeout: float = 30.0) -> _FakeResponse:
        seen_urls.append((url, timeout))
        return _FakeResponse(payload)

    monkeypatch.setattr("urllib.request.urlopen", fake_urlopen)
    monkeypatch.setattr("azp.updater.fetch_remote_main_sha", lambda **_kw: sha)
    ok, msg = update_from_github_zip(tmp_path, zip_url="https://example.invalid/archive.zip", timeout=12.0)
    assert ok is True
    assert "1" in msg
    assert seen_urls == [("https://example.invalid/archive.zip", 12.0)]
    assert (tmp_path / "azp" / "new_mod.py").read_text(encoding="utf-8") == "VALUE = 1\n"
    assert (tmp_path / ".azp_version").read_text(encoding="utf-8").strip() == sha

    # Kolejne auto_update_if_needed widzi zgodny SHA w .azp_version i nie pobiera ZIP-a
    monkeypatch.setattr(
        "urllib.request.urlopen",
        lambda *_args, **_kw: (_ for _ in ()).throw(AssertionError("ZIP nie powinien być pobierany")),
    )
    changed, auto_msg = auto_update_if_needed(project_root=tmp_path)
    assert changed is False
    assert "aktualna" in auto_msg.lower()


def test_update_from_git_handles_missing_git_dirty_tree_and_clean_pull(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from azp.updater import update_from_git

    # Brak .git -> (False, ...)
    ok_no_git, msg_no_git = update_from_git(tmp_path)
    assert ok_no_git is False
    assert "git" in msg_no_git.lower()

    (tmp_path / ".git").mkdir()
    monkeypatch.setattr("azp.updater.shutil.which", lambda _name: "git")

    # Lokalne modyfikacje -> pominięcie git pull, zwrot (True, ...)
    calls_dirty: list[list[str]] = []

    def fake_run_dirty(cmd: list[str], **_kwargs: object) -> subprocess.CompletedProcess[str]:
        calls_dirty.append(cmd)
        return subprocess.CompletedProcess(cmd, 0, stdout=" M azp/main.py\n", stderr="")

    monkeypatch.setattr("azp.updater.subprocess.run", fake_run_dirty)
    ok_dirty, msg_dirty = update_from_git(tmp_path)
    assert ok_dirty is True
    assert "lokalne zmiany" in msg_dirty.lower()
    assert calls_dirty == [["git", "status", "--porcelain", "--untracked-files=no"]]

    # Czyste repozytorium -> wywołanie git pull --ff-only
    calls_clean: list[list[str]] = []

    def fake_run_clean(cmd: list[str], **_kwargs: object) -> subprocess.CompletedProcess[str]:
        calls_clean.append(cmd)
        return subprocess.CompletedProcess(cmd, 0, stdout="", stderr="")

    monkeypatch.setattr("azp.updater.subprocess.run", fake_run_clean)
    ok_clean, msg_clean = update_from_git(tmp_path)
    assert ok_clean is True
    assert "git pull" in msg_clean.lower()
    assert calls_clean == [
        ["git", "status", "--porcelain", "--untracked-files=no"],
        ["git", "pull", "--ff-only"],
    ]


def test_update_project_code_falls_back_to_github_zip_when_no_git(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from azp.updater import update_project_code

    captured: list[tuple[Path, str]] = []

    def fake_zip_update(
        project_root: Path, zip_url: str = "", timeout: float = 30.0
    ) -> tuple[bool, str]:
        captured.append((project_root, zip_url))
        return True, "Pobrano i zaktualizowano 3 plików z GitHuba."

    monkeypatch.setattr("azp.updater.update_from_github_zip", fake_zip_update)
    ok, msg = update_project_code(project_root=tmp_path, zip_url="https://example.invalid/main.zip")
    assert ok is True
    assert "3" in msg
    assert captured == [(tmp_path, "https://example.invalid/main.zip")]


def test_update_project_code_does_not_fallback_to_zip_when_git_pull_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from azp.updater import update_project_code

    (tmp_path / ".git").mkdir()
    monkeypatch.setattr("azp.updater.shutil.which", lambda _name: "git")

    def fake_run(cmd: list[str], **_kwargs: object) -> subprocess.CompletedProcess[str]:
        if cmd[:2] == ["git", "status"]:
            return subprocess.CompletedProcess(cmd, 0, stdout="", stderr="")
        return subprocess.CompletedProcess(cmd, 1, stdout="", stderr="fatal: cannot lock ref")

    monkeypatch.setattr("azp.updater.subprocess.run", fake_run)

    zip_called: list[bool] = []
    monkeypatch.setattr(
        "azp.updater.update_from_github_zip",
        lambda *_args, **_kwargs: (zip_called.append(True), (True, "zip"))[1],
    )

    ok, msg = update_project_code(project_root=tmp_path)
    assert ok is False
    assert msg.startswith("git pull nie powiódł się:")
    assert "fatal: cannot lock ref" in msg
    assert zip_called == []


def test_main_update_flag_updates_code_and_installs_addin(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    from azp.main import main

    calls: list[str] = []
    monkeypatch.setattr(
        "azp.main.install_powerpoint_addin",
        lambda: (calls.append("addin"), tmp_path / "AsystentZotero.ppam")[1],
    )

    monkeypatch.setattr(
        "azp.main.update_project_code",
        lambda: (calls.append("update"), (True, "Zaktualizowano kod przez git pull."))[1],
    )
    assert main(["--update"]) == 0
    assert calls == ["update", "addin"]
    out = capsys.readouterr().out
    assert "Zaktualizowano kod przez git pull." in out
    assert "AsystentZotero.ppam" in out

    calls.clear()
    monkeypatch.setattr(
        "azp.main.update_project_code",
        lambda: (calls.append("update"), (False, "Błąd sieci"))[1],
    )
    assert main(["--update"]) == 1
    assert calls == ["update", "addin"]


def test_main_addin_and_update_handle_permission_error_gracefully(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    from azp.main import main

    def raise_locked() -> Path:
        raise PermissionError("Plik zablokowany")

    monkeypatch.setattr("azp.main.install_powerpoint_addin", raise_locked)
    monkeypatch.setattr(
        "azp.main.update_project_code",
        lambda: (True, "Zaktualizowano kod przez git pull."),
    )

    assert main(["--install-addin"]) == 1
    out_install = capsys.readouterr().out
    assert "Plik zablokowany" in out_install
    assert "zamknij PowerPointa" in out_install

    assert main(["--update"]) == 1
    out_update = capsys.readouterr().out
    assert "Zaktualizowano kod przez git pull." in out_update
    assert "Plik zablokowany" in out_update
    assert "zamknij PowerPointa" in out_update


def test_install_or_update_bat_and_readme_exist_and_have_expected_content() -> None:
    bat_path = REPO_ROOT / "instaluj_lub_aktualizuj.bat"
    readme_path = REPO_ROOT / "README.md"

    assert bat_path.is_file()
    assert readme_path.is_file()

    bat_raw = bat_path.read_bytes()
    assert b"\r\n" in bat_raw
    assert b"\n" not in bat_raw.replace(b"\r\n", b"")

    bat_text = bat_raw.decode("utf-8")
    assert "@echo off" in bat_text
    assert "chcp 65001" in bat_text
    assert 'cd /d "%~dp0"' in bat_text
    assert "python --version" in bat_text
    assert "Add python.exe to PATH" in bat_text
    assert "python -m pip install -r requirements.txt" in bat_text
    assert "python -m azp --update" in bat_text
    assert "python -m azp --install-addin" in bat_text
    assert "Asystent Zotero" in bat_text

    readme_text = readme_path.read_text(encoding="utf-8")
    assert "instaluj_lub_aktualizuj.bat" in readme_text
    assert "automatycznie w tle" in readme_text
    assert "IEEE v4" in readme_text
    assert "APA 7" in readme_text
    assert "config.json" in readme_text
    assert "start_azp.bat" in readme_text
    assert "install_pp_addin.bat" in readme_text
    assert "--uninstall-addin" in readme_text

    gitignore_lines = [
        line.strip()
        for line in (REPO_ROOT / ".gitignore").read_text(encoding="utf-8").splitlines()
    ]
    assert ".azp_version" in gitignore_lines


def test_fetch_remote_main_sha_supports_raw_hex_json_and_errors(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from azp.updater import fetch_remote_main_sha

    sha = "a" * 40

    class _Resp:
        def __init__(self, body: bytes) -> None:
            self._body = body

        def read(self) -> bytes:
            return self._body

        def __enter__(self) -> _Resp:
            return self

        def __exit__(self, *_args: object) -> None:
            return None

    monkeypatch.setattr("urllib.request.urlopen", lambda *_args, **_kw: _Resp(sha.encode("utf-8")))
    assert fetch_remote_main_sha() == sha

    monkeypatch.setattr(
        "urllib.request.urlopen",
        lambda *_args, **_kw: _Resp(json.dumps({"sha": sha.upper()}).encode("utf-8")),
    )
    assert fetch_remote_main_sha() == sha

    def raise_err(*_args: object, **_kw: object) -> object:
        raise OSError("network down")

    monkeypatch.setattr("urllib.request.urlopen", raise_err)
    assert fetch_remote_main_sha() is None


def test_auto_update_if_needed_git_mode(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from azp.updater import auto_update_if_needed

    (tmp_path / ".git").mkdir()
    monkeypatch.setattr("azp.updater.shutil.which", lambda _name: "git")

    # 1. Lokalne zmiany -> (False, ...) bez pulla
    calls_dirty: list[list[str]] = []

    def run_dirty(cmd: list[str], **_kw: object) -> subprocess.CompletedProcess[str]:
        calls_dirty.append(cmd)
        return subprocess.CompletedProcess(cmd, 0, stdout=" M azp/main.py\n", stderr="")

    monkeypatch.setattr("azp.updater.subprocess.run", run_dirty)
    changed, _msg = auto_update_if_needed(project_root=tmp_path)
    assert changed is False
    assert calls_dirty == [["git", "status", "--porcelain", "--untracked-files=no"]]

    # 2. old_sha == new_sha -> (False, ...)
    head_sha = "1" * 40

    def run_up_to_date(cmd: list[str], **_kw: object) -> subprocess.CompletedProcess[str]:
        if cmd[:2] == ["git", "rev-parse"]:
            return subprocess.CompletedProcess(cmd, 0, stdout=head_sha + "\n", stderr="")
        return subprocess.CompletedProcess(cmd, 0, stdout="", stderr="")

    monkeypatch.setattr("azp.updater.subprocess.run", run_up_to_date)
    changed_same, msg_same = auto_update_if_needed(project_root=tmp_path)
    assert changed_same is False
    assert "aktualna" in msg_same.lower()

    # 3. old_sha != new_sha -> (True, ...)
    new_sha = "2" * 40
    rev_calls = 0

    def run_updated(cmd: list[str], **_kw: object) -> subprocess.CompletedProcess[str]:
        nonlocal rev_calls
        if cmd[:2] == ["git", "rev-parse"]:
            rev_calls += 1
            sha = head_sha if rev_calls == 1 else new_sha
            return subprocess.CompletedProcess(cmd, 0, stdout=sha + "\n", stderr="")
        return subprocess.CompletedProcess(cmd, 0, stdout="", stderr="")

    monkeypatch.setattr("azp.updater.subprocess.run", run_updated)
    changed_new, msg_new = auto_update_if_needed(project_root=tmp_path)
    assert changed_new is True
    assert "git pull" in msg_new.lower()


def test_auto_update_if_needed_zip_mode(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from azp.updater import PRESERVED_RELATIVE_PATHS, auto_update_if_needed

    assert ".azp_version" in PRESERVED_RELATIVE_PATHS
    sha_v1 = "a" * 40
    sha_v2 = "b" * 40

    # 1. .azp_version == remote_sha -> brak pobierania ZIP-a, (False, ...)
    (tmp_path / ".azp_version").write_text(sha_v1, encoding="utf-8")
    monkeypatch.setattr("azp.updater.fetch_remote_main_sha", lambda **_kw: sha_v1)

    def fail_if_called(*_args: object, **_kw: object) -> object:
        raise AssertionError("ZIP nie powinien być pobierany przy zgodnym .azp_version")

    monkeypatch.setattr("urllib.request.urlopen", fail_if_called)
    changed_cached, _ = auto_update_if_needed(project_root=tmp_path)
    assert changed_cached is False

    # 2. Nowy SHA i zmiana pliku na dysku -> aktualizuje plik, zapisuje .azp_version, (True, ...)
    class _ZipResp:
        def __init__(self, data: bytes) -> None:
            self._data = data

        def read(self) -> bytes:
            return self._data

        def __enter__(self) -> _ZipResp:
            return self

        def __exit__(self, *_args: object) -> None:
            return None

    target_file = tmp_path / "azp" / "mod.py"
    target_file.parent.mkdir(parents=True, exist_ok=True)
    target_file.write_text("VER = 1\n", encoding="utf-8")

    zip_v2 = _build_zip_bytes({"zotero-ppt-assistant-main/azp/mod.py": b"VER = 2\n"})
    monkeypatch.setattr("azp.updater.fetch_remote_main_sha", lambda **_kw: sha_v2)
    monkeypatch.setattr("urllib.request.urlopen", lambda *_args, **_kw: _ZipResp(zip_v2))

    changed_v2, msg_v2 = auto_update_if_needed(project_root=tmp_path)
    assert changed_v2 is True
    assert "1" in msg_v2
    assert target_file.read_text(encoding="utf-8") == "VER = 2\n"
    assert (tmp_path / ".azp_version").read_text(encoding="utf-8").strip() == sha_v2

    # 3. Nowy SHA, ale pliki identyczne z dyskiem -> zapisuje nowy .azp_version, (False, ...)
    sha_v3 = "c" * 40
    monkeypatch.setattr("azp.updater.fetch_remote_main_sha", lambda **_kw: sha_v3)
    changed_same_bytes, _ = auto_update_if_needed(project_root=tmp_path)
    assert changed_same_bytes is False
    assert (tmp_path / ".azp_version").read_text(encoding="utf-8").strip() == sha_v3


def test_background_auto_update_and_reload_window(
    tk_root: tk.Tk, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from azp.main import main, reload_window_after_update, start_background_auto_update

    # 1. update_fn zwraca (False, ...) -> brak przeładowania, okienko pozostaje otwarte
    win_no_change = AzpWindow(
        zotero=FakeZoteroSource(),
        slide_doc=FakeSlideDoc(),
        master=tk_root,
        poll_interval_ms=None,
    )
    try:
        reloaded: list[str] = []
        thread = start_background_auto_update(
            win_no_change,
            update_fn=lambda: (False, "Wersja jest aktualna."),
            addin_fn=lambda: reloaded.append("addin"),
            reload_fn=lambda _w, msg: (reloaded.append(msg), True)[1],
        )
        thread.join(timeout=2.0)
        time.sleep(0.03)
        win_no_change.root.update()
        assert reloaded == []
        assert win_no_change.root.winfo_exists()
    finally:
        win_no_change.destroy()

    # 2. update_fn zwraca (True, "Nowa wersja") -> odświeża dodatek, wywołuje popen z --no-auto-update i zamyka okienko
    win_update = AzpWindow(
        zotero=FakeZoteroSource(),
        slide_doc=FakeSlideDoc(),
        master=tk_root,
        poll_interval_ms=None,
    )
    events: list[object] = []

    def fake_popen(cmd: list[str], **_kw: object) -> object:
        events.append(("popen", cmd))
        return object()

    thread_up = start_background_auto_update(
        win_update,
        update_fn=lambda: (True, "Nowa wersja"),
        addin_fn=lambda: (events.append("addin"), tmp_path / "AsystentZotero.ppam")[1],
        reload_fn=lambda w, msg: reload_window_after_update(w, msg, popen_fn=fake_popen),
    )
    thread_up.join(timeout=2.0)
    deadline = time.monotonic() + 1.0
    while win_update.root.winfo_exists() and time.monotonic() < deadline:
        win_update.root.update()
        time.sleep(0.01)
    assert events[0] == "addin"
    assert events[1][0] == "popen"
    assert "--no-auto-update" in events[1][1]
    assert not win_update.root.winfo_exists()

    # 3. Gdy popen_fn rzuca OSError -> okienko zostaje widoczne, pasek stanu informuje o restarcie
    win_err = AzpWindow(
        zotero=FakeZoteroSource(),
        slide_doc=FakeSlideDoc(),
        master=tk_root,
        poll_interval_ms=None,
    )
    try:
        def failing_popen(*_args: object, **_kw: object) -> object:
            raise OSError("cannot spawn")

        ok_reload = reload_window_after_update(win_err, "Nowa wersja", popen_fn=failing_popen)
        assert ok_reload is False
        assert win_err.root.winfo_exists()
        assert "uruchom okienko ponownie" in win_err.get_status_message().lower()
    finally:
        win_err.destroy()

    # 4. Flaga --no-auto-update w main() pomija focus_existing_azp_window i start_background_auto_update
    skipped: list[str] = []

    class _FakeWin:
        def run(self) -> None:
            skipped.append("run")

    monkeypatch.setattr(
        "azp.main.focus_existing_azp_window",
        lambda: (skipped.append("focus"), True)[1],
    )
    monkeypatch.setattr("azp.main.is_powerpoint_addin_installed", lambda: True)
    monkeypatch.setattr("azp.main.build_window", lambda: _FakeWin())
    monkeypatch.setattr(
        "azp.main.start_background_auto_update",
        lambda _w: skipped.append("auto_update"),
    )
    assert main(["--no-auto-update"]) == 0
    assert skipped == ["run"]


def test_git_updater_and_reload_hide_console_windows_on_windows(
    tk_root: tk.Tk, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Wszystkie wywołania subprocess w updater.py i main.py muszą ukrywać okno konsoli (CREATE_NO_WINDOW + SW_HIDE)."""
    from azp.main import reload_window_after_update
    from azp.updater import auto_update_if_needed, update_from_git

    (tmp_path / ".git").mkdir()
    monkeypatch.setattr("azp.updater.shutil.which", lambda _name: "git")
    monkeypatch.setattr("azp.updater.subprocess.CREATE_NO_WINDOW", 0x08000000, raising=False)

    captured_runs: list[tuple[list[str], dict[str, object]]] = []

    def fake_run(cmd: list[str], **kwargs: object) -> subprocess.CompletedProcess[str]:
        captured_runs.append((cmd, dict(kwargs)))
        if cmd[:2] == ["git", "rev-parse"]:
            return subprocess.CompletedProcess(cmd, 0, stdout=("1" * 40) + "\n", stderr="")
        return subprocess.CompletedProcess(cmd, 0, stdout="", stderr="")

    monkeypatch.setattr("azp.updater.subprocess.run", fake_run)

    ok, _ = update_from_git(tmp_path)
    assert ok is True
    changed, _ = auto_update_if_needed(project_root=tmp_path)
    assert changed is False
    assert len(captured_runs) == 6  # 2 w update_from_git + 4 w _auto_update_git

    for _cmd, kw in captured_runs:
        assert kw.get("creationflags") == 0x08000000
        si = kw.get("startupinfo")
        if hasattr(subprocess, "STARTUPINFO"):
            assert si is not None
            assert getattr(si, "wShowWindow", None) == 0
            assert (getattr(si, "dwFlags", 0) & getattr(subprocess, "STARTF_USESHOWWINDOW", 1)) != 0
        env = kw.get("env")
        assert isinstance(env, dict)
        assert env.get("GIT_TERMINAL_PROMPT") == "0"
        assert env.get("GCM_INTERACTIVE") == "never"

    # reload_window_after_update również przekazuje creationflags=CREATE_NO_WINDOW i używa pythonw.exe
    fake_py_dir = tmp_path / "py"
    fake_py_dir.mkdir()
    fake_python = fake_py_dir / "python.exe"
    fake_pythonw = fake_py_dir / "pythonw.exe"
    fake_python.write_bytes(b"MZ")
    fake_pythonw.write_bytes(b"MZ")
    monkeypatch.setattr("azp.main.sys.executable", str(fake_python))
    monkeypatch.setattr("azp.main.subprocess.CREATE_NO_WINDOW", 0x08000000, raising=False)

    popen_calls: list[tuple[list[str], dict[str, object]]] = []

    def fake_popen(cmd: list[str], **kwargs: object) -> object:
        popen_calls.append((cmd, dict(kwargs)))
        return object()

    win = AzpWindow(
        zotero=FakeZoteroSource(),
        slide_doc=FakeSlideDoc(),
        master=tk_root,
        poll_interval_ms=None,
    )
    assert reload_window_after_update(win, "Aktualizacja", popen_fn=fake_popen) is True
    assert len(popen_calls) == 1
    popen_cmd, popen_kw = popen_calls[0]
    assert Path(popen_cmd[0]).name.lower() == "pythonw.exe"
    assert popen_kw.get("creationflags") == 0x08000000
    si_popen = popen_kw.get("startupinfo")
    if hasattr(subprocess, "STARTUPINFO"):
        assert si_popen is not None
        assert getattr(si_popen, "wShowWindow", None) == 0
        assert (getattr(si_popen, "dwFlags", 0) & getattr(subprocess, "STARTF_USESHOWWINDOW", 1)) != 0

