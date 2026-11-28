"""Testy uruchamiania okienka z prawdziwymi zależnościami (`azp.main`, etap C1) – na atrapach."""

from __future__ import annotations

import importlib.util
import tkinter as tk

import pytest

from azp.contracts import Citation, NoTextCursorError
from azp.main import build_window, describe_error
from tests.fakes.fake_ppt import FakeSlideDoc
from tests.fakes.fake_zotero import FakeZoteroSource


class _FocusRecordingDoc(FakeSlideDoc):
    def __init__(self, **kwargs: object) -> None:
        super().__init__(**kwargs)  # type: ignore[arg-type]
        self.focus_calls = 0
        self.fail_insert_with: Exception | None = None

    def focus_presentation_window(self) -> None:
        self.focus_calls += 1

    def insert_citation_at_cursor(self, citation: Citation) -> None:
        if self.fail_insert_with is not None:
            raise self.fail_insert_with
        super().insert_citation_at_cursor(citation)


def test_main_module_is_runnable_with_python_m_azp() -> None:
    assert importlib.util.find_spec("azp.__main__") is not None


def test_build_window_wires_engine_and_returns_focus_after_cite_d20(tk_root: tk.Tk) -> None:
    doc = _FocusRecordingDoc()
    win = build_window(
        zotero=FakeZoteroSource(),
        slide_doc=doc,
        master=tk_root,
        poll_interval_ms=None,
        config_path=None,
    )
    try:
        win.select_result_keys(["KOWAL2023", "NOWAK2024"])
        assert win.cite_selected() is True
        # Silnik (CitationEngine) wstawia cytowanie zbiorcze i stopkę
        assert doc.read_slides()[0].citations[0].text == "[1, 2]"
        assert [e.key for e in doc.read_slides()[0].footer_entries] == ["KOWAL2023", "NOWAK2024"]
        assert doc.focus_calls == 1

        # Bez kursora na aktywnym slajdzie (D26): przypis trafia do stopki na dole slajdu
        doc.active_slide_index = 2
        doc.has_cursor = False
        win.select_result_keys(["DOVE2013"])
        assert win.cite_selected() is True
        assert doc.get_slide_body_text(2) == "Tekst slajdu 2"
        assert [e.key for e in doc.read_slides()[1].footer_entries] == ["DOVE2013"]
        assert doc.focus_calls == 2

        # Brak aktywnego slajdu: brak zmian i brak przełączania okna
        doc.active_slide_index = None
        assert win.cite_selected() is False
        assert doc.focus_calls == 2
        doc.active_slide_index = 1

        # Wstawienie bibliografii przez okienko (D27)
        assert win.insert_bibliography() is True
        assert [e.key for e in doc.get_slide_bibliography_entries(1)] == [
            "KOWAL2023",
            "NOWAK2024",
            "DOVE2013",
        ]

        # Przebudowa przez silnik – z kopią zapasową
        assert win.refresh_presentation() is True
        assert len(doc.backups_created) == 1
        assert "kopia zapasowa" in win.get_status_message()
    finally:
        win.destroy()


def test_build_window_turns_unexpected_powerpoint_errors_into_status(tk_root: tk.Tk) -> None:
    """Błąd COM (np. PowerPoint zajęty) nie wysypuje okienka – komunikat w pasku stanu."""
    doc = _FocusRecordingDoc()
    win = build_window(
        zotero=FakeZoteroSource(),
        slide_doc=doc,
        master=tk_root,
        poll_interval_ms=None,
        config_path=None,
    )
    try:
        doc.fail_insert_with = RuntimeError("Call was rejected by callee.")
        win.select_result_keys(["KOWAL2023"])
        assert win.cite_selected() is False
        msg = win.get_status_message()
        assert msg.startswith("PowerPoint zgłosił błąd")
        assert "Call was rejected by callee." in msg
        assert doc.focus_calls == 0

        # Wyjątki domenowe przechodzą bez zmian (D10)
        doc.fail_insert_with = NoTextCursorError("x")
        assert win.cite_selected() is False
        assert win.get_status_message() == "Ustaw kursor w polu tekstowym na slajdzie."
    finally:
        win.destroy()


def test_describe_error_reads_com_error_description() -> None:
    class _FakeComError(Exception):
        pass

    exc = _FakeComError(
        -2147352567,
        "Wystąpił wyjątek.",
        (0, "Microsoft PowerPoint", "Nieprawidłowa wartość zakresu.", None, 0, -2147024809),
        None,
    )
    assert describe_error(exc) == "Nieprawidłowa wartość zakresu."
    assert describe_error(RuntimeError("Zwykły błąd")) == "Zwykły błąd"
    assert describe_error(RuntimeError()) == "RuntimeError"



# ---------------------------------------------------------------------------
# Test końcowy C1 (@pytest.mark.com): prawdziwe okienko + prawdziwe Zotero + prawdziwy PowerPoint
# ---------------------------------------------------------------------------


def _is_powerpoint_running() -> bool:
    import subprocess

    res = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"],
        capture_output=True,
        text=True,
        check=False,
    )
    return "powerpnt.exe" in res.stdout.lower()


def _ensure_powerpoint_not_running(timeout_seconds: float = 6.0) -> bool:
    import time

    start = time.time()
    while time.time() - start < timeout_seconds:
        if not _is_powerpoint_running():
            return True
        time.sleep(0.3)
    return not _is_powerpoint_running()


@pytest.mark.com
def test_c1_window_drives_real_powerpoint_and_real_zotero_e2e(tk_root: tk.Tk) -> None:
    """Kompleksowy test końcowy C1: `AzpWindow` -> `CitationEngine` -> `ZoteroClient` + `PowerPointSlideDoc`."""
    if not _ensure_powerpoint_not_running():
        pytest.fail("PowerPoint jest uruchomiony! Zamknij PowerPointa przed testami COM.")

    from pathlib import Path
    import pythoncom
    import win32com.client

    from azp.contracts import MSG_SAVE_BEFORE_REBUILD, STYLE_APA, STYLE_IEEE_V4
    from azp.ppt import PowerPointSlideDoc, TAG_ROLE, TAG_ROLE_FOOTER
    from azp.zotero import ZoteroClient

    data_dir = Path(__file__).resolve().parent / "data"
    data_dir.mkdir(parents=True, exist_ok=True)
    pres_a_path = data_dir / "test_c1_e2e_a.pptx"
    pres_a_bak = data_dir / "test_c1_e2e_a.bak.pptx"
    pres_b_path = data_dir / "test_c1_e2e_b.pptx"
    pres_b_bak = data_dir / "test_c1_e2e_b.bak.pptx"
    for p in (pres_a_path, pres_a_bak, pres_b_path, pres_b_bak):
        if p.exists():
            p.unlink()

    pythoncom.CoInitialize()
    app = None
    pres_a = None
    pres_b = None
    doc = None
    win = None
    try:
        zotero = ZoteroClient()
        assert zotero.is_available() is True, "Lokalne API Zotero 10 musi być włączone do testu C1."

        app = win32com.client.DispatchEx("PowerPoint.Application")
        pres_a = app.Presentations.Add(0)

        slide1 = pres_a.Slides.Add(1, 12)
        box1 = slide1.Shapes.AddTextbox(1, 60, 120, 600, 120)
        box1.Name = "BodyBox1"
        box1.TextFrame.TextRange.Text = "Wstęp Ważne i wniosek"
        # Pogrubiamy słowo "Ważne" (znaki 7..11, 1-based), aby sprawdzić zachowanie formatowania (T10)
        box1.TextFrame.TextRange.Characters(7, 5).Font.Bold = -1

        slide2 = pres_a.Slides.Add(2, 12)
        box2 = slide2.Shapes.AddTextbox(1, 60, 120, 600, 120)
        box2.Name = "BodyBox2"
        box2.TextFrame.TextRange.Text = "Drugi slajd: metodyka"

        doc = PowerPointSlideDoc(app=app, presentation=pres_a, active_slide_index=1)
        win = build_window(
            zotero=zotero,
            slide_doc=doc,
            master=tk_root,
            poll_interval_ms=None,
            config_path=None,
        )

        items = win.get_result_items()
        assert len(items) >= 2, "W bibliotece Zotero muszą znajdować się co najmniej 2 pozycje."
        key1, key2 = items[0].key, items[1].key

        # 1. Cytowanie na niezapisanej prezentacji działa, ale przebudowa jest zablokowana (D18)
        win.select_result_keys([key1])
        doc.set_cursor_target(1, "BodyBox1", len("Wstęp Ważne"))
        assert win.cite_selected() is True
        assert box1.TextFrame.TextRange.Text == "Wstęp Ważne [1] i wniosek"
        # Pogrubienie słowa "Ważne" zostało zachowane (T10)!
        assert int(box1.TextFrame.TextRange.Characters(7, 5).Font.Bold) == -1

        assert win.set_style(STYLE_APA) is False
        assert win.get_status_message() == MSG_SAVE_BEFORE_REBUILD
        assert win.controller.settings.style_id == STYLE_IEEE_V4
        assert win.refresh_presentation() is False
        assert win.get_status_message() == MSG_SAVE_BEFORE_REBUILD

        # 3. Automatyczne scalenie sąsiedniego cytowania [1] + [2] -> [1, 2] (D13)
        win.select_result_keys([key2])
        assert win.cite_selected() is True
        assert box1.TextFrame.TextRange.Text == "Wstęp Ważne [1, 2] i wniosek"
        assert int(box1.TextFrame.TextRange.Characters(7, 5).Font.Bold) == -1

        # 4. Cytowanie na slajdzie 2: powtórzenie key2 -> [2] w trybie globalnym
        doc.set_cursor_target(2, "BodyBox2", None)
        win.select_result_keys([key2])
        assert win.cite_selected() is True
        assert box2.TextFrame.TextRange.Text == "Drugi slajd: metodyka [2]"

        # 5. Zapis prezentacji na dysk -> przebudowa i SaveCopyAs bez zapisywania oryginału (D17, D14)
        pres_a.SaveAs(str(pres_a_path))
        mtime_before = pres_a_path.stat().st_mtime_ns
        # Dodatkowa zmiana w pamięci po SaveAs, żeby sprawdzić, że create_backup nie woła Save() na oryginale
        box2.TextFrame.TextRange.Characters(1, 5).Font.Italic = -1

        assert win.set_numbering_mode("per_slide") is True
        assert pres_a_bak.exists()
        assert box2.TextFrame.TextRange.Text == "Drugi slajd: metodyka [1]"
        assert int(box1.TextFrame.TextRange.Characters(7, 5).Font.Bold) == -1
        # Oryginalny plik na dysku nie został nadpisany przez przebudowę (D17)
        assert pres_a_path.stat().st_mtime_ns == mtime_before

        # 6. Przełączenie na styl APA 7 oraz zmiana czcionki i położenia stopki (D3, D5, D9, D14)
        assert win.set_style(STYLE_APA) is True
        apa_text_s1 = str(box1.TextFrame.TextRange.Text)
        assert "[1, 2]" not in apa_text_s1
        assert "(" in apa_text_s1 and ";" in apa_text_s1 and ")" in apa_text_s1
        assert int(box1.TextFrame.TextRange.Characters(7, 5).Font.Bold) == -1

        assert win.set_font_name("Arial") is True
        assert win.set_font_size(11) is True
        assert win.set_footer_position("bottom_left") is True
        assert win.set_footer_header(False) is True

        footer_s1 = None
        for i in range(1, int(slide1.Shapes.Count) + 1):
            sh = slide1.Shapes(i)
            if sh.Tags(TAG_ROLE) == TAG_ROLE_FOOTER:
                footer_s1 = sh
                break
        assert footer_s1 is not None
        assert not str(footer_s1.TextFrame.TextRange.Text).startswith("Źródła:")
        assert footer_s1.TextFrame.TextRange.Font.Name == "Arial"
        assert int(footer_s1.TextFrame.TextRange.Font.Size) == 11

        # 7. Powrót do IEEE v4 globalnie, zamiana kolejności slajdów i Odśwież / Przenumeruj
        assert win.set_style(STYLE_IEEE_V4) is True
        assert win.set_numbering_mode("global") is True
        assert box1.TextFrame.TextRange.Text == "Wstęp Ważne [1, 2] i wniosek"
        assert box2.TextFrame.TextRange.Text == "Drugi slajd: metodyka [2]"

        pres_a.Slides(2).MoveTo(1)
        assert win.refresh_presentation() is True
        # Teraz dawny slajd 2 jest slajdem 1, więc key2 dostaje numer [1], a na slajdzie 2 zostaje [1, 2]
        assert box2.TextFrame.TextRange.Text == "Drugi slajd: metodyka [1]"
        assert box1.TextFrame.TextRange.Text == "Wstęp Ważne [1, 2] i wniosek"
        assert int(box1.TextFrame.TextRange.Characters(7, 5).Font.Bold) == -1

        # 8. Śledzenie aktywnej prezentacji przy dwóch otwartych plikach (D19)
        pres_b = app.Presentations.Add(0)
        pres_b.Slides.Add(1, 12)
        pres_b.SaveAs(str(pres_b_path))

        # Przełączamy PowerPointSlideDoc na dynamiczny wybór prezentacji z app
        doc._presentation = pres_b
        win.check_active_presentation()
        assert win.set_style(STYLE_APA) is True

        doc._presentation = pres_a
        assert win.check_active_presentation() is True
        assert win.controller.settings.style_id == STYLE_IEEE_V4

        doc._presentation = pres_b
        assert win.check_active_presentation() is True
        assert win.controller.settings.style_id == STYLE_APA
    finally:
        if win is not None:
            win.destroy()
        slide1 = slide2 = box1 = box2 = footer_s1 = None
        if doc is not None:
            doc.invalidate_cache()
            doc._presentation = None
            doc._app = None
            doc = None
        for p_obj in (pres_b, pres_a):
            if p_obj is not None:
                try:
                    p_obj.Saved = True
                    p_obj.Close()
                except Exception:
                    pass
        pres_a = pres_b = None
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        app = None
        import gc
        import time

        gc.collect()
        pythoncom.CoUninitialize()
        for _ in range(25):
            if not _is_powerpoint_running():
                break
            time.sleep(0.2)
        for p in (pres_a_path, pres_a_bak, pres_b_path, pres_b_bak):
            if p.exists():
                try:
                    p.unlink()
                except Exception:
                    pass
