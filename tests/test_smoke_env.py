"""Testy dymne środowiska (Etap A1): lokalne API Zotero 10 oraz COM PowerPointa."""

from __future__ import annotations

from pathlib import Path
import subprocess

import pytest
import requests

ZOTERO_API_BASE = "http://127.0.0.1:23119/api/users/0"
STYLE_IEEE_V4 = "http://www.zotero.org/styles/ieee-praca-inzynierska-access-date-v4"
STYLE_APA = "http://www.zotero.org/styles/apa"
TEST_PPTX_PATH = Path(__file__).resolve().parent / "data" / "test_smoke.pptx"


def _is_powerpoint_running() -> bool:
    """Sprawdza, czy proces POWERPNT.EXE jest uruchomiony."""
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


def test_zotero_local_api_responds_ieee_v4_and_apa() -> None:
    """Sprawdza, czy lokalne API Zotero 10 odpowiada i formatuje w stylu v4 oraz APA."""
    for style_id, expected_marker in ((STYLE_IEEE_V4, "[1]"), (STYLE_APA, "csl-entry")):
        resp = requests.get(
            f"{ZOTERO_API_BASE}/items/top",
            params={
                "limit": 1,
                "include": "data,bib,citation",
                "style": style_id,
            },
            timeout=5,
        )
        assert resp.status_code == 200, f"Zotero API zwróciło {resp.status_code}: {resp.text}"
        items = resp.json()
        assert isinstance(items, list) and len(items) >= 1
        item = items[0]
        assert "key" in item
        assert "bib" in item and expected_marker in item["bib"]
        assert "citation" in item and len(item["citation"].strip()) > 0


@pytest.mark.com
def test_powerpoint_com_open_tag_save_and_close() -> None:
    """Sprawdza utworzenie, zapis, ponowne otwarcie i zamknięcie prezentacji testowej przez COM."""
    if not _ensure_powerpoint_not_running():
        pytest.fail("PowerPoint jest uruchomiony! Zamknij PowerPointa przed testami COM.")

    import pythoncom
    import win32com.client

    TEST_PPTX_PATH.parent.mkdir(parents=True, exist_ok=True)
    if TEST_PPTX_PATH.exists():
        TEST_PPTX_PATH.unlink()

    pythoncom.CoInitialize()
    app = None
    pres = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        # Tworzymy nową prezentację bez okna (WithWindow=False = 0)
        pres = app.Presentations.Add(0)
        pres.Tags.Add("AZP_SMOKE", "V1")

        # Dodajemy pusty slajd (ppLayoutBlank = 12) i pole tekstowe
        slide = pres.Slides.Add(1, 12)
        slide.Tags.Add("AZP_SLIDE_TAG", "SLIDE_1")
        shape = slide.Shapes.AddTextbox(1, 50, 50, 400, 100)
        shape.TextFrame.TextRange.Text = "Test dymny AZP [1]"
        shape.Tags.Add("AZP_SHAPE_TAG", "FOOTER_1")

        pres.SaveAs(str(TEST_PPTX_PATH))
        pres.Close()
        pres = None

        assert TEST_PPTX_PATH.exists()

        # Ponowne otwarcie i weryfikacja trwałości slajdu, tekstu oraz Tags
        pres = app.Presentations.Open(str(TEST_PPTX_PATH), ReadOnly=False, Untitled=False, WithWindow=0)
        assert pres.Slides.Count == 1
        assert pres.Tags("AZP_SMOKE") == "V1"
        reopened_slide = pres.Slides(1)
        assert reopened_slide.Tags("AZP_SLIDE_TAG") == "SLIDE_1"
        reopened_shape = reopened_slide.Shapes(1)
        assert reopened_shape.TextFrame.TextRange.Text == "Test dymny AZP [1]"
        assert reopened_shape.Tags("AZP_SHAPE_TAG") == "FOOTER_1"
        pres.Close()
        pres = None
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
