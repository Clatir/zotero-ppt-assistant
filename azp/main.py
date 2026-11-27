"""Punkt sklejenia aplikacji AZP (`build_window`, `main` – Etap C1)."""

from __future__ import annotations

import sys
import tkinter as tk
from pathlib import Path
from typing import Sequence

from azp.config import DEFAULT_CONFIG_PATH
from azp.contracts import AzpError, Settings, SlideDoc, ZoteroSource
from azp.engine import CitationEngine
from azp.gui import AzpWindow
from azp.ppt import (
    PowerPointSlideDoc,
    install_powerpoint_addin,
    is_powerpoint_addin_installed,
    uninstall_powerpoint_addin,
)
from azp.zotero import ZoteroClient


def describe_error(exc: BaseException) -> str:
    """Wyciąga czytelny opis błędu (w tym `excepinfo` z błędów COM `pywintypes.com_error`)."""
    args = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 3:
        desc = args[2][2]
        if isinstance(desc, str) and desc.strip():
            return desc.strip()
    if len(args) >= 2 and isinstance(args[1], str) and args[1].strip():
        return args[1].strip()
    msg = str(exc).strip()
    if msg:
        return msg
    return type(exc).__name__


def build_window(
    zotero: ZoteroSource | None = None,
    slide_doc: SlideDoc | None = None,
    settings: Settings | None = None,
    master: tk.Misc | None = None,
    initial_search: bool = True,
    poll_interval_ms: int | None = 500,
    config_path: Path | None = DEFAULT_CONFIG_PATH,
) -> AzpWindow:
    """Tworzy okienko `AzpWindow` połączone z `ZoteroClient`, `PowerPointSlideDoc` i `CitationEngine`."""
    zot: ZoteroSource = zotero if zotero is not None else ZoteroClient()
    doc: SlideDoc = slide_doc if slide_doc is not None else PowerPointSlideDoc()
    engine = CitationEngine(zotero=zot, slide_doc=doc)

    raw_on_cite = engine.make_cite_callback()
    raw_on_rebuild = engine.make_rebuild_callback()

    def _on_cite(keys: tuple[str, ...], current_settings: Settings) -> str | None:
        try:
            msg = raw_on_cite(keys, current_settings)
        except AzpError:
            raise
        except Exception as exc:
            raise AzpError(f"PowerPoint zgłosił błąd: {describe_error(exc)}") from exc

        focus_fn = getattr(doc, "focus_presentation_window", None)
        if callable(focus_fn):
            try:
                focus_fn()
            except Exception:
                pass
        return msg

    def _on_rebuild(current_settings: Settings) -> str | None:
        try:
            return raw_on_rebuild(current_settings)
        except AzpError:
            raise
        except Exception as exc:
            raise AzpError(f"PowerPoint zgłosił błąd: {describe_error(exc)}") from exc

    return AzpWindow(
        zotero=zot,
        slide_doc=doc,
        settings=settings,
        master=master,
        on_cite=_on_cite,
        on_rebuild=_on_rebuild,
        initial_search=initial_search,
        poll_interval_ms=poll_interval_ms,
        config_path=config_path,
    )


def focus_existing_azp_window() -> bool:
    """Jeśli okienko AZP jest już otwarte, przywraca je na wierzch i zwraca `True`."""
    try:
        import win32con
        import win32gui

        found: list[int] = []

        def _cb(hwnd: int, _extra: object) -> None:
            if not win32gui.IsWindowVisible(hwnd):
                return
            title = (win32gui.GetWindowText(hwnd) or "").strip()
            if title == "Asystent Zotero dla PowerPointa":
                found.append(hwnd)

        win32gui.EnumWindows(_cb, None)
        if found:
            target = found[0]
            try:
                if win32gui.IsIconic(target):
                    win32gui.ShowWindow(target, win32con.SW_RESTORE)
                else:
                    win32gui.ShowWindow(target, win32con.SW_SHOW)
            except Exception:
                pass
            try:
                win32gui.SetForegroundWindow(target)
            except Exception:
                pass
            return True
    except Exception:
        pass
    return False


def main(argv: Sequence[str] | None = None) -> int:
    """Uruchamia pływające okienko AZP lub instalator dodatku wstążki PowerPointa."""
    args = list(sys.argv[1:] if argv is None else argv)
    if "--install-addin" in args:
        ppam = install_powerpoint_addin()
        print(f"Zainstalowano dodatek wstążki PowerPointa: {ppam}")
        return 0
    if "--uninstall-addin" in args:
        uninstall_powerpoint_addin()
        print("Odinstalowano dodatek wstążki PowerPointa.")
        return 0

    if focus_existing_azp_window():
        return 0

    try:
        if not is_powerpoint_addin_installed():
            install_powerpoint_addin()
    except Exception:
        pass

    window = build_window()
    window.run()
    return 0


__all__ = [
    "build_window",
    "describe_error",
    "focus_existing_azp_window",
    "main",
]
