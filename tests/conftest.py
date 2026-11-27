"""Wspólne fikstury pytest dla testów AZP."""

from __future__ import annotations

import tkinter as tk

import pytest


@pytest.fixture(scope="session")
def tk_root():
    """Pojedynczy ukryty korzeń `tk.Tk` współdzielony przez wszystkie moduły testowe."""
    root = tk.Tk()
    root.withdraw()
    yield root
    try:
        root.destroy()
    except tk.TclError:
        pass
