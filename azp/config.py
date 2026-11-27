"""Obsługa trwałej konfiguracji użytkownika w pliku `config.json` (etap C2)."""

from __future__ import annotations

import json
from pathlib import Path

from azp.contracts import Settings

DEFAULT_CONFIG_PATH: Path = Path(__file__).resolve().parent.parent / "config.json"


def load_config(path: Path | None = None) -> Settings:
    """Wczytuje ustawienia z pliku JSON; przy błędzie zwraca domyślne `Settings()`."""
    target = path if path is not None else DEFAULT_CONFIG_PATH
    try:
        raw = target.read_text(encoding="utf-8")
        if not raw.strip():
            return Settings()
        data = json.loads(raw)
        if not isinstance(data, dict):
            return Settings()
        return Settings.from_dict(data)
    except (OSError, ValueError, TypeError):
        return Settings()


def save_config(settings: Settings, path: Path | None = None) -> Path:
    """Zapisuje ustawienia do pliku JSON w kodowaniu UTF-8 i zwraca ścieżkę."""
    target = path if path is not None else DEFAULT_CONFIG_PATH
    target.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(settings.to_dict(), indent=2, ensure_ascii=False) + "\n"
    target.write_text(payload, encoding="utf-8")
    return target


__all__ = [
    "DEFAULT_CONFIG_PATH",
    "load_config",
    "save_config",
]
