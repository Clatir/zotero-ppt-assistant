"""Aktualizacja kodu AZP z repozytorium Git lub archiwum ZIP z GitHuba."""

from __future__ import annotations

import io
import shutil
import subprocess
import urllib.request
import zipfile
from pathlib import Path, PurePosixPath

GITHUB_ARCHIVE_ZIP_URL = "https://github.com/Clatir/zotero-ppt-assistant/archive/refs/heads/main.zip"
PRESERVED_RELATIVE_PATHS = frozenset({"config.json"})


def update_from_git(project_root: Path) -> tuple[bool, str]:
    """Aktualizuje repozytorium przez ``git pull --ff-only``, jeśli brak lokalnych zmian."""
    if not (project_root / ".git").exists() or not shutil.which("git"):
        return False, "Brak repozytorium Git lub programu git."

    try:
        status = subprocess.run(
            ["git", "status", "--porcelain", "--untracked-files=no"],
            cwd=project_root,
            capture_output=True,
            text=True,
            timeout=30,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        return False, f"git status nie powiódł się: {exc}"

    if status.returncode != 0:
        err = (status.stderr or status.stdout or "").strip()
        return False, f"git status nie powiódł się: {err}"

    if status.stdout.strip():
        return (
            True,
            "Wykryto lokalne zmiany w repozytorium Git – pominięto git pull, używam bieżących plików.",
        )

    try:
        pull = subprocess.run(
            ["git", "pull", "--ff-only"],
            cwd=project_root,
            capture_output=True,
            text=True,
            timeout=30,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        return False, f"git pull nie powiódł się: {exc}"

    if pull.returncode == 0:
        return True, "Zaktualizowano kod przez git pull."

    stderr_or_stdout = (pull.stderr or pull.stdout or "").strip()
    return False, f"git pull nie powiódł się: {stderr_or_stdout}"


def extract_github_zip_bytes(zip_bytes: bytes, project_root: Path) -> int:
    """Rozpakowuje archiwum ZIP z GitHuba do ``project_root`` i zwraca liczbę plików."""
    root_resolved = project_root.resolve()
    pending_writes: list[tuple[zipfile.ZipInfo, Path]] = []
    seen_project_files = 0

    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for info in zf.infolist():
            raw_name = info.filename.replace("\\", "/")
            if not raw_name:
                continue

            posix_path = PurePosixPath(raw_name)
            if posix_path.is_absolute() or Path(raw_name).is_absolute():
                raise ValueError(f"Niebezpieczna ścieżka w archiwum ZIP: {info.filename}")

            parts = [part for part in posix_path.parts if part not in ("", ".")]
            if not parts:
                continue
            if ".." in parts or any(":" in part for part in parts):
                raise ValueError(f"Niebezpieczna ścieżka w archiwum ZIP: {info.filename}")

            if info.is_dir() or raw_name.endswith("/"):
                continue

            # Archiwum z GitHuba ma prefiks pierwszego katalogu, np. zotero-ppt-assistant-main/...
            if len(parts) < 2:
                continue

            rel_parts = parts[1:]
            if ".git" in rel_parts:
                continue

            target_path = (root_resolved / Path(*rel_parts)).resolve()
            if target_path == root_resolved or not target_path.is_relative_to(root_resolved):
                raise ValueError(f"Niebezpieczna ścieżka w archiwum ZIP: {info.filename}")

            seen_project_files += 1
            rel_posix = "/".join(rel_parts)
            if rel_posix in PRESERVED_RELATIVE_PATHS and target_path.exists():
                continue

            pending_writes.append((info, target_path))

        if seen_project_files == 0:
            raise ValueError("Archiwum ZIP nie zawiera plików projektu.")

        for info, target_path in pending_writes:
            target_path.parent.mkdir(parents=True, exist_ok=True)
            target_path.write_bytes(zf.read(info))

    return len(pending_writes)


def update_from_github_zip(
    project_root: Path,
    zip_url: str = GITHUB_ARCHIVE_ZIP_URL,
    timeout: float = 30.0,
) -> tuple[bool, str]:
    """Pobiera archiwum ZIP z GitHuba i aktualizuje pliki projektu w ``project_root``."""
    try:
        with urllib.request.urlopen(zip_url, timeout=timeout) as response:
            data = response.read()
        count = extract_github_zip_bytes(data, project_root)
        return True, f"Pobrano i zaktualizowano {count} plików z GitHuba."
    except Exception as exc:
        return False, f"Nie udało się pobrać aktualizacji z GitHuba: {exc}"


def update_project_code(
    project_root: Path | None = None,
    zip_url: str = GITHUB_ARCHIVE_ZIP_URL,
) -> tuple[bool, str]:
    """Aktualizuje kod projektu przez Git lub (gdy brak repozytorium Git) przez ZIP z GitHuba."""
    root = project_root if project_root is not None else Path(__file__).resolve().parent.parent
    if (root / ".git").exists() and shutil.which("git"):
        return update_from_git(root)
    return update_from_github_zip(root, zip_url=zip_url)


__all__ = [
    "GITHUB_ARCHIVE_ZIP_URL",
    "PRESERVED_RELATIVE_PATHS",
    "extract_github_zip_bytes",
    "update_from_git",
    "update_from_github_zip",
    "update_project_code",
]
