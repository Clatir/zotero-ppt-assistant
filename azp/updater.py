"""Aktualizacja kodu AZP z repozytorium Git lub archiwum ZIP z GitHuba."""

from __future__ import annotations

import io
import json
import re
import shutil
import subprocess
import urllib.request
import zipfile
from pathlib import Path, PurePosixPath

GITHUB_ARCHIVE_ZIP_URL = "https://github.com/Clatir/zotero-ppt-assistant/archive/refs/heads/main.zip"
GITHUB_COMMIT_SHA_URL = "https://api.github.com/repos/Clatir/zotero-ppt-assistant/commits/main"
PRESERVED_RELATIVE_PATHS = frozenset({"config.json", ".azp_version"})
_SHA_RE = re.compile(r"^[0-9a-fA-F]{40}$")


def fetch_remote_main_sha(
    sha_url: str = GITHUB_COMMIT_SHA_URL,
    timeout: float = 5.0,
) -> str | None:
    """Pobiera bieżący hash commita ``main`` z GitHub API (lub zwraca ``None`` przy błędzie)."""
    try:
        req = urllib.request.Request(
            sha_url,
            headers={
                "Accept": "application/vnd.github.sha",
                "User-Agent": "AsystentZotero-Updater",
            },
        )
        with urllib.request.urlopen(req, timeout=timeout) as response:
            raw_text = response.read().decode("utf-8", errors="replace").strip()
    except Exception:
        return None

    if _SHA_RE.fullmatch(raw_text):
        return raw_text.lower()

    try:
        payload = json.loads(raw_text)
    except ValueError:
        return None
    if isinstance(payload, dict):
        sha = payload.get("sha")
        if isinstance(sha, str) and _SHA_RE.fullmatch(sha.strip()):
            return sha.strip().lower()
    return None


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


def _collect_validated_zip_entries(
    zf: zipfile.ZipFile,
    project_root: Path,
) -> list[tuple[zipfile.ZipInfo, Path]]:
    """Waliduje wszystkie wpisy archiwum ZIP i zwraca listę ``(info, target_path)`` do zapisu."""
    root_resolved = project_root.resolve()
    pending_writes: list[tuple[zipfile.ZipInfo, Path]] = []
    seen_project_files = 0

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

    return pending_writes


def extract_github_zip_bytes(zip_bytes: bytes, project_root: Path) -> int:
    """Rozpakowuje archiwum ZIP z GitHuba do ``project_root`` i zwraca liczbę plików."""
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        pending_writes = _collect_validated_zip_entries(zf, project_root)
        for info, target_path in pending_writes:
            target_path.parent.mkdir(parents=True, exist_ok=True)
            target_path.write_bytes(zf.read(info))
    return len(pending_writes)


def update_from_github_zip(
    project_root: Path,
    zip_url: str = GITHUB_ARCHIVE_ZIP_URL,
    timeout: float = 30.0,
    sha_url: str = GITHUB_COMMIT_SHA_URL,
) -> tuple[bool, str]:
    """Pobiera archiwum ZIP z GitHuba i aktualizuje pliki projektu w ``project_root``."""
    try:
        with urllib.request.urlopen(zip_url, timeout=timeout) as response:
            data = response.read()
        count = extract_github_zip_bytes(data, project_root)
        remote_sha = fetch_remote_main_sha(sha_url=sha_url, timeout=min(timeout, 5.0))
        if remote_sha is not None:
            try:
                (project_root / ".azp_version").write_text(remote_sha, encoding="utf-8")
            except OSError:
                pass
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


def _auto_update_git(root: Path, timeout: float) -> tuple[bool, str]:
    try:
        status = subprocess.run(
            ["git", "status", "--porcelain", "--untracked-files=no"],
            cwd=root,
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except (OSError, subprocess.SubprocessError):
        return False, "Pominięto auto-aktualizację (lokalne zmiany lub błąd git status)."

    if status.returncode != 0 or status.stdout.strip():
        return False, "Pominięto auto-aktualizację (lokalne zmiany lub błąd git status)."

    try:
        old_head = subprocess.run(
            ["git", "rev-parse", "HEAD"],
            cwd=root,
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except (OSError, subprocess.SubprocessError):
        return False, "Nie udało się odczytać HEAD."

    if old_head.returncode != 0:
        return False, "Nie udało się odczytać HEAD."
    old_sha = old_head.stdout.strip()

    try:
        pull = subprocess.run(
            ["git", "pull", "--ff-only"],
            cwd=root,
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except (OSError, subprocess.SubprocessError):
        return False, "git pull nie powiódł się."

    if pull.returncode != 0:
        return False, "git pull nie powiódł się."

    try:
        new_head = subprocess.run(
            ["git", "rev-parse", "HEAD"],
            cwd=root,
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        new_sha = new_head.stdout.strip() if new_head.returncode == 0 else ""
    except (OSError, subprocess.SubprocessError):
        new_sha = ""

    if new_sha and old_sha != new_sha:
        return True, "Pobrano nową wersję z GitHuba (git pull)."
    return False, "Wersja jest aktualna."


def auto_update_if_needed(
    project_root: Path | None = None,
    zip_url: str = GITHUB_ARCHIVE_ZIP_URL,
    sha_url: str = GITHUB_COMMIT_SHA_URL,
    timeout: float = 10.0,
) -> tuple[bool, str]:
    """Sprawdza aktualizacje w tle; zwraca ``(True, msg)`` tylko po zmianie plików."""
    root = project_root if project_root is not None else Path(__file__).resolve().parent.parent
    if (root / ".git").exists() and shutil.which("git"):
        return _auto_update_git(root, timeout=timeout)

    remote_sha = fetch_remote_main_sha(sha_url=sha_url, timeout=min(timeout, 5.0))
    version_file = root / ".azp_version"
    if remote_sha is not None and version_file.is_file():
        try:
            if version_file.read_text(encoding="utf-8").strip() == remote_sha:
                return False, "Wersja jest aktualna."
        except OSError:
            pass

    try:
        with urllib.request.urlopen(zip_url, timeout=timeout) as response:
            zip_bytes = response.read()
        changed_count = 0
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            pending_writes = _collect_validated_zip_entries(zf, root)
            for info, target_path in pending_writes:
                new_bytes = zf.read(info)
                if not target_path.exists() or target_path.read_bytes() != new_bytes:
                    target_path.parent.mkdir(parents=True, exist_ok=True)
                    target_path.write_bytes(new_bytes)
                    changed_count += 1
    except Exception as exc:
        return False, f"Nie udało się sprawdzić aktualizacji: {exc}"

    if remote_sha is not None:
        try:
            version_file.write_text(remote_sha, encoding="utf-8")
        except OSError:
            pass

    if changed_count > 0:
        return True, f"Pobrano nową wersję z GitHuba ({changed_count} zmienionych plików)."
    return False, "Wersja jest aktualna."


__all__ = [
    "GITHUB_ARCHIVE_ZIP_URL",
    "GITHUB_COMMIT_SHA_URL",
    "PRESERVED_RELATIVE_PATHS",
    "auto_update_if_needed",
    "extract_github_zip_bytes",
    "fetch_remote_main_sha",
    "update_from_git",
    "update_from_github_zip",
    "update_project_code",
]
