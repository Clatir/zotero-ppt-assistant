"""Generator i instalator dodatku wstążki PowerPointa (`AsystentZotero.ppam`).

Tworzy natywny dodatek OpenXML PowerPointa (`.ppam`) z przyciskiem
„Asystent Zotero” na karcie **Dodatki** (`TabAddIns`) oraz **Narzędzia główne** (`TabHome`)
na wstążce PowerPointa i rejestruje go w `HKCU\\Software\\Microsoft\\Office\\<ver>\\PowerPoint\\AddIns`.
"""

from __future__ import annotations

import base64
import os
import sys
import zipfile
import zlib
from pathlib import Path

ADDIN_NAME = "AsystentZotero"
ADDIN_FILENAME = "AsystentZotero.ppam"
TRUSTED_LOCATION_KEY_NAME = "AsystentZoteroAddIn"
OFFICE_VERSIONS: tuple[str, ...] = ("16.0", "15.0", "14.0")

# Skompilowany przez PowerPoint projekt VBA (`ppt/vbaProject.bin`) zawierający moduł `AzpLauncher`
# z procedurami `OnLaunchAzp(ByVal control As Object)` oraz `LaunchAzp()`.
_VBA_PROJECT_BIN_ZLIB_B64 = (
    "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"
    "L/wvH9vxOw=="
)

_CONTENT_TYPES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/ppt/presentation.xml" '
    'ContentType="application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"/>'
    "</Types>"
)

_ROOT_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="ppt/presentation.xml"/>'
    '<Relationship Id="rId2" '
    'Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" '
    'Target="customUI/customUI.xml"/>'
    '<Relationship Id="rId3" '
    'Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" '
    'Target="customUI/customUI14.xml"/>'
    "</Relationships>"
)

_PRESENTATION_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
    'Target="vbaProject.bin"/>'
    "</Relationships>"
)

_CUSTOM_UI_BODY = (
    "  <ribbon>\r\n"
    "    <tabs>\r\n"
    '      <tab idMso="TabAddIns">\r\n'
    '        <group id="AzpAddInsGroup" label="Asystent Zotero">\r\n'
    '          <button id="AzpLaunchButtonAddIns"\r\n'
    '                  label="Asystent Zotero"\r\n'
    '                  size="large"\r\n'
    '                  imageMso="CitationInsert"\r\n'
    '                  screentip="Uruchom Asystenta Zotero dla PowerPointa"\r\n'
    '                  supertip="Otwiera pływające okienko Asystenta Zotero (AZP) do wstawiania cytowań i źródeł na slajdach."\r\n'
    '                  onAction="OnLaunchAzp" />\r\n'
    "        </group>\r\n"
    "      </tab>\r\n"
    '      <tab idMso="TabHome">\r\n'
    '        <group id="AzpHomeGroup" label="Zotero">\r\n'
    '          <button id="AzpLaunchButtonHome"\r\n'
    '                  label="Asystent Zotero"\r\n'
    '                  size="large"\r\n'
    '                  imageMso="CitationInsert"\r\n'
    '                  screentip="Uruchom Asystenta Zotero dla PowerPointa"\r\n'
    '                  supertip="Otwiera pływające okienko Asystenta Zotero (AZP) do wstawiania cytowań i źródeł na slajdach."\r\n'
    '                  onAction="OnLaunchAzp" />\r\n'
    "        </group>\r\n"
    "      </tab>\r\n"
    "    </tabs>\r\n"
    "  </ribbon>\r\n"
)

_CUSTOM_UI_2006_XML = (
    '<?xml version="1.0" encoding="UTF-8"?>\r\n'
    '<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">\r\n'
    f"{_CUSTOM_UI_BODY}"
    "</customUI>\r\n"
)

_CUSTOM_UI_2009_XML = (
    '<?xml version="1.0" encoding="UTF-8"?>\r\n'
    '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">\r\n'
    f"{_CUSTOM_UI_BODY}"
    "</customUI>\r\n"
)


def _default_project_root() -> Path:
    return Path(__file__).resolve().parent.parent.parent


def _default_addin_dir() -> Path:
    appdata = os.environ.get("APPDATA")
    base = Path(appdata) if appdata else (Path.home() / "AppData" / "Roaming")
    return base / "Microsoft" / "AddIns"


def _default_azp_appdata_dir() -> Path:
    appdata = os.environ.get("APPDATA")
    base = Path(appdata) if appdata else (Path.home() / "AppData" / "Roaming")
    return base / "AsystentZotero"


def _resolve_pythonw(python_executable: Path | None = None) -> Path:
    exe = Path(python_executable) if python_executable is not None else Path(sys.executable)
    candidate = exe.with_name("pythonw.exe")
    if candidate.exists():
        return candidate
    return exe


def build_ppam_addin(output_path: Path) -> Path:
    """Buduje plik dodatku PowerPointa `.ppam` z wstążką CustomUI i modułem VBA."""
    target = Path(output_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    vba_bytes = zlib.decompress(base64.b64decode(_VBA_PROJECT_BIN_ZLIB_B64))

    with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES_XML.encode("utf-8"))
        zf.writestr("_rels/.rels", _ROOT_RELS_XML.encode("utf-8"))
        zf.writestr("ppt/presentation.xml", b"")
        zf.writestr(
            "ppt/_rels/presentation.xml.rels",
            _PRESENTATION_RELS_XML.encode("utf-8"),
        )
        zf.writestr("ppt/vbaProject.bin", vba_bytes)
        zf.writestr("customUI/customUI.xml", _CUSTOM_UI_2006_XML.encode("utf-8"))
        zf.writestr("customUI/customUI14.xml", _CUSTOM_UI_2009_XML.encode("utf-8"))

    return target


def build_vbs_launcher(
    vbs_path: Path,
    project_root: Path | None = None,
    python_executable: Path | None = None,
) -> Path:
    """Tworzy bezokienkowy skrypt `launch_azp.vbs` wywoływany przez przycisk na wstążce PowerPointa."""
    target = Path(vbs_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    root_dir = Path(project_root) if project_root is not None else _default_project_root()
    pyw = _resolve_pythonw(python_executable)
    start_script = root_dir / "start_azp.pyw"

    pyw_escaped = str(pyw).replace('"', '""')
    script_escaped = str(start_script).replace('"', '""')
    root_escaped = str(root_dir).replace('"', '""')

    content = (
        'Option Explicit\r\n'
        'Dim sh\r\n'
        'Set sh = CreateObject("WScript.Shell")\r\n'
        f'sh.CurrentDirectory = "{root_escaped}"\r\n'
        f'sh.Run """{pyw_escaped}"" ""{script_escaped}""", 0, False\r\n'
    )
    target.write_text(content, encoding="utf-8")
    return target


def _detect_office_powerpoint_versions() -> list[str]:
    versions: list[str] = []
    try:
        import winreg

        for ver in OFFICE_VERSIONS:
            try:
                with winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    rf"Software\Microsoft\Office\{ver}\PowerPoint",
                ):
                    versions.append(ver)
            except FileNotFoundError:
                continue
    except Exception:
        pass
    return versions or ["16.0"]


def install_powerpoint_addin(
    addin_dir: Path | None = None,
    appdata_dir: Path | None = None,
    project_root: Path | None = None,
    python_executable: Path | None = None,
    register_hkcu: bool = True,
) -> Path:
    """Instaluje dodatek `AsystentZotero.ppam` i rejestruje go na wstążce PowerPointa."""
    target_addin_dir = Path(addin_dir) if addin_dir is not None else _default_addin_dir()
    target_appdata_dir = (
        Path(appdata_dir) if appdata_dir is not None else _default_azp_appdata_dir()
    )

    ppam_path = build_ppam_addin(target_addin_dir / ADDIN_FILENAME)
    build_vbs_launcher(
        target_appdata_dir / "launch_azp.vbs",
        project_root=project_root,
        python_executable=python_executable,
    )
    build_vbs_launcher(
        target_addin_dir / "launch_azp.vbs",
        project_root=project_root,
        python_executable=python_executable,
    )

    if register_hkcu:
        import winreg

        for ver in _detect_office_powerpoint_versions():
            addin_reg_path = (
                rf"Software\Microsoft\Office\{ver}\PowerPoint\AddIns\{ADDIN_NAME}"
            )
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, addin_reg_path) as key:
                winreg.SetValueEx(key, "Path", 0, winreg.REG_SZ, str(ppam_path))
                winreg.SetValueEx(key, "AutoLoad", 0, winreg.REG_DWORD, 1)

            trusted_reg_path = (
                rf"Software\Microsoft\Office\{ver}\PowerPoint\Security"
                rf"\Trusted Locations\{TRUSTED_LOCATION_KEY_NAME}"
            )
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, trusted_reg_path) as tkey:
                folder_str = str(target_addin_dir.resolve())
                if not folder_str.endswith("\\"):
                    folder_str += "\\"
                winreg.SetValueEx(tkey, "Path", 0, winreg.REG_SZ, folder_str)
                winreg.SetValueEx(tkey, "AllowSubfolders", 0, winreg.REG_DWORD, 1)
                winreg.SetValueEx(
                    tkey,
                    "Description",
                    0,
                    winreg.REG_SZ,
                    "Asystent Zotero dla PowerPointa",
                )

    return ppam_path


def uninstall_powerpoint_addin(
    addin_dir: Path | None = None,
    appdata_dir: Path | None = None,
    unregister_hkcu: bool = True,
) -> None:
    """Usuwa pliki dodatku `AsystentZotero.ppam` oraz jego wpisy z rejestru `HKCU`."""
    target_addin_dir = Path(addin_dir) if addin_dir is not None else _default_addin_dir()
    target_appdata_dir = (
        Path(appdata_dir) if appdata_dir is not None else _default_azp_appdata_dir()
    )

    if unregister_hkcu:
        try:
            import winreg

            for ver in OFFICE_VERSIONS:
                for subkey in (
                    rf"Software\Microsoft\Office\{ver}\PowerPoint\AddIns\{ADDIN_NAME}",
                    rf"Software\Microsoft\Office\{ver}\PowerPoint\Security\Trusted Locations\{TRUSTED_LOCATION_KEY_NAME}",
                ):
                    try:
                        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, subkey)
                    except FileNotFoundError:
                        pass
        except Exception:
            pass

    for file_path in (
        target_addin_dir / ADDIN_FILENAME,
        target_addin_dir / "launch_azp.vbs",
        target_appdata_dir / "launch_azp.vbs",
    ):
        try:
            if file_path.exists():
                file_path.unlink()
        except OSError:
            pass


def is_powerpoint_addin_installed(
    addin_dir: Path | None = None,
    check_hkcu: bool = True,
) -> bool:
    """Zwraca `True`, jeśli plik `AsystentZotero.ppam` istnieje i jest zarejestrowany w PowerPoincie."""
    target_addin_dir = Path(addin_dir) if addin_dir is not None else _default_addin_dir()
    ppam_path = target_addin_dir / ADDIN_FILENAME
    if not ppam_path.is_file():
        return False
    if not check_hkcu:
        return True

    try:
        import winreg

        for ver in _detect_office_powerpoint_versions():
            addin_reg_path = (
                rf"Software\Microsoft\Office\{ver}\PowerPoint\AddIns\{ADDIN_NAME}"
            )
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, addin_reg_path) as key:
                    reg_path, _ = winreg.QueryValueEx(key, "Path")
                    auto_load, _ = winreg.QueryValueEx(key, "AutoLoad")
                    if str(reg_path).strip() and int(auto_load) == 1:
                        return True
            except FileNotFoundError:
                continue
    except Exception:
        return False
    return False


__all__ = [
    "ADDIN_FILENAME",
    "ADDIN_NAME",
    "TRUSTED_LOCATION_KEY_NAME",
    "build_ppam_addin",
    "build_vbs_launcher",
    "install_powerpoint_addin",
    "is_powerpoint_addin_installed",
    "uninstall_powerpoint_addin",
]
