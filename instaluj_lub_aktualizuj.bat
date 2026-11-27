@echo off
chcp 65001 >nul
cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    echo [BŁĄD] Nie znaleziono programu Python w systemie.
    echo Zainstaluj Python 3.12 lub nowszy ze strony https://www.python.org/downloads/
    echo Podczas instalacji zaznacz opcję "Add python.exe to PATH".
    pause
    exit /b 1
)

echo [1/4] Sprawdzanie i instalacja wymaganych bibliotek...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [BŁĄD] Nie udało się zainstalować wymaganych bibliotek z requirements.txt.
    pause
    exit /b 1
)

echo [2/4] Pobieranie aktualizacji kodu...
python -m azp --update
if errorlevel 1 (
    echo [UWAGA] Aktualizacja kodu z GitHuba nie powiodła się; użyto bieżących plików.
)

echo [3/4] Weryfikacja zależności po aktualizacji...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [BŁĄD] Nie udało się zaktualizować bibliotek po pobraniu nowej wersji.
    pause
    exit /b 1
)

echo [4/4] Odświeżanie dodatku wstążki PowerPointa...
python -m azp --install-addin
if errorlevel 1 (
    echo [BŁĄD] Nie udało się zainstalować lub odświeżyć dodatku wstążki PowerPointa.
    pause
    exit /b 1
)

echo.
echo Gotowe! Uruchom lub zrestartuj PowerPointa – przycisk Asystent Zotero znajduje się na karcie Narzędzia główne oraz Dodatki.
pause
