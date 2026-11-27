# Asystent Zotero dla PowerPointa (AZP)

Pływające okienko dla systemu Windows do wstawiania cytowań z **Zotero** bezpośrednio do prezentacji **Microsoft PowerPoint**. Narzędzie obsługuje style cytowania **IEEE v4** oraz **APA 7**, automatycznie utrzymuje źródła w stopce pod slajdem, przenumerowuje cytowania w całej prezentacji, udostępnia przycisk **Asystent Zotero** bezpośrednio na wstążce PowerPointa oraz wspiera tryb ciemny.

## Wymagania

- System **Windows** z programem **Microsoft PowerPoint**
- **Zotero** (np. Zotero 7 / 10) z włączonym lokalnym API:
  *Edycja → Ustawienia → Zaawansowane → Zezwalaj innym aplikacjom na tym komputerze na komunikację z Zotero*
- **Python 3.12+** (podczas instalacji zaznacz opcję **Add python.exe to PATH**)

## Instalacja na nowym komputerze

1. Pobierz repozytorium przez `git clone https://github.com/Clatir/zotero-ppt-assistant.git` lub wybierz **Code → Download ZIP** na GitHubie i rozpakuj archiwum do stałego folderu.
2. Kliknij dwukrotnie plik **`instaluj_lub_aktualizuj.bat`** (zainstaluje wymagane pakiety Pythona i dodatek wstążki PowerPointa).
3. Otwórz PowerPointa i kliknij **Asystent Zotero** na wstążce (karta *Narzędzia główne* lub *Dodatki*).

## Aktualizacja do najnowszej wersji

Kliknij dwukrotnie **`instaluj_lub_aktualizuj.bat`** – skrypt sam pobierze najnowszą wersję z GitHuba przez `git pull` lub ZIP, zaktualizuje pakiety oraz odświeży dodatek wstążki PowerPointa, zachowując Twoje ustawienia w `config.json`.

## Uruchamianie ręczne i komendy

- `start_azp.bat` – uruchamia okienko AZP bez czarnego okna konsoli
- `install_pp_addin.bat` – instaluje lub odświeża dodatek wstążki PowerPointa
- `python -m azp --update` – aktualizuje kod z GitHuba i odświeża dodatek PowerPointa
- `python -m azp --uninstall-addin` – odinstalowuje dodatek wstążki PowerPointa
