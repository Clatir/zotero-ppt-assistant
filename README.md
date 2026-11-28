# Asystent Zotero dla PowerPointa (AZP)

Pływające okienko dla systemu Windows do wstawiania cytowań z **Zotero** bezpośrednio do prezentacji **Microsoft PowerPoint**. Narzędzie obsługuje style cytowania **IEEE v4** oraz **APA 7**, wybór kolekcji Zotero, przypisy w stopce pod slajdem (także bez ustawiania kursora w tekście i na slajdach ze samymi zdjęciami), wielokolumnową bibliografię bez nagłówka z wyborem czcionki, rozmiaru i trybu dopasowania, przycisk **Asystent Zotero** na wstążce PowerPointa oraz tryb ciemny.

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

Program **aktualizuje się automatycznie w tle** po otwarciu okienka – pobiera nową wersję z GitHuba przez Git lub ZIP (zachowując Twoje ustawienia w `config.json`), odświeża dodatek PowerPointa i automatycznie przeładowuje okienko.

Ręcznie możesz też w każdej chwili kliknąć dwukrotnie **`instaluj_lub_aktualizuj.bat`**, aby wymusić aktualizację kodu, pakietów Pythona oraz dodatku wstążki.

## Uruchamianie ręczne i komendy

- `start_azp.bat` – uruchamia okienko AZP bez czarnego okna konsoli
- `install_pp_addin.bat` – instaluje lub odświeża dodatek wstążki PowerPointa
- `python -m azp --update` – aktualizuje kod z GitHuba i odświeża dodatek PowerPointa
- `python -m azp --uninstall-addin` – odinstalowuje dodatek wstążki PowerPointa
