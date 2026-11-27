# Decyzje użytkownika (D)

| ID | Data | Decyzja |
|---|---|---|
| D1 | 2026-11-27 | Forma narzędzia: pływające mini-okienko obok PowerPointa (opcja 1), nie skrypt wsadowy. |
| D2 | 2026-11-27 | Styl domyślny: własny styl Zotero „IEEE Praca Inzynierska Access Date v4” (`http://www.zotero.org/styles/ieee-praca-inzynierska-access-date-v4`). |
| D3 | 2026-11-27 | Drugi dostępny styl: APA 7 (`http://www.zotero.org/styles/apa`). Styl wybierany w okienku. |
| D4 | 2026-11-27 | Numeracja IEEE: klasyczna, globalna (pierwsze wystąpienie); w okienku wybór trybu (globalna / od [1] na każdym slajdzie). |
| D5 | 2026-11-27 | Czcionka stopek wybierana w okienku z listy czcionek zainstalowanych w systemie, z rozmiarem. |
| D6 | 2026-11-27 | Bez slajdu „Bibliografia” na końcu. |
| D7 | 2026-11-27 | Prace pogrupowane w ciągi, części niezależne równolegle. |
| D8 | 2026-11-27 | Zasady pracy według `Clatir/RegulatedDay` (Zasady współpracy, skill `etap`), jak najdokładniej. |
| D9 | 2026-11-27 | Położenie pola źródeł wybierane w okienku: (a) dół slajdu, cała szerokość, nad stopką PowerPointa; (b) lewy dolny róg, węższe; (c) w miejscu stopki PowerPointa. Do wyboru też: z nagłówkiem „Źródła:” albo same źródła bez nagłówka. |
| D10 | 2026-11-27 | „Cytuj” bez kursora w polu tekstowym: nic nie zmieniać, tylko poprosić o ustawienie kursora. |
| D11 | 2026-11-27 | Kod tylko lokalnie (git na tym komputerze), bez GitHuba. *(zastąpione przez D12)* |
| D12 | 2026-11-27 | Publiczne repozytorium na GitHubie (`Clatir/zotero-ppt-assistant`), commity wypychane na bieżąco; bez danych użytkownika i osobistych treści (zastępuje D11). |
| D13 | 2026-11-27 | Cytowania zbiorcze (np. `[1, 2]` lub `(Kowalski, 2023; Nowak, 2024)`): obsługiwany zarówno wielokrotny wybór na liście wyników w okienku (Ctrl/Shift), jak i automatyczne scalanie, gdy kursor przy wstawianiu stoi tuż przy istniejącym cytowaniu. |
| D14 | 2026-11-27 | Zmiana stylu, trybu numeracji, czcionki, rozmiaru lub położenia stopki na liście wyboru w okienku od razu przebudowuje całą prezentację (po automatycznym utworzeniu kopii zapasowej `.bak.pptx`). |
| D15 | 2026-11-27 | Skróty klawiszowe (C2): skróty wewnątrz okienka (`Enter` / `Ctrl+Enter` = *Cytuj* na liście wyników, `F5` / `Ctrl+R` = *Odśwież / Przenumeruj*, `Ctrl+F` = fokus na polu *Szukaj*), bez globalnego przechwytywania klawiatury w Windows. |
| D16 | 2026-11-27 | Uruchamianie ikoną (C2): skrypt startowy `start_azp.bat` / `start_azp.pyw` (uruchamiający okienko bez czarnej konsoli w tle), bez pakowania do `.exe`. |
| D17 | 2026-11-27 | Kopia zapasowa przed przebudową przez „Zapisz kopię jako” (`SaveCopyAs`) do `<nazwa>.bak.pptx` obok pliku; plik użytkownika na dysku nie jest zapisywany bez jego wiedzy. |
| D18 | 2026-11-27 | Niezapisana prezentacja (bez pliku): przebudowa (zmiana ustawień, *Odśwież / Przenumeruj*) nie jest wykonywana – komunikat „Zapisz prezentację, żeby można było zrobić kopię zapasową”, ustawienie wraca do poprzedniego; *Cytuj* działa normalnie. |
| D19 | 2026-11-27 | Okienko śledzi aktywną prezentację: po przełączeniu wczytuje jej ustawienia (styl, numeracja, czcionka, położenie, nagłówek); prezentacja bez ustawień AZP dostaje ustawienia aktualnie widoczne w okienku. |
| D20 | 2026-11-27 | Po *Cytuj* fokus wraca do PowerPointa, kursor stoi za wstawionym cytowaniem. |
| D21 | 2026-11-27 | Tryb ciemny okienka: przełącznik „Tryb ciemny” w okienku (zapisywany w `config.json`, bez wywoływania przebudowy prezentacji), zmieniający kolorystykę wszystkich kontrolek oraz paska tytułowego Windows. |
| D22 | 2026-11-27 | Uruchamianie z poziomu wstążki PowerPointa: natywny dodatek `AsystentZotero.ppam` instalowany w `%APPDATA%\Microsoft\AddIns` i rejestrowany w `HKCU\Software\Microsoft\Office\<ver>\PowerPoint\AddIns\AsystentZotero` (`AutoLoad=1`), dodający przycisk **Asystent Zotero** na karcie **Dodatki** oraz **Narzędzia główne** na wstążce PowerPointa (`install_pp_addin.bat` / `python -m azp --install-addin`). |
| D23 | 2026-11-27 | Instalacja i aktualizacja jednym kliknięciem: skrypt `instaluj_lub_aktualizuj.bat` (`python -m azp --update`) oraz instrukcja `README.md` – pobiera najnowszą wersję z GitHuba (`git pull` lub archiwum ZIP, gdy projekt pobrano bez Gita), instaluje pakiety z `requirements.txt` i odświeża dodatek wstążki PowerPointa z zachowaniem ustawień użytkownika w `config.json`. |
