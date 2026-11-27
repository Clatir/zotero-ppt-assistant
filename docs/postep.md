# Postęp

## Stan
- **Zakończone ciągi / etapy:**
  - **Ciąg A – Fundament** (A1, A2).
  - **Ciąg B – Rdzeń** (Tor B1: **B1.1 Klient Zotero**, **B1.2 Okienko**; Tor B2: **B2.1 Warstwa PowerPoint**, **B2.2 Silnik numeracji i stylów**).
  - **Ciąg C – Integracja i odbiór:** etapy kodowe **C1 Integracja** (`xhigh`), **C2 Dodatki** (`high`) oraz poprawki audytowe + **Tryb ciemny** (D21) + **Dodatek wstążki PowerPointa `.ppam`** (D22, T11, T12) + **Instalator i auto-aktualizacja w tle** (D23, D24, T13) zakończone, po przeglądach kodu, commity wypchnięte na GitHuba.
- **Pełny zestaw testów:** `155 passed` bez COM + `3 passed` COM (`-m com`).
- **Następny krok:** **Etap C3 Odbiór na Twojej prezentacji** (wspólny odbiór na żywo według listy w `docs/odbior.md`).

## Kolejka
| Ciąg | Stan |
|---|---|
| A – Fundament | gotowe (A1, A2) |
| B – Rdzeń (tory B1 i B2 równolegle) | gotowe (B1.1, B1.2, B2.1, B2.2) |
| C – Integracja i odbiór | C1, C2 + poprawki audytu, tryb ciemny (D21), wstążka PP (D22) i auto-aktualizacja (D23, D24) gotowe; gotowe do C3 (Odbiór na Twojej prezentacji) |

## Zrealizowane etapy

### Audyt kodu, Tryb ciemny, Dodatek wstążki PowerPointa i Auto-aktualizacja (2026-11-27)
- **Zakres:**
  - **Naprawa błędów „działających tylko teoretycznie” (T11):**
    1. *Martwe cytowania („duchy”) po ręcznym usunięciu w PowerPoincie (`azp/ppt/com_doc.py`)*: weryfikacja tagów `AZP_CITATIONS` względem aktualnego tekstu kształtów slajdu w `_get_citations_by_shape()` i `read_slides()`, sortowanie cytowań po faktycznej pozycji w tekście (obsługa przeklejania zdań) oraz usunięcie błędnego doklejania nieznalezionych cytowań z usuniętych pól do `shape_cits[0]`.
    2. *Utrata kursora i wykrywanie zaznaczenia w prawdziwym oknie PowerPointa (`azp/ppt/com_doc.py`)*: ponowne zaznaczanie pozycji kursora (`_restore_last_insert_cursor`) po utworzeniu/modyfikacji kształtu stopki w `set_slide_footer()` oraz po `focus_presentation_window()`; obsługa kursora wewnątrz zgrupowanych kształtów (`GroupItems` / `HasChildShapeRange`) i komórek tabeli (`HasTable`); odfiltrowanie panelu notatek prelegenta (`NotesPage`); obsługa zaznaczonego tekstu (`Length > 0`) oraz odzyskiwanie kursora po przełączeniu zaznaczenia na kształt (`PP_SELECTION_SHAPES`).
    3. *Formatowanie stopki i odporność tagów (`azp/ppt/com_doc.py`)*: ustawianie `WordWrap` i `Width` po `AutoSize` (zapobiega zmianie szerokości stopki przy autodopasowaniu), wywołania COM `Characters` tylko dla przebiegów z kursywą/pogrubieniem, odporność `deserialize_citations` / `deserialize_footer_entries` / `deserialize_settings` na uszkodzony JSON i błędne typy.
    4. *Spójność `AzpController` i odporność na błędy COM (`azp/gui/app.py`)*: zastąpienie zduplikowanych metod `_cite_builtin` / `_rebuild_builtin` w `AzpController` pełnym `CitationEngine` (scalanie sąsiednich cytowań D13 i globalne przenumerowanie na wszystkich slajdach), przechwytywanie surowych wyjątków COM we wszystkich wywołaniach `slide_doc`, dwuklik `<Double-Button-1>` na liście wyników (= *Cytuj*) oraz strzałka `<Down>` w polu szukania (= przejście do listy wyników).
    5. *Wyszukiwanie wielowyrazowe i fallback stylu w `ZoteroClient` (`azp/zotero/client.py`)*: fallback wyszukiwania po tokenach przy zapytaniach łączących różne pola (autor + tytuł + rok), dociąganie wyników przy pustym zapytaniu ponad odfiltrowane załączniki/notatki oraz automatyczny fallback stylu `STYLE_IEEE_V4` → `http://www.zotero.org/styles/ieee` przy braku niestandardowego stylu v4 w Zotero (HTTP 400).
  - **Tryb ciemny okienka (D21):** pole `Settings.dark_mode: bool = False` w `azp/contracts.py`, przełącznik *Tryb ciemny* w `AzpWindow` (`azp/gui/app.py`), palety `DARK_THEME_COLORS` i `LIGHT_THEME_COLORS` ze stylizacją wszystkich widgetów `ttk`/`tk`, rozwijanych list `Combobox` oraz ciemnego paska tytułowego Windows 10/11 (DWM), z zapisem w `config.json` bez wywoływania przebudowy prezentacji.
  - **Dodatek na wstążce PowerPointa (D22, T12):** moduł `azp/ppt/addin.py` (`build_ppam_addin`, `build_vbs_launcher`, `install_powerpoint_addin`, `uninstall_powerpoint_addin`, `is_powerpoint_addin_installed`), skrypt `install_pp_addin.bat`, flagi CLI `--install-addin` / `--uninstall-addin` oraz `focus_existing_azp_window()` w `azp/main.py`. Tworzy i rejestruje w `HKCU` dodatek `AsystentZotero.ppam` z przyciskiem **Asystent Zotero** na kartach **Dodatki** (`TabAddIns`) oraz **Narzędzia główne** (`TabHome`) na wstążce PowerPointa.
  - **Instalator jednym kliknięciem i auto-aktualizacja w tle (D23, D24, T13):** moduł `azp/updater.py` (`fetch_remote_main_sha`, `auto_update_if_needed`, `update_from_git`, `extract_github_zip_bytes`, `update_from_github_zip`, `update_project_code`), wątek auto-aktualizacji w tle `start_background_auto_update` i `reload_window_after_update` w `azp/main.py` (z flagami `--update` oraz `--no-auto-update`), zbiorczy skrypt `instaluj_lub_aktualizuj.bat` oraz polska instrukcja `README.md` (po przeglądach `high`).
- **Decyzje:** D21–D24, T11–T13.
- **Testy:** `pytest -q` (`155 passed`).

### C2 Dodatki (2026-11-27)
- **Zakres:** `azp/config.py` (`DEFAULT_CONFIG_PATH`, `load_config`, `save_config` – zapis i odczyt domyślnych ustawień w `config.json` z bezpiecznym powrotem do `Settings()` przy braku pliku lub uszkodzonym JSON-ie), `.gitignore` (`config.json`), `azp/contracts.py` (ścisła walidacja typów w `Settings.__post_init__` i `Settings.from_dict` – M1), `azp/gui/app.py` i `azp/main.py` (integracja `config_path`, skróty klawiszowe wewnątrz okienka D15: `Enter` / `Ctrl+Enter` = *Cytuj*, `F5` / `Ctrl+R` = *Odśwież / Przenumeruj*, `Ctrl+F` = fokus i zaznaczenie pola *Szukaj*), `start_azp.pyw` i `start_azp.bat` (uruchamianie ikoną bez czarnej konsoli w tle D16), `tests/test_c2_extras.py` (26 przypadków testowych, w tym poprawki H1 i M1 z 1. rundy przeglądu `high`).
- **Decyzje:** realizacja D15, D16 (oraz D14, D18, D19).
- **Testy:** `pytest -q tests/test_c2_extras.py` (`26 passed`), pełny zestaw `pytest -q` (`133 passed`).
- **Czas faktyczny:** ~12 min (szacunek: 20–30 min).

### C1 Integracja (2026-11-27)
- **Zakres:** `azp/main.py`, `azp/__main__.py` (`build_window`, `describe_error`, `main` – połączenie `AzpWindow`, `CitationEngine`, `ZoteroClient` i `PowerPointSlideDoc`, powrót fokusu do PowerPointa po *Cytuj* D20, konwersja nieoczekiwanych błędów COM na komunikaty paska stanu), `azp/ppt/com_doc.py` (`TextEdit`, `plan_citation_insert`, `plan_citation_replacements`, `_apply_text_edit_to_text_range` – edycje wyłącznie na podzakresach `Characters(start, len)` bez utraty formatowania slajdu T10, korekta pozycji kursora po `update_slide_citations` H2/M2, kopia zapasowa przez `SaveCopyAs` bez zapisu oryginału D17, `get_presentation_key` D19, `choose_powerpoint_window` i `focus_presentation_window` z odfiltrowaniem własnego okna AZP D20/H1), `azp/engine/service.py` i `azp/gui/app.py` (blokada przebudowy niezapisanej prezentacji z komunikatem `MSG_SAVE_BEFORE_REBUILD` i wycofaniem wyboru D18, śledzenie aktywnej prezentacji `sync_with_active_presentation` / `check_active_presentation` D19/M1), `tests/conftest.py`, `tests/fakes/fake_com.py`, `tests/test_main.py` oraz nowe testy w `tests/test_ppt.py`, `tests/test_engine.py`, `tests/test_gui.py` (łącznie 23 nowe testy jednostkowe + kompleksowy test końcowy COM `@pytest.mark.com` okienko + żywe Zotero 10 + żywy PowerPoint, po 3 rundach przeglądu `xhigh`).
- **Decyzje:** D17–D20, T10.
- **Testy:** `pytest -q` (`107 passed`), `pytest -q -m com` (`3 passed`).
- **Czas faktyczny:** ~30 min (szacunek: 30–45 min).

### B2.2 Silnik numeracji i stylów (2026-11-27)
- **Zakres:** `azp/engine/numbering.py`, `azp/engine/service.py`, `azp/engine/__init__.py` (czyste funkcje `is_numeric_style`, `assign_global_numbers`, `assign_slide_numbers`, `merge_citation_keys`, `format_citation_text`, `build_slide_footer_entries`, `compute_presentation_plan` oraz koordynator `CitationEngine`: numeracja IEEE `global` i `per_slide` D2/D4, ten sam numer dla powtórzeń, scalanie `[1, 2]` i `(Autor1, Rok1; Autor2, Rok2)` D13 wraz z integracją `get_adjacent_citation_at_cursor()`, sortowanie alfabetyczne stopki w APA D3, blokada bez kursora D10, natychmiastowa przebudowa z kopią zapasową `.bak.pptx` po zmianie ustawień D14, adaptery callbacków `make_cite_callback` i `make_rebuild_callback`), `tests/test_engine.py` (20 testów jednostkowych i integracyjnych, w tym 4 poprawki H1, M1–M3 z 1. rundy przeglądu `xhigh`).
- **Decyzje:** realizacja D2–D6, D9, D10, D13, D14, T4, T8, T9.
- **Testy:** `pytest -q tests/test_engine.py` (`20 passed in 0.07s`), pełny zestaw `pytest -q` (`84 passed`) + `pytest -q -m com` (`2 passed`).
- **Czas faktyczny:** ~15 min (szacunek: 1–1,5 h).

### B1.2 Okienko (2026-11-27)
- **Zakres:** `azp/gui/app.py`, `azp/gui/__init__.py` (`AzpController`, `AzpWindow` – pływające okienko tkinter po polsku: wyszukiwarka z obsługą `<Return>`, lista wyników z wielokrotnym wyborem Ctrl/Shift D13, przyciski *Cytuj* i *Odśwież / Przenumeruj*, listy wyboru stylu D2/D3, trybu numeracji D4, czcionki z systemu `tkinter.font.families()` i rozmiaru D5, 3 położeń pola źródeł D9, nagłówka „Źródła:” tak/nie D9, natychmiastowa przebudowa z kopią zapasową `.bak.pptx` po zmianie ustawień D14 wraz z wycofaniem wyboru przy błędzie H1, blokada *Cytuj* bez kursora D10, przełącznik „Zawsze na wierzchu”, obsługa `ZoteroOfflineError`), `tests/test_gui.py` (16 testów z atrapami `FakeZoteroSource` i `FakeSlideDoc`, w tym 2 poprawki H1 i M1 z 1. rundy przeglądu `high`).
- **Decyzje:** realizacja D1–D5, D9, D10, D13, D14.
- **Testy:** `pytest -q tests/test_gui.py` (`16 passed in 1.93s`).
- **Czas faktyczny:** ~12 min (szacunek: 45–60 min).

### B2.1 Warstwa PowerPoint (2026-11-27)
- **Zakres:** `azp/ppt/com_doc.py`, `azp/ppt/powerpoint.py`, `azp/ppt/__init__.py` (`PowerPointSlideDoc` implementujący `SlideDoc`, wykrywanie aktywnego slajdu i kursora z `ActiveWindow.Selection` oraz w trybie bezokienkowym, blokada D10 bez kursora, wstawianie i automatyczne scalanie sąsiednich cytowań D13, izolacja cytowań per kształt tekstowy T9, pole źródeł w 3 położeniach D9 z nagłówkiem „Źródła:” lub bez, czcionka/rozmiar/kursywa na podzakresach `Characters(start, len)` i autodopasowanie, tagi `AZP_*`, dwuprzebiegowa podmiana cytowań, kopia zapasowa `.bak.pptx`), `tests/test_ppt.py` (11 testów jednostkowych + kompleksowy test COM, w tym 4 poprawki H1, H2, M1, M2 z 1. rundy przeglądu `xhigh`).
- **Decyzje:** T9 (oraz realizacja D5, D9, D10, D13, D14, T3, T4, T7, T8).
- **Testy:** `pytest -q tests/test_ppt.py` (`11 passed in 0.10s`), `pytest -q -m com tests/test_ppt.py` (`1 passed in 15.57s`).
- **Czas faktyczny:** ~20 min (szacunek: 1–1,5 h).

### B1.1 Klient Zotero (2026-11-27)
- **Zakres:** `azp/zotero/html_parser.py` (`parse_csl_bib_html`, `parse_csl_citation_html` – konwersja HTML CSL na `TextSpan` z kursywą i pogrubieniem, dekodowaniem encji HTML, obsługą znaczników void `<br>` oraz usuwaniem wiodącego `[n]` w stylach numerycznych zgodnie z T4 i T8), `azp/zotero/client.py` (`ZoteroClient` implementujący `ZoteroSource`: `is_available`, `search_items` po autorze/tytule/roku, `get_formatted_ref`, `get_formatted_refs`, `list_styles`, obsługa `ZoteroOfflineError` z komunikatem „Zotero nie odpowiada”), `tests/test_zotero.py` (25 testów jednostkowych i integracyjnych z żywym Zotero 10, w tym 3 poprawki M1–M3 z 1. rundy przeglądu `high`).
- **Decyzje:** realizacja D2, D3, T2, T4, T8.
- **Testy:** `pytest -q tests/test_zotero.py` (`25 passed in 0.90s`).
- **Czas faktyczny:** ~10 min (szacunek: 30–45 min).

### A2 Kontrakty i atrapy (2026-11-27)
- **Zakres:** `azp/contracts.py` (`TextSpan`, `StyleInfo`, `ZoteroItem`, `FormattedRef`, `Citation`, `FooterEntry`, `SlideRecord`, `Settings`, wyjątki domenowe, protokoły `@runtime_checkable` `ZoteroSource` i `SlideDoc`), atrapy `tests/fakes/fake_zotero.py` (`FakeZoteroSource`) i `tests/fakes/fake_ppt.py` (`FakeSlideDoc`), zestaw 11 testów kontraktów w `tests/test_contracts.py` (w tym poprawka z 1. rundy przeglądu na dwuprzebiegową podmianę tekstów cytowań przy duplikatach i zamianie miejscami).
- **Decyzje:** T8 (oraz realizacja D2–D5, D9, D10 w kontraktach).
- **Testy:** `pytest -q` (`12 passed in 0.65s`).
- **Czas faktyczny:** ~6 min (szacunek: 15 min).

### A1 Środowisko (2026-11-27)
- **Zakres:** przeniesienie repozytorium do `c:\Agrvt-Projekty\PP`, `.gitignore`, `requirements.txt` (`pywin32`, `requests`, `pytest`), `pytest.ini`, szkielet pakietów `azp/*`, testy dymne `tests/test_smoke_env.py` (lokalne API Zotero ze stylami v4 i APA oraz cykl życia prezentacji i `Tags` przez COM PowerPointa).
- **Decyzje:** T7.
- **Testy:** `pytest -q` (1 passed w 0,79 s), `pytest -q -m com` (1 passed w 8,76 s).
- **Czas faktyczny:** ~5 min (szacunek: 10 min).

## Uwagi na później
- `low`/`nit` (A1, `tests/test_smoke_env.py`): plik `tests/data/test_smoke.pptx` jest usuwany przed startem testu i zostaje po teście w ignorowanym katalogu `tests/data/` do ewentualnego podglądu.
- `low`/`nit` (B1.1, `tests/test_zotero.py`): test integracyjny z żywym Zotero zakłada, że pierwsza pozycja z `/items/top` posiada element formatowany kursywą w stylu IEEE v4.
