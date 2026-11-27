# Postęp

## Stan
- **Zakończone ciągi:**
  - **Ciąg A – Fundament** (A1, A2).
  - **Ciąg B – Rdzeń** (Tor B1: **B1.1 Klient Zotero**, **B1.2 Okienko**; Tor B2: **B2.1 Warstwa PowerPoint**, **B2.2 Silnik numeracji i stylów** – wszystkie 4 etapy po przeglądach kodu `high`/`xhigh`, commity wypchnięte na GitHuba).
- **Pełny zestaw testów po C1:** `107 passed` bez COM + `3 passed` COM (`-m com`).
- **Przed ciągiem C:** pytania rozstrzygnięte i zapisane jako D15–D20.
- **Następny krok:** **Etap C2 Dodatki** (w ramach Ciągu C), a następnie pauza przed `C3 Odbiór na Twojej prezentacji`.

## Kolejka
| Ciąg | Stan |
|---|---|
| A – Fundament | gotowe (A1, A2) |
| B – Rdzeń (tory B1 i B2 równolegle) | gotowe (B1.1, B1.2, B2.1, B2.2) |
| C – Integracja i odbiór | w toku: C1 gotowe → C2 Dodatki → pauza przed C3 |

## Zrealizowane etapy

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
- `low`/`nit` (A2, `azp/contracts.py:183`): w Pythonie `bool` dziedziczy po `int`, więc można opcjonalnie odrzucać `isinstance(font_size_pt, bool)` w walidacji `Settings`.
- `low`/`nit` (B1.1, `tests/test_zotero.py:517`): test integracyjny z żywym Zotero zakłada, że pierwsza pozycja z `/items/top` posiada element formatowany kursywą w stylu IEEE v4.
- `low`/`nit` (C1, `azp/gui/app.py:500`, `azp/ppt/com_doc.py:726`): wywołanie `save_settings` przy `rebuild=False` oraz odczyt `pres.Slides.Count` na początku `_resolve_cursor_context()` można dodatkowo opakować w `try...except Exception:`.
- `low`/`nit` (C1, `azp/ppt/com_doc.py:571`): `deserialize_settings` może dodatkowo łapać `(ValueError, TypeError)` na wypadek ręcznie uszkodzonego tagu `AZP_SETTINGS`.
