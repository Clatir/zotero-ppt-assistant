# Postęp

## Stan
- **Zakończony ciąg:** Ciąg A – Fundament (A1, A2).
- **W trakcie:** Ciąg B – Rdzeń:
  - Tor B1: **B1.1 Klient Zotero** ukończony i zatwierdzony; trwa **B1.2 Okienko**.
  - Tor B2: **B2.1 Warstwa PowerPoint** w przeglądzie kodu (`xhigh`); następnie **B2.2 Silnik numeracji i stylów**.

## Kolejka
| Ciąg | Stan |
|---|---|
| A – Fundament | gotowe (A1, A2) |
| B – Rdzeń (tory B1 i B2 równolegle) | w trakcie (B1.1 gotowe; trwa B1.2 oraz B2.1→B2.2) |
| C – Integracja i odbiór | – |

## Zrealizowane etapy

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
