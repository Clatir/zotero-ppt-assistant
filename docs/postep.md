# Postęp

## Stan
- **Zakończony ciąg:** Ciąg A – Fundament (etapy A1 i A2 ukończone, przegląd kodu zatwierdzony, commity na GitHubie).
- **Brama A→B:** kontrakty (`azp/contracts.py`) i atrapy (`tests/fakes/`) gotowe i przetestowane (`12 passed` bez COM + `1 passed` COM).
- **Przed ciągiem B:** pytania rozstrzygnięte i zapisane jako D13–D14.
- **Następny krok:** Ciąg B (dwa tory równolegle: B1.1→B1.2 oraz B2.1→B2.2) po słowie „dalej”.

## Kolejka
| Ciąg | Stan |
|---|---|
| A – Fundament | gotowe (A1, A2) |
| B – Rdzeń (tory B1 i B2 równolegle) | czeka na „dalej” (brama A→B) |
| C – Integracja i odbiór | – |

## Zrealizowane etapy

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
