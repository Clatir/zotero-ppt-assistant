# Postęp

## Stan
- **Zakończony etap:** A1 Środowisko (przegląd `medium`: brak uwag critical/high/medium).
- **W trakcie:** A2 Kontrakty i atrapy (ciąg A).
- **Następny krok po A2:** Brama A→B i pauza przed ciągiem B.

## Kolejka
| Ciąg | Stan |
|---|---|
| A – Fundament | w trakcie (A1 gotowe, trwa A2) |
| B – Rdzeń | – |
| C – Integracja i odbiór | – |

## Zrealizowane etapy

### A1 Środowisko (2026-11-27)
- **Zakres:** przeniesienie repozytorium do `c:\Agrvt-Projekty\PP`, `.gitignore`, `requirements.txt` (`pywin32`, `requests`, `pytest`), `pytest.ini`, szkielet pakietów `azp/*`, testy dymne `tests/test_smoke_env.py` (lokalne API Zotero ze stylami v4 i APA oraz cykl życia prezentacji i `Tags` przez COM PowerPointa).
- **Decyzje:** T7.
- **Testy:** `pytest -q` (1 passed w 0,79 s), `pytest -q -m com` (1 passed w 8,76 s).
- **Czas faktyczny:** ~5 min (szacunek: 10 min).

## Uwagi na później
- `low`/`nit` (A1, `tests/test_smoke_env.py`): plik `tests/data/test_smoke.pptx` jest usuwany przed startem testu i zostaje po teście w ignorowanym katalogu `tests/data/` do ewentualnego podglądu.
