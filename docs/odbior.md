# Odbiór na Windows

Każdy punkt: **Automatycznie:** (które testy to sprawdzają) albo **Tylko Ty:**, oraz Jak sprawdzić / Oczekiwany efekt /
Nie przechodzi, gdy.

## Do odebrania
- **A1.1 Odpowiedź lokalnego API Zotero dla stylów IEEE v4 i APA** – **Automatycznie:** `tests/test_smoke_env.py::test_zotero_local_api_responds_ieee_v4_and_apa`
  - *Jak sprawdzić:* przy włączonym Zotero 10 uruchom `python -m pytest -q tests/test_smoke_env.py`.
  - *Oczekiwany efekt:* test przechodzi (`1 passed`); Zotero zwraca `bib` i `citation` dla stylu `ieee-praca-inzynierska-access-date-v4` (z `[1]`) oraz `apa`.
  - *Nie przechodzi, gdy:* Zotero jest wyłączone, lokalne API zwraca 403 albo brak stylu v4 w Zotero.
- **A1.2 Sterowanie PowerPointem przez COM i trwałość Tags** – **Automatycznie:** `tests/test_smoke_env.py::test_powerpoint_com_open_tag_save_and_close`
  - *Jak sprawdzić:* przy zamkniętym PowerPoincie uruchom `python -m pytest -q -m com tests/test_smoke_env.py`.
  - *Oczekiwany efekt:* test przechodzi (`1 passed`); tworzy, zapisuje, otwiera ponownie i zamyka `tests/data/test_smoke.pptx` z zachowaniem tekstu oraz znaczników `Tags`.
  - *Nie przechodzi, gdy:* PowerPoint jest otwarty przed testem albo `pywin32` nie może utworzyć instancji `PowerPoint.Application`.
- **A2.1 Zgodność kontraktów `ZoteroSource` i `SlideDoc` oraz atrap testowych** – **Automatycznie:** `tests/test_contracts.py` (11 testów)
  - *Jak sprawdzić:* uruchom `python -m pytest -q tests/test_contracts.py`.
  - *Oczekiwany efekt:* 11 testów przechodzi (`11 passed`); walidacja `Settings` (D2–D5, D9), blokada cytowania bez kursora (D10), zachowanie kursywy w `TextSpan`/`FooterEntry` (T4, T8) oraz dwuprzebiegowa podmiana numerów cytowań działają poprawnie.
  - *Nie przechodzi, gdy:* którykolwiek z 11 testów kontraktów zgłasza błąd.
- **B1.1 Klient lokalnego API Zotero i parser HTML CSL** – **Automatycznie:** `tests/test_zotero.py` (25 testów)
  - *Jak sprawdzić:* przy włączonym Zotero 10 uruchom `python -m pytest -q tests/test_zotero.py`.
  - *Oczekiwany efekt:* 25 testów przechodzi (`25 passed`); wyszukiwanie po autorze/tytule/roku, formatowanie cytowań i opisów dla IEEE v4 oraz APA 7, konwersja HTML CSL na przebiegi `TextSpan` z kursywą bez wiodącego `[1]` (T4, T8) oraz komunikat „Zotero nie odpowiada” przy wyłączonym API działają poprawnie.
  - *Nie przechodzi, gdy:* którykolwiek z 25 testów w `tests/test_zotero.py` zgłasza błąd.
- **B2.1 Warstwa COM PowerPoint (`PowerPointSlideDoc`)** – **Automatycznie:** `tests/test_ppt.py` (11 testów jednostkowych + 1 test COM `@pytest.mark.com`)
  - *Jak sprawdzić:* przy zamkniętym PowerPoincie uruchom `python -m pytest -q tests/test_ppt.py` oraz `python -m pytest -q -m com tests/test_ppt.py`.
  - *Oczekiwany efekt:* przechodzą wszystkie testy jednostkowe (`11 passed`) oraz test COM (`1 passed`); sprawdzane są: blokada wstawienia cytowania bez kursora (D10), wstawienie i automatyczne scalanie sąsiednich cytowań (D13), izolacja cytowań przy wielu polach tekstowych na slajdzie (T9), tworzenie pola źródeł w 3 położeniach (`bottom_full`, `bottom_left`, `replace_footer` – D9) z nagłówkiem „Źródła:” i bez, czcionka, rozmiar i kursywa na podzakresach `Characters(start, len)` (D5, T4), trwałość tagów po ponownym otwarciu pliku, dwuprzebiegowa podmiana cytowań oraz tworzenie kopii zapasowej `.bak.pptx`.
  - *Nie przechodzi, gdy:* PowerPoint użytkownika jest otwarty przed testem COM lub którykolwiek z testów w `tests/test_ppt.py` zgłasza błąd.
- **B1.2 Okienko GUI (`AzpWindow` / `AzpController`)** – **Automatycznie:** `tests/test_gui.py` (16 testów z atrapami `FakeZoteroSource` i `FakeSlideDoc`)
  - *Jak sprawdzić:* uruchom `python -m pytest -q tests/test_gui.py`.
  - *Oczekiwany efekt:* 16 testów przechodzi (`16 passed`); wyszukiwarka, wielokrotny wybór wyników (Ctrl/Shift – D13), blokada *Cytuj* bez kursora w polu tekstowym (D10), lista czcionek z systemu (D5), 3 położenia pola źródeł i przełącznik nagłówka „Źródła:” (D9), natychmiastowa przebudowa prezentacji z kopią zapasową `.bak.pptx` po zmianie ustawień (D14) wraz z wycofaniem wyboru przy błędzie Zotero oraz opcja „Zawsze na wierzchu” działają poprawnie.
  - *Nie przechodzi, gdy:* którykolwiek z 16 testów w `tests/test_gui.py` zgłasza błąd.
- **B2.2 Silnik numeracji i stylów (`azp.engine`)** – **Automatycznie:** `tests/test_engine.py` (21 testów czystych funkcji i `CitationEngine`)
  - *Jak sprawdzić:* uruchom `python -m pytest -q tests/test_engine.py`.
  - *Oczekiwany efekt:* 21 testów przechodzi (`21 passed`); numeracja IEEE globalnie i per slajd (D4), ten sam numer dla powtórzeń, cytowania zbiorcze i automatyczne scalanie sąsiednich cytowań przy kursorze (`[1, 2]` oraz `(Autor1; Autor2)` – D13), alfabetyczne sortowanie stopki w stylu APA 7 (D3), blokada bez kursora (D10), blokada przebudowy niezapisanej prezentacji (D18) oraz natychmiastowa przebudowa z kopią zapasową `.bak.pptx` po zmianie ustawień lub kolejności slajdów (D14) działają poprawnie.
  - *Nie przechodzi, gdy:* którykolwiek z 21 testów w `tests/test_engine.py` zgłasza błąd.
- **C1.1 Integracja okienka z prawdziwym PowerPointem i Zotero (`azp.main`, `python -m azp`)** – **Automatycznie:** `tests/test_main.py` (4 testy jednostkowe + 1 test końcowy COM `@pytest.mark.com`), `tests/test_ppt.py` (25 testów jednostkowych + 1 test COM), `tests/test_gui.py` (20 testów)
  - *Jak sprawdzić:* przy włączonym Zotero 10 i zamkniętym PowerPoincie uruchom `python -m pytest -q` oraz `python -m pytest -q -m com`.
  - *Oczekiwany efekt:* przechodzą wszystkie testy bez COM (`107 passed`) oraz wszystkie testy COM (`3 passed`); sprawdzane są: zachowanie formatowania tekstu użytkownika na slajdzie przy wstawianiu, scalaniu i przenumerowaniu cytowań (T10), kopia zapasowa przez `SaveCopyAs` bez zapisu oryginalnego pliku użytkownika (D17), odmowa przebudowy niezapisanej prezentacji z komunikatem „Zapisz prezentację, żeby można było zrobić kopię zapasową.” i wycofaniem zmiany na liście (D18), śledzenie ustawień aktywnej prezentacji przy przełączaniu między plikami (D19) oraz powrót fokusu i kursora za wstawionym cytowaniem po *Cytuj* (D20).
  - *Nie przechodzi, gdy:* Zotero 10 jest wyłączone, PowerPoint użytkownika jest otwarty przed testami COM albo którykolwiek test zgłasza błąd.

## Wszystkie punkty
- A1.1 Odpowiedź lokalnego API Zotero dla stylów IEEE v4 i APA
- A1.2 Sterowanie PowerPointem przez COM i trwałość Tags
- A2.1 Zgodność kontraktów `ZoteroSource` i `SlideDoc` oraz atrap testowych
- B1.1 Klient lokalnego API Zotero i parser HTML CSL
- B2.1 Warstwa COM PowerPoint (`PowerPointSlideDoc`)
- B1.2 Okienko GUI (`AzpWindow` / `AzpController`)
- B2.2 Silnik numeracji i stylów (`azp.engine`)
- C1.1 Integracja okienka z prawdziwym PowerPointem i Zotero (`azp.main`, `python -m azp`)
