# Plan – praca do zrobienia

Ciąg = etapy idące jednym ciągiem: „Przed etapem” na początku ciągu, pauza do „dalej” na końcu ciągu albo przy bramie.
Po każdym etapie (także w ciągu): testy → przegląd → commit → wpis w postępie → punkty odbioru.

## Ciąg A – Fundament (główna sesja, sekwencyjnie)
| Etap | Rodzaj | Wykonawca | Przegląd | Szacunek | Gotowe, gdy |
|---|---|---|---|---|---|
| A1 Środowisko | konfiguracja | główna sesja (Opus 5.5) | medium | 10 min | `git init`, `requirements.txt`, `pywin32`, `pytest`; test „Zotero odpowiada” i test COM „otwiera i zamyka prezentację testową” przechodzą |
| A2 Kontrakty i atrapy | kod | główna sesja (Opus 5.5) | high | 15 min | `azp/contracts.py` (typy: `ZoteroItem`, `Citation`, `Settings`; interfejsy `ZoteroSource`, `SlideDoc`); atrapy `tests/fakes/`; testy kontraktów zielone |

**Brama A→B:** kontrakty zatwierdzone (bez nich tory B nie mogą iść równolegle).

## Ciąg B – Rdzeń (dwa tory równolegle)
| Etap | Rodzaj | Wykonawca | Przegląd | Szacunek | Gotowe, gdy |
|---|---|---|---|---|---|
| B1.1 Klient Zotero | kod | subagent `flash` | high | 30–45 min | wyszukiwanie po autorze/tytule/roku; cytowanie i opis dla v4 i APA; lista stylów; HTML → przebiegi tekstu z kursywą; komunikat „Zotero nie odpowiada” |
| B1.2 Okienko | kod | subagent `flash` | high | 45–60 min | wyszukiwarka, wyniki, *Cytuj*, *Odśwież / Przenumeruj*; listy: styl, tryb numeracji, czcionka (z systemu), rozmiar; „zawsze na wierzchu”; działa z atrapą `SlideDoc` |
| B2.1 Warstwa PowerPoint | kod | subagent `inherit` (Opus 5.5) | **xhigh** | 1–1,5 h | aktywny slajd i kursor; wstawienie cytowania; pole „Źródła” na dole slajdu (czcionka, rozmiar, kursywa, autodopasowanie); tagi; testy COM na prezentacji testowej |
| B2.2 Silnik numeracji i stylów | kod | subagent `inherit` (Opus 5.5) | **xhigh** | 1–1,5 h | IEEE globalnie / per slajd, ten sam numer dla powtórzeń, `[1, 2]`; APA autor-rok, alfabetycznie; przebudowa po zmianie stylu; kopia zapasowa przed przebudową; czyste funkcje z testami |

Tor B1 = B1.1 → B1.2; tor B2 = B2.1 → B2.2; tory równolegle, rozłączne katalogi (`azp/zotero`, `azp/gui` | `azp/ppt`, `azp/engine`).

## Ciąg C – Integracja i odbiór (główna sesja + Ty)
| Etap | Rodzaj | Wykonawca | Przegląd | Szacunek | Gotowe, gdy |
|---|---|---|---|---|---|
| C1 Integracja | kod | główna sesja (Opus 5.5) | xhigh | 30–45 min | okienko steruje prawdziwym PowerPointem i Zotero; test końcowy na prezentacji testowej |
| C2 Dodatki | kod | subagent `flash` | high | 20–30 min | skrót klawiszowy, `config.json`, uruchamianie ikoną (`.bat`, opcjonalnie `.exe`) |
| C3 Odbiór na Twojej prezentacji | odbiór | Ty + główna sesja | – | 30–60 min | wszystkie punkty w `docs/odbior.md` odebrane |

**Łącznie:** ~3–4 h zegarowo (ciąg B równolegle).
