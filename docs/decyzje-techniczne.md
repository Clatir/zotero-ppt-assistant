# Decyzje techniczne (T)

| ID | Decyzja | Powód |
|---|---|---|
| T1 | Python 3.12 + tkinter + pywin32, nie Office Web Add-in ani VSTO. | Python już jest; brak instalacji dodatku w Office; najkrótsza droga. |
| T2 | Dane i formatowanie ze Zotero tylko przez lokalne API (`/api/users/0/...`, `format=bib`, `include=citation,bib`, `style=<ID>`). | Formatowanie identyczne jak w Wordzie (silnik CSL Zotero); `zotero.sqlite` jest zablokowany, gdy Zotero działa. Sprawdzone 2026-11-27: styl v4 zwraca HTML z `<i>`. |
| T3 | Powiązania cytowań trzymane w *Tags* PowerPointa (prezentacja, slajd, kształt stopki). | Przetrwają zapis i ponowne otwarcie pliku; pozwalają przebudować numerację i styl. |
| T4 | Kursywa i inne znaczniki z HTML-a CSL przenoszone do formatowania tekstu stopki. | Zgodność wyglądu z Wordem. |
| T5 | Wykonawcy w Antigravity: subagenci (`inherit` = Opus 5.5 do zadań xhigh, `flash` do prostych). Odpowiednik `wykonawca-opus` / `wykonawca-haiku` z RD. | Brak plików `.claude/agents` w Antigravity. |
| T6 | Odpowiednik `/code-review`: subagent-recenzent z poziomem ryzyka w poleceniu; `/clear` → `/compact` albo nowa rozmowa. | Brak tych komend w Antigravity. |
