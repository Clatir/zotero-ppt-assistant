# Asystent Zotero dla PowerPointa (AZP)

Pływające okienko (Windows) do cytowania pozycji ze Zotero 10 w PowerPoincie: cytowanie w tekście slajdu,
źródła w stopce pod slajdem, przenumerowanie i przełączanie stylów (IEEE v4 / APA). Interfejs po polsku.
Repozytorium publiczne (`Clatir/zotero-ppt-assistant`) – bez danych użytkownika i osobistych treści (D12).
Zasady pracy przeniesione z `Clatir/RegulatedDay` (CLAUDE.md, „Zasady współpracy”, skill `etap`), dostosowane do Antigravity.

## Zasady współpracy
- **Przed każdym ciągiem / etapem** komunikat „Przed etapem N”: co trzeba wiedzieć, pytania z rekomendacją (zawsze
  klikalne), co przygotować, model i thinking. **Po etapie:** testy → przegląd kodu → commit → wpis w `docs/postep.md`
  → lista „do sprawdzenia” w `docs/odbior.md` → pauza do „dalej”. Bez „dalej” nie zaczynaj kolejnego etapu.
  W ciągu (`docs/plan.md`) „Przed etapem” tylko na jego początku, pauza dopiero na końcu ciągu albo przy bramie.
- Zachowanie widoczne dla użytkownika zmienia się tylko po pytaniu, nigdy zgodą hurtem. Sprawy techniczne rozstrzygaj
  sam i dopisz jako T-decyzję (`docs/decyzje-techniczne.md`). Decyzje użytkownika – D (`docs/decyzje.md`); nie zmieniaj
  D bez pytania, nowa decyzja wskazuje, którą zastępuje.
- **Główna sesja** (Claude Opus 5.5) zleca test i kod każdego etapu **wykonawcy** (subagent) według kolumny „Wykonawca”
  w `docs/plan.md`, a sama sprawdza, robi przegląd, commit i wpisy w dokumentach. Wykonawca nie commituje.
- Do gałęzi `main` pisze naraz jedna sesja. Wykonawcy równoległi piszą tylko w swoich, rozłącznych katalogach.
- Po każdym etapie commit i `git push` na GitHuba (`Clatir/zotero-ppt-assistant`, D12).
- **Przegląd kodu** (subagent-recenzent) według ryzyka: *xhigh* – COM PowerPointa i silnik numeracji (zmieniają plik
  użytkownika); *high* – reszta kodu; *medium* – dokumenty. Najwyżej trzy rundy (druga i trzecia tylko na poprawkach).
  Poprawiamy critical i high zawsze, z testem; medium – gdy da się napisać test, który bez poprawki pada. Reszta →
  „Uwagi na później” w postępie. Z przeglądu nie wchodzi zmiana zachowania widocznego dla użytkownika bez pytania.
- Po wznowieniu (limit, nowa rozmowa) najpierw przeczytaj `docs/postep.md`.
- **Czysty start (`/compact` albo nowa rozmowa):** przy każdej pauzie powiedz wprost, czy można go teraz bezpiecznie
  zrobić. Bezpiecznie = wszystko w commitach wypchniętych na GitHuba, `docs/postep.md` mówi, co dalej, nic nie zależy od pamięci rozmowy.

## Zasady produktu
- Narzędzie nigdy nie psuje prezentacji: przed każdą przebudową całej prezentacji kopia zapasowa pliku.
- Zotero tylko do odczytu (lokalne API), nigdy zapis do `zotero.sqlite`.
- Okienko „zawsze na wierzchu” – z możliwością wyłączenia.

## Dokumenty
- `docs/plan.md` – tylko praca do zrobienia: ciągi, etapy, rodzaj, wykonawca, szacunek, „gotowe, gdy”.
- `docs/postep.md` – stan bieżący i kolejka (≤ 20 KB).
- `docs/odbior.md` – punkty do odbioru na Windows (Jak sprawdzić / Oczekiwany efekt / Nie przechodzi, gdy).
- `docs/decyzje.md` (D), `docs/decyzje-techniczne.md` (T).

## Stos
- Python 3.12, tkinter (GUI), pywin32 (COM PowerPointa), requests (lokalne API Zotero `127.0.0.1:23119/api/`), pytest.
- `azp/zotero/` – klient Zotero; `azp/ppt/` – warstwa COM; `azp/engine/` – numeracja i style (czyste funkcje);
  `azp/gui/` – okienko; `azp/contracts.py` – wspólne interfejsy; `tests/` – testy (`fakes/` – atrapy).

## Testy
- Każda poprawka ma test, który na starym kodzie nie przechodzi. W trakcie tylko powiązany plik (`pytest -q plik`),
  pełny zestaw raz przed commitem.
- Testy COM (`@pytest.mark.com`) działają na prezentacji testowej w `tests/data/`, nigdy na plikach użytkownika;
  uruchamiane tylko przy zamkniętym PowerPoincie użytkownika.

## Komendy
```
python -m pip install -r requirements.txt
python -m pytest -q                 # bez testów COM
python -m pytest -q -m com          # testy COM (PowerPoint zamknięty)
python -m azp                       # uruchom okienko
```
