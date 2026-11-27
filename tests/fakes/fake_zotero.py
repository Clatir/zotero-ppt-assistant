"""Atrapa `ZoteroSource` do niezależnego testowania GUI (B1.2) oraz silnika (B2.2)."""

from __future__ import annotations

from typing import Sequence

from azp.contracts import (
    STYLE_APA,
    STYLE_IEEE_V4,
    FormattedRef,
    StyleInfo,
    TextSpan,
    ZoteroItem,
    ZoteroOfflineError,
)


def _default_items() -> list[ZoteroItem]:
    return [
        ZoteroItem(
            key="KOWAL2023",
            title="Klasyfikacja systemów sztucznej inteligencji wysokiego ryzyka",
            authors=("Kowalski, Adam",),
            year="2023",
            item_type="journalArticle",
        ),
        ZoteroItem(
            key="NOWAK2024",
            title="Algorytmy uczenia maszynowego w praktyce inżynierskiej",
            authors=("Nowak, Barbara", "Wiśniewski, Piotr"),
            year="2024",
            item_type="book",
        ),
        ZoteroItem(
            key="DOVE2013",
            title="An ethics safe harbor for international genomics research?",
            authors=("Dove, Edward S.", "Knoppers, Bartha M.", "Zawati, Ma'n H."),
            year="2013",
            item_type="journalArticle",
        ),
        ZoteroItem(
            key="ZIEL2022",
            title="Przegląd metod walidacji modeli predykcyjnych",
            authors=("Zieliński, Marek",),
            year="2022",
            item_type="conferencePaper",
        ),
    ]


def _default_formatted_refs() -> dict[tuple[str, str], FormattedRef]:
    return {
        ("KOWAL2023", STYLE_IEEE_V4): FormattedRef(
            key="KOWAL2023",
            style_id=STYLE_IEEE_V4,
            in_text="[1]",
            bib_body_spans=(
                TextSpan('A. Kowalski, "Klasyfikacja systemów sztucznej inteligencji wysokiego ryzyka," '),
                TextSpan("Przegląd Informatyczny", italic=True),
                TextSpan(", vol. 12, no. 3, pp. 45-58, 2023, [dostęp: 2026-11-27]."),
            ),
            sort_key="Kowalski, Adam (2023)",
        ),
        ("KOWAL2023", STYLE_APA): FormattedRef(
            key="KOWAL2023",
            style_id=STYLE_APA,
            in_text="(Kowalski, 2023)",
            bib_body_spans=(
                TextSpan("Kowalski, A. (2023). Klasyfikacja systemów sztucznej inteligencji wysokiego ryzyka. "),
                TextSpan("Przegląd Informatyczny", italic=True),
                TextSpan(", "),
                TextSpan("12", italic=True),
                TextSpan("(3), 45–58."),
            ),
            sort_key="Kowalski, A. (2023)",
        ),
        ("NOWAK2024", STYLE_IEEE_V4): FormattedRef(
            key="NOWAK2024",
            style_id=STYLE_IEEE_V4,
            in_text="[1]",
            bib_body_spans=(
                TextSpan("B. Nowak and P. Wiśniewski, "),
                TextSpan("Algorytmy uczenia maszynowego w praktyce inżynierskiej", italic=True),
                TextSpan(". Warszawa: Wydawnictwo Naukowe, 2024."),
            ),
            sort_key="Nowak, Barbara (2024)",
        ),
        ("NOWAK2024", STYLE_APA): FormattedRef(
            key="NOWAK2024",
            style_id=STYLE_APA,
            in_text="(Nowak & Wiśniewski, 2024)",
            bib_body_spans=(
                TextSpan("Nowak, B., & Wiśniewski, P. (2024). "),
                TextSpan("Algorytmy uczenia maszynowego w praktyce inżynierskiej", italic=True),
                TextSpan(". Wydawnictwo Naukowe."),
            ),
            sort_key="Nowak, B., & Wiśniewski, P. (2024)",
        ),
        ("DOVE2013", STYLE_IEEE_V4): FormattedRef(
            key="DOVE2013",
            style_id=STYLE_IEEE_V4,
            in_text="[1]",
            bib_body_spans=(
                TextSpan('E. S. Dove, B. M. Knoppers, and M. H. Zawati, "An ethics safe harbor for international genomics research?," '),
                TextSpan("Genome Med", italic=True),
                TextSpan(", vol. 5, no. 11, p. 99, 2013, doi: 10.1186/gm503."),
            ),
            sort_key="Dove, Edward S. (2013)",
        ),
        ("DOVE2013", STYLE_APA): FormattedRef(
            key="DOVE2013",
            style_id=STYLE_APA,
            in_text="(Dove et al., 2013)",
            bib_body_spans=(
                TextSpan("Dove, E. S., Knoppers, B. M., & Zawati, M. H. (2013). An ethics safe harbor for international genomics research? "),
                TextSpan("Genome Medicine", italic=True),
                TextSpan(", "),
                TextSpan("5", italic=True),
                TextSpan("(11), 99. https://doi.org/10.1186/gm503"),
            ),
            sort_key="Dove, E. S., Knoppers, B. M., & Zawati, M. H. (2013)",
        ),
        ("ZIEL2022", STYLE_IEEE_V4): FormattedRef(
            key="ZIEL2022",
            style_id=STYLE_IEEE_V4,
            in_text="[1]",
            bib_body_spans=(
                TextSpan('M. Zieliński, "Przegląd metod walidacji modeli predykcyjnych," in '),
                TextSpan("Proc. Konf. Inżynierii Oprogramowania", italic=True),
                TextSpan(", Kraków, 2022, pp. 101-108."),
            ),
            sort_key="Zieliński, Marek (2022)",
        ),
        ("ZIEL2022", STYLE_APA): FormattedRef(
            key="ZIEL2022",
            style_id=STYLE_APA,
            in_text="(Zieliński, 2022)",
            bib_body_spans=(
                TextSpan("Zieliński, M. (2022). Przegląd metod walidacji modeli predykcyjnych. "),
                TextSpan("Proc. Konf. Inżynierii Oprogramowania", italic=True),
                TextSpan(", 101–108."),
            ),
            sort_key="Zieliński, M. (2022)",
        ),
    }


class FakeZoteroSource:
    """Atrapa źródła Zotero implementująca protokół `ZoteroSource`."""

    def __init__(
        self,
        items: Sequence[ZoteroItem] | None = None,
        refs: dict[tuple[str, str], FormattedRef] | None = None,
        offline: bool = False,
    ) -> None:
        self._items: list[ZoteroItem] = list(items) if items is not None else _default_items()
        self._refs: dict[tuple[str, str], FormattedRef] = (
            dict(refs) if refs is not None else _default_formatted_refs()
        )
        self.offline: bool = offline
        self._styles: list[StyleInfo] = [
            StyleInfo(
                style_id=STYLE_IEEE_V4,
                title="IEEE Praca Inzynierska Access Date v4",
                is_numeric=True,
            ),
            StyleInfo(
                style_id=STYLE_APA,
                title="APA Style 7th edition",
                is_numeric=False,
            ),
        ]

    def _check_online(self) -> None:
        if self.offline:
            raise ZoteroOfflineError(
                "Zotero nie odpowiada. Uruchom Zotero 10 i upewnij się, że lokalne API jest włączone."
            )

    def is_available(self) -> bool:
        return not self.offline

    def search_items(self, query: str = "", limit: int = 25) -> list[ZoteroItem]:
        self._check_online()
        q = query.strip().casefold()
        if not q:
            return self._items[:limit]

        results: list[ZoteroItem] = []
        for item in self._items:
            haystack = " ".join(
                [item.title, item.year, *item.authors, item.key]
            ).casefold()
            if q in haystack:
                results.append(item)
                if len(results) >= limit:
                    break
        return results

    def get_formatted_ref(self, key: str, style_id: str) -> FormattedRef:
        self._check_online()
        pair = (key, style_id)
        if pair not in self._refs:
            raise KeyError(f"Brak sformatowanego opisu dla klucza {key!r} i stylu {style_id!r}")
        return self._refs[pair]

    def get_formatted_refs(
        self, keys: Sequence[str], style_id: str
    ) -> dict[str, FormattedRef]:
        self._check_online()
        return {k: self.get_formatted_ref(k, style_id) for k in keys}

    def list_styles(self) -> list[StyleInfo]:
        return list(self._styles)
