"""Silnik operacji cytowania i przebudowy prezentacji (`CitationEngine`, Etap B2.2)."""

from __future__ import annotations

from pathlib import Path
from typing import Callable, Sequence
import uuid

from azp.contracts import (
    MSG_SAVE_BEFORE_REBUILD,
    AzpError,
    BackupError,
    Citation,
    NoActiveSlideError,
    NoTextCursorError,
    Settings,
    SlideDoc,
    SlideRecord,
    StyleInfo,
    ZoteroSource,
)
from azp.engine.numbering import (
    assign_global_numbers,
    assign_slide_numbers,
    compute_presentation_plan,
    format_citation_text,
    is_numeric_style,
    merge_citation_keys,
)


class CitationEngine:
    """Koordynator łączący źródło danych Zotero, czystą numerację i dokument PowerPoint."""

    def __init__(self, zotero: ZoteroSource, slide_doc: SlideDoc) -> None:
        self.zotero = zotero
        self.slide_doc = slide_doc
        self.last_backup_path: Path | None = None
        self.last_citation: Citation | None = None

    def _resolve_settings(self, settings: Settings | None) -> Settings:
        if settings is not None:
            return settings
        loaded = self.slide_doc.load_settings()
        return loaded if loaded is not None else Settings()

    def _get_styles(self) -> list[StyleInfo]:
        try:
            return list(self.zotero.list_styles())
        except AzpError:
            return []

    @staticmethod
    def _new_citation_id() -> str:
        return f"cit-{uuid.uuid4().hex[:12]}"

    def cite(
        self,
        keys: Sequence[str],
        settings: Settings | None = None,
    ) -> SlideRecord:
        """Wstawia cytowanie (pojedyncze lub zbiorcze) w miejscu kursora i przelicza slajdy.

        Przestrzega D10 (brak kursora w polu tekstowym -> `NoTextCursorError`, zero zmian)
        oraz D13 (wielokrotny wybór i automatyczne scalanie sąsiednich cytowań przy kursorze).
        Nie tworzy kopii zapasowej przy pojedynczym wstawieniu cytowania.
        """
        unique_keys = tuple(dict.fromkeys(keys))
        if not unique_keys or any(
            not isinstance(k, str) or not k.strip() for k in unique_keys
        ):
            raise ValueError("Lista kluczy do zacytowania nie może być pusta.")

        if not self.slide_doc.has_open_presentation():
            raise NoActiveSlideError(
                "Brak otwartej prezentacji lub aktywnego slajdu w PowerPoincie."
            )

        # D10: wymagamy kursora w polu tekstowym przed jakąkolwiek zmianą i przed pobieraniem indeksu slajdu
        if not self.slide_doc.has_text_cursor():
            raise NoTextCursorError(
                "Ustaw kursor w polu tekstowym na slajdzie przed wstawieniem cytowania."
            )

        active_slide_index = self.slide_doc.get_active_slide_index()

        effective_settings = self._resolve_settings(settings)
        slides_before = self.slide_doc.read_slides()

        get_adj = getattr(self.slide_doc, "get_adjacent_citation_at_cursor", None)
        adjacent_citation: Citation | None = (
            get_adj() if callable(get_adj) else None
        )
        if adjacent_citation is not None:
            target_citation_id = adjacent_citation.citation_id
            target_keys = merge_citation_keys(adjacent_citation.keys, unique_keys)
        else:
            target_citation_id = self._new_citation_id()
            target_keys = unique_keys

        all_keys = tuple(
            dict.fromkeys(
                [
                    *(k for s in slides_before for c in s.citations for k in c.keys),
                    *target_keys,
                ]
            )
        )
        # Pobierz dane z Zotero PRZED modyfikacją slajdu (ochrona przed ZoteroOfflineError)
        refs = self.zotero.get_formatted_refs(all_keys, effective_settings.style_id)
        styles = self._get_styles()
        numeric = is_numeric_style(effective_settings.style_id, styles)

        if numeric:
            if effective_settings.numbering_mode == "global":
                slides_up_to_active = [
                    s
                    for s in slides_before
                    if s.slide_index <= active_slide_index
                ]
                base_map = assign_global_numbers(slides_up_to_active)
            else:
                active_before = next(
                    (
                        s
                        for s in slides_before
                        if s.slide_index == active_slide_index
                    ),
                    None,
                )
                base_map = (
                    assign_slide_numbers(active_before.citations)
                    if active_before is not None
                    else {}
                )
            provisional_map = dict(base_map)
            for key in target_keys:
                if key not in provisional_map:
                    provisional_map[key] = len(provisional_map) + 1
        else:
            provisional_map = None

        initial_text = format_citation_text(
            target_keys,
            refs,
            is_numeric=numeric,
            number_map=provisional_map,
        )
        new_citation = Citation(
            citation_id=target_citation_id,
            keys=target_keys,
            text=initial_text,
        )

        self.slide_doc.insert_citation_at_cursor(new_citation)

        slides_after = self.slide_doc.read_slides()
        planned_slides = compute_presentation_plan(
            slides=slides_after,
            refs=refs,
            settings=effective_settings,
            styles=styles,
        )

        active_planned: SlideRecord | None = None
        for current, planned in zip(slides_after, planned_slides):
            is_active = planned.slide_index == active_slide_index
            citations_changed = planned.citations != current.citations
            footer_changed = planned.footer_entries != current.footer_entries
            if citations_changed:
                self.slide_doc.update_slide_citations(
                    planned.slide_index,
                    planned.citations,
                )
            if is_active or citations_changed or footer_changed:
                self.slide_doc.set_slide_footer(
                    planned.slide_index,
                    planned.footer_entries,
                    effective_settings,
                )
            if is_active:
                active_planned = planned

        self.slide_doc.save_settings(effective_settings)

        if active_planned is None:
            raise NoActiveSlideError(
                f"Nie znaleziono aktywnego slajdu o indeksie {active_slide_index}."
            )

        matched_cit = next(
            (
                c
                for c in active_planned.citations
                if c.citation_id == target_citation_id
            ),
            None,
        )
        if matched_cit is None:
            matched_cit = next(
                (
                    c
                    for c in reversed(active_planned.citations)
                    if all(k in c.keys for k in unique_keys)
                ),
                active_planned.citations[-1] if active_planned.citations else new_citation,
            )
        self.last_citation = matched_cit
        return active_planned

    def rebuild(
        self,
        settings: Settings | None = None,
        create_backup: bool = True,
    ) -> list[SlideRecord]:
        """Przebudowuje numerację, styl i stopki w całej prezentacji (z kopią zapasową)."""
        if not self.slide_doc.has_open_presentation():
            raise NoActiveSlideError(
                "Brak otwartej prezentacji lub aktywnego slajdu w PowerPoincie."
            )

        if create_backup and self.slide_doc.get_presentation_path() is None:
            raise BackupError(MSG_SAVE_BEFORE_REBUILD)

        effective_settings = self._resolve_settings(settings)
        slides = self.slide_doc.read_slides()

        all_keys = tuple(
            dict.fromkeys(
                k for slide in slides for cit in slide.citations for k in cit.keys
            )
        )
        # Pobierz opisy ze Zotero PRZED utworzeniem kopii zapasowej i zmianą prezentacji
        refs = (
            self.zotero.get_formatted_refs(all_keys, effective_settings.style_id)
            if all_keys
            else {}
        )
        styles = self._get_styles()

        if create_backup:
            self.last_backup_path = self.slide_doc.create_backup()
            if self.last_backup_path is None:
                raise BackupError(MSG_SAVE_BEFORE_REBUILD)
        else:
            self.last_backup_path = None

        planned_slides = compute_presentation_plan(
            slides=slides,
            refs=refs,
            settings=effective_settings,
            styles=styles,
        )

        for planned in planned_slides:
            self.slide_doc.update_slide_citations(
                planned.slide_index,
                planned.citations,
            )
            self.slide_doc.set_slide_footer(
                planned.slide_index,
                planned.footer_entries,
                effective_settings,
            )

        self.slide_doc.save_settings(effective_settings)
        return planned_slides

    def apply_settings(self, new_settings: Settings) -> list[SlideRecord]:
        """Natychmiastowo przebudowuje całą prezentację z kopią zapasową po zmianie ustawień (D14)."""
        return self.rebuild(settings=new_settings, create_backup=True)

    def make_cite_callback(
        self,
    ) -> Callable[[tuple[str, ...], Settings], str | None]:
        """Zwraca callback `on_cite` kompatybilny z `AzpController`."""

        def _on_cite(keys: tuple[str, ...], settings: Settings) -> str | None:
            try:
                slide_record = self.cite(keys, settings=settings)
            except KeyError as exc:
                detail = exc.args[0] if exc.args else str(exc)
                raise AzpError(
                    f"Nie znaleziono opisu pozycji w Zotero: {detail}"
                ) from exc

            cit_text = (
                self.last_citation.text
                if self.last_citation is not None
                else (
                    slide_record.citations[-1].text
                    if slide_record.citations
                    else ""
                )
            )
            return (
                f"Wstawiono cytowanie {cit_text} na slajdzie {slide_record.slide_index}."
                if cit_text
                else f"Wstawiono cytowanie na slajdzie {slide_record.slide_index}."
            )

        return _on_cite

    def make_rebuild_callback(self) -> Callable[[Settings], str | None]:
        """Zwraca callback `on_rebuild` kompatybilny z `AzpController`."""

        def _on_rebuild(settings: Settings) -> str | None:
            try:
                self.rebuild(settings=settings, create_backup=True)
            except KeyError as exc:
                detail = exc.args[0] if exc.args else str(exc)
                raise AzpError(
                    f"Nie znaleziono opisu pozycji w Zotero: {detail}"
                ) from exc

            if self.last_backup_path is not None:
                return (
                    f"Przebudowano prezentację (kopia zapasowa: {self.last_backup_path.name})."
                )
            return "Przebudowano prezentację."

        return _on_rebuild
