"""Testy kontraktów (`azp/contracts.py`) oraz atrap (`tests/fakes/`)."""

from __future__ import annotations

import pytest

from azp.contracts import (
    STYLE_APA,
    STYLE_IEEE_V4,
    Citation,
    FooterEntry,
    NoActiveSlideError,
    NoTextCursorError,
    Settings,
    SlideDoc,
    TextSpan,
    ZoteroItem,
    ZoteroOfflineError,
    ZoteroSource,
    spans_to_plain_text,
)
from tests.fakes import FakeSlideDoc, FakeZoteroSource


def test_fakes_satisfy_runtime_protocols() -> None:
    zotero = FakeZoteroSource()
    doc = FakeSlideDoc()
    assert isinstance(zotero, ZoteroSource)
    assert isinstance(doc, SlideDoc)


def test_text_spans_and_footer_entry_formatting() -> None:
    spans = (
        TextSpan("A. Kowalski, "),
        TextSpan("Genome Med", italic=True),
        TextSpan(", 2023."),
    )
    assert spans_to_plain_text(spans) == "A. Kowalski, Genome Med, 2023."

    entry = FooterEntry(key="KOWAL2023", prefix="[1] ", spans=spans)
    full = entry.full_spans()
    assert full[0] == TextSpan("[1] ", italic=False)
    assert full[2].italic is True
    assert entry.plain_text == "[1] A. Kowalski, Genome Med, 2023."


def test_zotero_item_labels() -> None:
    no_author = ZoteroItem(key="K0", title="Tytuł", authors=())
    assert no_author.authors_label == "(brak autora)"

    one_author = ZoteroItem(key="K1", title="Tytuł", authors=("Kowalski, A.",), year="2023")
    assert one_author.display_label == "Kowalski, A. (2023) – Tytuł"

    two_authors = ZoteroItem(
        key="K2", title="Tytuł", authors=("Nowak, B.", "Wiśniewski, P."), year="2024"
    )
    assert two_authors.authors_label == "Nowak, B., Wiśniewski, P."

    three_authors = ZoteroItem(
        key="K3",
        title="Tytuł",
        authors=("Dove, E.", "Knoppers, B.", "Zawati, M."),
        year="2013",
    )
    assert three_authors.authors_label == "Dove, E. i in."


def test_settings_validation_and_roundtrip() -> None:
    s = Settings(
        style_id=STYLE_APA,
        numbering_mode="per_slide",
        font_name="Arial",
        font_size_pt=11,
        footer_position="bottom_left",
        footer_header=False,
        always_on_top=False,
    )
    restored = Settings.from_dict(s.to_dict())
    assert restored == s

    with pytest.raises(ValueError, match="tryb numeracji"):
        Settings(numbering_mode="invalid")  # type: ignore[arg-type]

    with pytest.raises(ValueError, match="położenie stopki"):
        Settings(footer_position="top")  # type: ignore[arg-type]

    with pytest.raises(ValueError, match="poza zakresem"):
        Settings(font_size_pt=4)

    with pytest.raises(ValueError, match="Nazwa czcionki"):
        Settings(font_name="   ")


def test_citation_requires_non_empty_keys() -> None:
    with pytest.raises(ValueError, match="nie może być puste"):
        Citation(citation_id="c1", keys=())


def test_fake_zotero_search_and_formatting() -> None:
    zotero = FakeZoteroSource()
    assert zotero.is_available() is True

    styles = zotero.list_styles()
    style_ids = {st.style_id for st in styles}
    assert STYLE_IEEE_V4 in style_ids
    assert STYLE_APA in style_ids

    found = zotero.search_items("Kowalski")
    assert len(found) == 1
    assert found[0].key == "KOWAL2023"

    ref_ieee = zotero.get_formatted_ref("KOWAL2023", STYLE_IEEE_V4)
    assert ref_ieee.in_text == "[1]"
    assert any(span.italic for span in ref_ieee.bib_body_spans)
    assert "Przegląd Informatyczny" in ref_ieee.bib_plain

    ref_apa = zotero.get_formatted_ref("KOWAL2023", STYLE_APA)
    assert ref_apa.in_text == "(Kowalski, 2023)"

    batch = zotero.get_formatted_refs(["KOWAL2023", "NOWAK2024"], STYLE_APA)
    assert set(batch.keys()) == {"KOWAL2023", "NOWAK2024"}


def test_fake_zotero_offline_raises_error() -> None:
    zotero = FakeZoteroSource(offline=True)
    assert zotero.is_available() is False
    with pytest.raises(ZoteroOfflineError, match="Zotero nie odpowiada"):
        zotero.search_items("AI")
    with pytest.raises(ZoteroOfflineError):
        zotero.get_formatted_ref("KOWAL2023", STYLE_IEEE_V4)


def test_fake_slide_doc_no_cursor_leaves_presentation_unchanged_d10() -> None:
    doc = FakeSlideDoc(slide_count=2, active_slide_index=1, has_cursor=False)
    before_text = doc.get_slide_body_text(1)
    before_slides = doc.read_slides()

    with pytest.raises(NoTextCursorError, match="kursor"):
        doc.insert_citation_at_cursor(
            Citation(citation_id="cit-1", keys=("KOWAL2023",), text="[1]")
        )

    assert doc.get_slide_body_text(1) == before_text
    assert doc.read_slides() == before_slides


def test_fake_slide_doc_no_active_slide_raises() -> None:
    doc = FakeSlideDoc(slide_count=2, active_slide_index=None)
    assert doc.has_open_presentation() is False
    with pytest.raises(NoActiveSlideError):
        doc.get_active_slide_index()


def test_fake_slide_doc_citation_footer_reorder_and_backup() -> None:
    doc = FakeSlideDoc(slide_count=2, active_slide_index=1, has_cursor=True)
    settings = Settings(footer_position="replace_footer", footer_header=False)
    doc.save_settings(settings)
    assert doc.load_settings() == settings

    cit = Citation(citation_id="cit-1", keys=("KOWAL2023",), text="[1]")
    doc.insert_citation_at_cursor(cit)
    assert doc.get_slide_body_text(1).endswith("[1]")

    entry = FooterEntry(
        key="KOWAL2023",
        prefix="[1] ",
        spans=(TextSpan("A. Kowalski, 2023."),),
    )
    doc.set_slide_footer(1, [entry], settings)
    assert doc.get_slide_footer_settings(1) == settings

    # Aktualizacja tekstu cytowania (np. po przenumerowaniu lub zmianie na APA)
    updated_cit = Citation(citation_id="cit-1", keys=("KOWAL2023",), text="(Kowalski, 2023)")
    doc.update_slide_citations(1, [updated_cit])
    assert "(Kowalski, 2023)" in doc.get_slide_body_text(1)

    # Przestawienie slajdów
    doc.reorder_slides([2, 1])
    slides = doc.read_slides()
    assert slides[0].citations == ()
    assert slides[1].citations == (updated_cit,)

    # Kopia zapasowa
    backup = doc.create_backup()
    assert backup is not None and backup.name.endswith(".bak.pptx")
    assert len(doc.backups_created) == 1


def test_fake_slide_doc_update_citations_handles_swaps_and_duplicates() -> None:
    doc = FakeSlideDoc(slide_count=1, active_slide_index=1, has_cursor=True)
    # Dwa cytowania wstawione z tym samym tekstem początkowym "[1]"
    c1 = Citation(citation_id="cit-1", keys=("KOWAL2023",), text="[1]")
    c2 = Citation(citation_id="cit-2", keys=("NOWAK2024",), text="[1]")
    doc.insert_citation_at_cursor(c1)
    doc.insert_citation_at_cursor(c2)
    assert doc.get_slide_body_text(1) == "Tekst slajdu 1 [1] [1]"

    # Pierwsze przenumerowanie: cit-1 -> [1], cit-2 -> [2]
    doc.update_slide_citations(
        1,
        [
            Citation(citation_id="cit-1", keys=("KOWAL2023",), text="[1]"),
            Citation(citation_id="cit-2", keys=("NOWAK2024",), text="[2]"),
        ],
    )
    assert doc.get_slide_body_text(1) == "Tekst slajdu 1 [1] [2]"

    # Zamiana numerów miejscami: cit-1 -> [2], cit-2 -> [1]
    doc.update_slide_citations(
        1,
        [
            Citation(citation_id="cit-1", keys=("KOWAL2023",), text="[2]"),
            Citation(citation_id="cit-2", keys=("NOWAK2024",), text="[1]"),
        ],
    )
    assert doc.get_slide_body_text(1) == "Tekst slajdu 1 [2] [1]"

