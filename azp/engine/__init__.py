"""Silnik numeracji cytowań i przebudowy stylów (Etap B2.2)."""

from __future__ import annotations

from azp.engine.numbering import (
    assign_global_numbers,
    assign_slide_numbers,
    build_presentation_bibliography_entries,
    build_slide_footer_entries,
    collect_presentation_keys,
    compute_presentation_plan,
    format_citation_text,
    is_numeric_style,
    merge_citation_keys,
)
from azp.engine.service import CitationEngine

__all__ = [
    "CitationEngine",
    "assign_global_numbers",
    "assign_slide_numbers",
    "build_presentation_bibliography_entries",
    "build_slide_footer_entries",
    "collect_presentation_keys",
    "compute_presentation_plan",
    "format_citation_text",
    "is_numeric_style",
    "merge_citation_keys",
]
