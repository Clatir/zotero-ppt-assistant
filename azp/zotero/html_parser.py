"""Parser HTML formatowania CSL zwracanego przez lokalne API Zotero."""

from __future__ import annotations

from html.parser import HTMLParser
import re

from azp.contracts import TextSpan, spans_to_plain_text

_LEADING_NUMERIC_PREFIX_RE = re.compile(r"^\s*\[\d+\]\s*")
_WHITESPACE_RUN_RE = re.compile(r"[ \t\r\n]+")
_SKIPPED_CLASSES = frozenset({"csl-left-margin", "Z3988"})
_SKIPPED_TAGS = frozenset({"script", "style"})
_VOID_TAGS = frozenset(
    {
        "area",
        "base",
        "br",
        "col",
        "embed",
        "hr",
        "img",
        "input",
        "link",
        "meta",
        "param",
        "source",
        "track",
        "wbr",
    }
)


class _CslHtmlParser(HTMLParser):
    """Konwerter fragmentów HTML CSL na surową listę `TextSpan`."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._spans: list[TextSpan] = []
        self._style_stack: list[tuple[bool, bool]] = [(False, False)]
        self._skip_depth: int = 0

    @property
    def spans(self) -> list[TextSpan]:
        return self._spans

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag_lower = tag.lower()
        if tag_lower in _VOID_TAGS:
            if tag_lower == "br" and self._skip_depth == 0 and self._spans:
                if not self._spans[-1].text.endswith(" "):
                    italic, bold = self._style_stack[-1]
                    self._spans.append(TextSpan(text=" ", italic=italic, bold=bold))
            return

        if self._skip_depth > 0:
            self._skip_depth += 1
            return

        attr_map = {k.lower(): (v or "") for k, v in attrs}
        classes = set(attr_map.get("class", "").split())
        if tag_lower in _SKIPPED_TAGS or bool(classes & _SKIPPED_CLASSES):
            self._skip_depth = 1
            return

        cur_italic, cur_bold = self._style_stack[-1]
        next_italic, next_bold = _next_style(tag_lower, attr_map.get("style", ""), cur_italic, cur_bold)
        self._style_stack.append((next_italic, next_bold))

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() in _VOID_TAGS:
            return

        if self._skip_depth > 0:
            self._skip_depth -= 1
            return

        if len(self._style_stack) > 1:
            self._style_stack.pop()

    def handle_data(self, data: str) -> None:
        if self._skip_depth > 0 or not data:
            return

        normalized = data.replace("\xa0", " ")
        normalized = _WHITESPACE_RUN_RE.sub(" ", normalized)
        if self._spans and self._spans[-1].text.endswith(" ") and normalized.startswith(" "):
            normalized = normalized.lstrip(" ")
        if not normalized:
            return

        italic, bold = self._style_stack[-1]
        self._spans.append(TextSpan(text=normalized, italic=italic, bold=bold))


def _next_style(
    tag: str,
    inline_style: str,
    cur_italic: bool,
    cur_bold: bool,
) -> tuple[bool, bool]:
    italic = cur_italic or tag in ("i", "em")
    bold = cur_bold or tag in ("b", "strong")

    if inline_style:
        style_lower = inline_style.lower()
        if "font-style" in style_lower:
            if re.search(r"font-style\s*:\s*(?:italic|oblique)", style_lower):
                italic = True
            elif re.search(r"font-style\s*:\s*normal", style_lower):
                italic = False
        if "font-weight" in style_lower:
            if re.search(r"font-weight\s*:\s*(?:bold|bolder|[6-9]\d{2})", style_lower):
                bold = True
            elif re.search(r"font-weight\s*:\s*normal", style_lower):
                bold = False

    return italic, bold


def _trim_leading_whitespace(spans: list[TextSpan]) -> list[TextSpan]:
    result = list(spans)
    while result:
        trimmed = result[0].text.lstrip()
        if trimmed:
            result[0] = TextSpan(
                text=trimmed,
                italic=result[0].italic,
                bold=result[0].bold,
            )
            break
        result.pop(0)
    return result


def _trim_trailing_whitespace(spans: list[TextSpan]) -> list[TextSpan]:
    result = list(spans)
    while result:
        trimmed = result[-1].text.rstrip()
        if trimmed:
            result[-1] = TextSpan(
                text=trimmed,
                italic=result[-1].italic,
                bold=result[-1].bold,
            )
            break
        result.pop()
    return result


def _strip_prefix_chars(spans: list[TextSpan], count: int) -> list[TextSpan]:
    """Usuwa dokładnie `count` początkowych znaków z listy `TextSpan`."""
    remaining = count
    result: list[TextSpan] = []
    for span in spans:
        if remaining <= 0:
            result.append(span)
            continue
        if len(span.text) <= remaining:
            remaining -= len(span.text)
            continue
        result.append(
            TextSpan(
                text=span.text[remaining:],
                italic=span.italic,
                bold=span.bold,
            )
        )
        remaining = 0
    return result


def _merge_spans(spans: list[TextSpan]) -> tuple[TextSpan, ...]:
    merged: list[TextSpan] = []
    for span in spans:
        if not span.text:
            continue
        if merged and merged[-1].italic == span.italic and merged[-1].bold == span.bold:
            prev = merged[-1]
            merged[-1] = TextSpan(
                text=prev.text + span.text,
                italic=prev.italic,
                bold=prev.bold,
            )
        else:
            merged.append(span)
    return tuple(merged)


def parse_csl_bib_html(html_text: str) -> tuple[TextSpan, ...]:
    """Konwertuje HTML wpisu bibliograficznego CSL na krotkę `TextSpan`.

    Zachowuje kursywę i pogrubienie, dekoduje encje HTML, zamienia niełamiące
    spacje na zwykłe oraz usuwa prefiks numeryczny (np. `[1] ` lub blok
    `csl-left-margin`).
    """
    if not html_text or not html_text.strip():
        return ()

    parser = _CslHtmlParser()
    parser.feed(html_text)
    parser.close()

    spans = [s for s in parser.spans if s.text]
    spans = _trim_leading_whitespace(spans)
    spans = _trim_trailing_whitespace(spans)
    if not spans:
        return ()

    full_text = spans_to_plain_text(spans)
    match = _LEADING_NUMERIC_PREFIX_RE.match(full_text)
    if match:
        spans = _strip_prefix_chars(spans, match.end())
        spans = _trim_leading_whitespace(spans)

    return _merge_spans(spans)


def parse_csl_citation_html(html_text: str) -> str:
    """Konwertuje HTML cytowania w tekście CSL (np. `[1]`, `(Dove et al., 2013)`) na czysty tekst."""
    if not html_text or not html_text.strip():
        return ""

    parser = _CslHtmlParser()
    parser.feed(html_text)
    parser.close()

    spans = _trim_leading_whitespace([s for s in parser.spans if s.text])
    spans = _trim_trailing_whitespace(spans)
    return spans_to_plain_text(spans)
