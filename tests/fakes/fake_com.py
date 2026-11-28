"""Atrapy obiektów COM PowerPointa (TextRange z podzakresami) do testów bez PowerPointa."""

from __future__ import annotations


class MockFont:
    """Atrapa obiektu `Font` na `TextRange` i `MockSubRange`."""

    def __init__(self) -> None:
        self.Name: str = "Calibri"
        self.Size: int = 12
        self.Italic: int = 0
        self.Bold: int = 0


class MockTextRange:
    """Atrapa `TextRange` z podzakresami `Characters(start, length)` zapisującymi do rodzica.

    Liczy przypisania całego tekstu (`whole_text_assignments`), żeby testy mogły sprawdzić,
    że wstawianie i podmiana cytowań nie nadpisują całego pola (utrata formatowania – T10).
    """

    def __init__(self, text: str = "", start: int = 1) -> None:
        self._text = text
        self.Start = start
        self.whole_text_assignments = 0
        self.edits: list[tuple[int, int, str]] = []
        self.selected_at: int | None = None
        self.Font = MockFont()
        self.subrange_formats: list[tuple[int, int, int, int]] = []

    @property
    def Text(self) -> str:
        return self._text

    @Text.setter
    def Text(self, value: str) -> None:
        self._text = value
        self.whole_text_assignments += 1

    def Characters(self, start: int, length: int = 0) -> MockSubRange:
        return MockSubRange(self, start, length)

    def Select(self) -> None:
        self.selected_at = self.Start


class _SubRangeFont:
    def __init__(self, subrange: MockSubRange) -> None:
        self._sub = subrange
        self.Name: str = "Calibri"
        self.Size: int = 12
        self._italic: int = 0
        self._bold: int = 0

    def _record(self) -> None:
        self._sub._parent.subrange_formats.append(
            (self._sub.Start, self._sub.Length, self._italic, self._bold)
        )

    @property
    def Italic(self) -> int:
        return self._italic

    @Italic.setter
    def Italic(self, val: int) -> None:
        self._italic = val
        self._record()

    @property
    def Bold(self) -> int:
        return self._bold

    @Bold.setter
    def Bold(self, val: int) -> None:
        self._bold = val
        self._record()


class MockSubRange:
    """Podzakres `TextRange` (1-based `start`, `length`) – zmiany trafiają do tekstu rodzica."""

    def __init__(self, parent: MockTextRange, start: int, length: int) -> None:
        if start < 1 or length < 0 or start - 1 + length > len(parent._text):
            raise ValueError(
                f"Characters({start}, {length}) poza tekstem o długości {len(parent._text)}"
            )
        self._parent = parent
        self.Start = start
        self.Length = length
        self.Font = _SubRangeFont(self)

    @property
    def Text(self) -> str:
        p = self._parent._text
        return p[self.Start - 1 : self.Start - 1 + self.Length]

    @Text.setter
    def Text(self, value: str) -> None:
        p = self._parent._text
        begin = self.Start - 1
        self._parent._text = p[:begin] + value + p[begin + self.Length :]
        self._parent.edits.append((begin, self.Length, value))
        self.Length = len(value)

    def InsertAfter(self, value: str) -> MockSubRange:
        p = self._parent._text
        pos = self.Start - 1 + self.Length
        self._parent._text = p[:pos] + value + p[pos:]
        self._parent.edits.append((pos, 0, value))
        return MockSubRange(self._parent, pos + 1, len(value))

    def Select(self) -> None:
        self._parent.selected_at = self.Start


class MockTags:
    """Atrapa kolekcji `Tags` (prezentacja, slajd, kształt)."""

    def __init__(self) -> None:
        self._d: dict[str, str] = {}

    def __call__(self, key: str) -> str:
        return self._d.get(key, "")

    def Add(self, key: str, val: str) -> None:
        self._d[key] = val

    def Delete(self, key: str) -> None:
        self._d.pop(key, None)


class _MockTextFrame:
    def __init__(self, text: str) -> None:
        self.TextRange = MockTextRange(text)
        self.WordWrap: int = -1
        self.MarginLeft: float = 0.0
        self.MarginRight: float = 0.0
        self.MarginTop: float = 0.0
        self.MarginBottom: float = 0.0
        self.AutoSize: int = 0


class MockShape:
    """Atrapa kształtu z polem tekstowym (lub bez pola tekstowego, gdy `has_text_frame=0`)."""

    def __init__(
        self,
        name: str,
        text: str = "",
        *,
        has_text_frame: int = -1,
        left: float = 0.0,
        top: float = 0.0,
        width: float = 100.0,
        height: float = 36.0,
    ) -> None:
        self.Name = name
        self.HasTextFrame = has_text_frame
        self.TextFrame = _MockTextFrame(text)
        self.Tags = MockTags()
        self.Left = left
        self.Top = top
        self.Width = width
        self.Height = height
        self.Visible = -1
        self.Parent: MockSlide | None = None

    def Delete(self) -> None:
        if self.Parent is not None and self in self.Parent._shapes:
            self.Parent._shapes.remove(self)


class _MockShapes:
    def __init__(self, items: list[MockShape], slide: MockSlide | None = None) -> None:
        self._items = items
        self._slide = slide
        self.Count = len(items)

    def __call__(self, key: int | str) -> MockShape:
        if isinstance(key, int):
            return self._items[key - 1]
        for sh in self._items:
            if sh.Name == key:
                return sh
        raise KeyError(key)

    def AddTextbox(
        self,
        orientation: int,
        left: float,
        top: float,
        width: float,
        height: float,
    ) -> MockShape:
        _ = orientation
        shape = MockShape(
            f"TextBox_{len(self._items) + 1}",
            "",
            left=float(left),
            top=float(top),
            width=float(width),
            height=float(height),
        )
        shape.Parent = self._slide
        self._items.append(shape)
        self.Count = len(self._items)
        return shape


class MockSlide:
    """Atrapa slajdu z listą kształtów."""

    def __init__(self, shapes: list[MockShape], slide_id: int = 256) -> None:
        self.SlideIndex = 1
        self.SlideID = slide_id
        self._shapes = shapes
        for sh in self._shapes:
            sh.Parent = self
        self.Tags = MockTags()
        self._pres: MockPresentation | None = None

    @property
    def Shapes(self) -> _MockShapes:
        return _MockShapes(self._shapes, slide=self)

    def Delete(self) -> None:
        if self._pres is not None and self in self._pres._slides:
            self._pres._slides.remove(self)
            self._pres._reindex_slides()


class _MockSlides:
    def __init__(self, slides: list[MockSlide], pres: MockPresentation | None = None) -> None:
        self._slides = slides
        self._pres = pres
        self.Count = len(slides)

    def __call__(self, idx: int) -> MockSlide:
        return self._slides[idx - 1]

    def Add(self, index: int, layout: int = 12) -> MockSlide:
        _ = layout
        next_id = max((s.SlideID for s in self._slides), default=255) + 1
        new_slide = MockSlide([], slide_id=next_id)
        new_slide._pres = self._pres
        insert_pos = max(0, min(int(index) - 1, len(self._slides)))
        self._slides.insert(insert_pos, new_slide)
        if self._pres is not None:
            self._pres._reindex_slides()
        self.Count = len(self._slides)
        return new_slide


class _MockPageSetup:
    def __init__(self, width: float = 960.0, height: float = 540.0) -> None:
        self.SlideWidth = width
        self.SlideHeight = height


class MockPresentation:
    """Atrapa prezentacji: slajdy, tagi, ścieżka i `SaveCopyAs` / `Save` z rejestrem wywołań."""

    def __init__(
        self,
        slides: list[MockSlide],
        full_name: str = "Prezentacja1",
        path: str = "",
        slide_width: float = 960.0,
        slide_height: float = 540.0,
    ) -> None:
        self._slides = slides
        self._reindex_slides()
        self.FullName = full_name
        self.Path = path
        self.Name = full_name.replace("\\", "/").rsplit("/", 1)[-1]
        self.Tags = MockTags()
        self.PageSetup = _MockPageSetup(slide_width, slide_height)
        self.save_calls = 0
        self.save_copy_as_calls: list[str] = []
        self.save_copy_as_error: Exception | None = None

    def _reindex_slides(self) -> None:
        for idx, s in enumerate(self._slides, start=1):
            s.SlideIndex = idx
            s._pres = self

    @property
    def Slides(self) -> _MockSlides:
        return _MockSlides(self._slides, pres=self)

    def Save(self) -> None:
        self.save_calls += 1

    def SaveCopyAs(self, filename: str) -> None:
        if self.save_copy_as_error is not None:
            raise self.save_copy_as_error
        self.save_copy_as_calls.append(filename)
