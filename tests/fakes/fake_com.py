"""Atrapy obiektów COM PowerPointa (TextRange z podzakresami) do testów bez PowerPointa."""

from __future__ import annotations


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


class MockShape:
    """Atrapa kształtu z polem tekstowym."""

    def __init__(self, name: str, text: str) -> None:
        self.Name = name
        self.HasTextFrame = -1
        self.TextFrame = _MockTextFrame(text)
        self.Tags = MockTags()


class _MockShapes:
    def __init__(self, items: list[MockShape]) -> None:
        self._items = items
        self.Count = len(items)

    def __call__(self, key: int | str) -> MockShape:
        if isinstance(key, int):
            return self._items[key - 1]
        for sh in self._items:
            if sh.Name == key:
                return sh
        raise KeyError(key)


class MockSlide:
    """Atrapa slajdu z listą kształtów."""

    def __init__(self, shapes: list[MockShape], slide_id: int = 256) -> None:
        self.SlideIndex = 1
        self.SlideID = slide_id
        self._shapes = shapes
        self.Tags = MockTags()

    @property
    def Shapes(self) -> _MockShapes:
        return _MockShapes(self._shapes)


class _MockSlides:
    def __init__(self, slides: list[MockSlide]) -> None:
        self._slides = slides
        self.Count = len(slides)

    def __call__(self, idx: int) -> MockSlide:
        return self._slides[idx - 1]


class MockPresentation:
    """Atrapa prezentacji: slajdy, tagi, ścieżka i `SaveCopyAs` / `Save` z rejestrem wywołań."""

    def __init__(
        self,
        slides: list[MockSlide],
        full_name: str = "Prezentacja1",
        path: str = "",
    ) -> None:
        self._slides = slides
        self.FullName = full_name
        self.Path = path
        self.Name = full_name.replace("\\", "/").rsplit("/", 1)[-1]
        self.Tags = MockTags()
        self.save_calls = 0
        self.save_copy_as_calls: list[str] = []
        self.save_copy_as_error: Exception | None = None

    @property
    def Slides(self) -> _MockSlides:
        return _MockSlides(self._slides)

    def Save(self) -> None:
        self.save_calls += 1

    def SaveCopyAs(self, filename: str) -> None:
        if self.save_copy_as_error is not None:
            raise self.save_copy_as_error
        self.save_copy_as_calls.append(filename)
