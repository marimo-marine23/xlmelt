"""Semantic model for representing Excel document structure."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


def _col_to_letter(col: int) -> str:
    """Convert 1-based column number to Excel letter (1→A, 27→AA)."""
    if col < 1:
        return "?"
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _region_to_a1(r: "Region") -> str:
    """Convert Region to A1-style range string (e.g., 'A1:E10')."""
    start = f"{_col_to_letter(r.min_col)}{r.min_row}"
    end = f"{_col_to_letter(r.max_col)}{r.max_row}"
    return f"{start}:{end}"


class SectionType(Enum):
    """Type of document section."""

    HEADING = "heading"
    TABLE = "table"
    KEY_VALUE = "key_value"
    TEXT = "text"
    LIST = "list"
    IMAGE = "image"
    CHART = "chart"


@dataclass
class CellInfo:
    """Raw cell information from Excel."""

    row: int
    col: int
    value: Any
    font_size: float | None = None
    font_bold: bool = False
    font_color: str | None = None
    fill_color: str | None = None
    border_top: bool = False
    border_bottom: bool = False
    border_left: bool = False
    border_right: bool = False
    merge_width: int = 1
    merge_height: int = 1
    is_merged_origin: bool = False
    is_merged_cell: bool = False
    number_format: str | None = None
    alignment_horizontal: str | None = None
    alignment_vertical: str | None = None
    indent: int = 0


@dataclass
class Region:
    """A rectangular region of cells in a sheet."""

    min_row: int
    min_col: int
    max_row: int
    max_col: int

    @property
    def width(self) -> int:
        return self.max_col - self.min_col + 1

    @property
    def height(self) -> int:
        return self.max_row - self.min_row + 1

    def contains(self, row: int, col: int) -> bool:
        return (
            self.min_row <= row <= self.max_row
            and self.min_col <= col <= self.max_col
        )

    def __repr__(self) -> str:
        return f"Region({self.min_row},{self.min_col}:{self.max_row},{self.max_col})"


@dataclass
class Section:
    """A semantic section of the document."""

    type: SectionType
    level: int = 0
    title: str | None = None
    content: Any = None
    children: list[Section] = field(default_factory=list)
    source_region: Region | None = None

    def to_dict(self) -> dict:
        """Convert to dictionary for JSON serialization."""
        result: dict[str, Any] = {"type": self.type.value}
        if self.level:
            result["level"] = self.level
        if self.title is not None:
            result["title"] = self.title
        if self.content is not None:
            result["content"] = self.content
        if self.children:
            result["children"] = [c.to_dict() for c in self.children]
        if self.source_region:
            r = self.source_region
            result["source_range"] = f"R{r.min_row}C{r.min_col}:R{r.max_row}C{r.max_col}"
            result["source_range_a1"] = _region_to_a1(r)
        return result


@dataclass
class SheetModel:
    """Semantic model for a single sheet."""

    name: str
    sections: list[Section] = field(default_factory=list)
    row_count: int = 0
    col_count: int = 0

    def to_dict(self) -> dict:
        result: dict[str, Any] = {
            "name": self.name,
            "sections": [s.to_dict() for s in self.sections],
        }
        # Section type summary for quick overview
        type_counts: dict[str, int] = {}
        for s in self.sections:
            t = s.type.value
            type_counts[t] = type_counts.get(t, 0) + 1
        result["section_summary"] = type_counts
        return result


@dataclass
class ImageInfo:
    """Information about an embedded image."""

    path: str
    format: str
    sheet_name: str
    anchor_cell: str | None = None
    alt_text: str | None = None
    width: int | None = None
    height: int | None = None


@dataclass
class DocumentModel:
    """Semantic model for an entire Excel document."""

    title: str
    source_file: str
    sheets: list[SheetModel] = field(default_factory=list)
    images: list[ImageInfo] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict:
        from .. import __version__

        result: dict[str, Any] = {
            "xlmelt_version": __version__,
            "schema_version": 1,
            "document": {
                "title": self.title,
                "source": self.source_file,
                "sheets": [s.to_dict() for s in self.sheets],
            }
        }
        if self.images:
            result["document"]["images"] = [
                {"path": img.path, "format": img.format, "sheet": img.sheet_name}
                for img in self.images
            ]
        if self.metadata:
            result["document"]["metadata"] = self.metadata
        return result
