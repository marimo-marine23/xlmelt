"""Excel .xls (legacy format) parser using xlrd."""

from __future__ import annotations

from pathlib import Path
from typing import Any

try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False

from .model import CellInfo, ImageInfo


class XlsParser:
    """Parse legacy .xls files and extract cell information.

    Provides the same interface as ExcelParser for .xls files.
    """

    def __init__(self, file_path: str | Path):
        if not HAS_XLRD:
            raise ImportError(
                "xlrd is required for .xls files. Install it with: pip install xlrd"
            )
        self.file_path = Path(file_path)
        self._wb: Any = None

    def open(self) -> None:
        """Open the workbook."""
        self._wb = xlrd.open_workbook(
            str(self.file_path), formatting_info=True
        )

    def close(self) -> None:
        """Close the workbook."""
        if self._wb:
            self._wb.release_resources()
            self._wb = None

    def __enter__(self) -> XlsParser:
        self.open()
        return self

    def __exit__(self, *args: Any) -> None:
        self.close()

    @property
    def wb(self) -> Any:
        if self._wb is None:
            raise RuntimeError("Workbook not opened. Call open() first.")
        return self._wb

    @property
    def sheet_names(self) -> list[str]:
        return self.wb.sheet_names()

    def parse_sheet(self, sheet_name: str) -> tuple[list[list[CellInfo | None]], int, int]:
        """Parse a sheet into a 2D grid of CellInfo.

        Returns (grid, row_count, col_count).
        Grid is 1-indexed: grid[row][col], with row/col starting at 1.
        """
        ws = self.wb.sheet_by_name(sheet_name)
        max_row = ws.nrows
        max_col = ws.ncols

        if max_row == 0 or max_col == 0:
            return [], 0, 0

        # Build merge map
        merge_map = self._build_merge_map(ws)

        # Find actual used range
        actual_max_row = 0
        actual_max_col = 0
        for row in range(max_row):
            for col in range(max_col):
                cell = ws.cell(row, col)
                if cell.value not in (None, ""):
                    actual_max_row = max(actual_max_row, row + 1)
                    actual_max_col = max(actual_max_col, col + 1)

        # Account for merges
        for (rlo, rhi, clo, chi) in ws.merged_cells:
            actual_max_row = max(actual_max_row, rhi)
            actual_max_col = max(actual_max_col, chi)

        if actual_max_row == 0:
            return [], 0, 0

        # Build grid (1-indexed)
        grid: list[list[CellInfo | None]] = [
            [None] * (actual_max_col + 1) for _ in range(actual_max_row + 1)
        ]

        for row in range(actual_max_row):
            for col in range(actual_max_col):
                info = self._parse_cell(ws, row, col, merge_map)
                grid[row + 1][col + 1] = info

        return grid, actual_max_row, actual_max_col

    def _build_merge_map(self, ws: Any) -> dict[tuple[int, int], tuple[int, int, int, int]]:
        """Build a map of merged cells (0-indexed internally)."""
        merge_map: dict[tuple[int, int], tuple[int, int, int, int]] = {}
        for rlo, rhi, clo, chi in ws.merged_cells:
            for r in range(rlo, rhi):
                for c in range(clo, chi):
                    merge_map[(r, c)] = (rlo, clo, rhi - 1, chi - 1)
        return merge_map

    def _parse_cell(
        self,
        ws: Any,
        row: int,
        col: int,
        merge_map: dict[tuple[int, int], tuple[int, int, int, int]],
    ) -> CellInfo:
        """Extract CellInfo from an xlrd cell (0-indexed input → 1-indexed output)."""
        cell = ws.cell(row, col)

        merge_info = merge_map.get((row, col))
        is_merged = merge_info is not None
        is_origin = False
        merge_width = 1
        merge_height = 1

        if merge_info:
            min_r, min_c, max_r, max_c = merge_info
            is_origin = row == min_r and col == min_c
            if is_origin:
                merge_width = max_c - min_c + 1
                merge_height = max_r - min_r + 1

        # Get value
        value = None
        if cell.value not in (None, ""):
            value = str(cell.value)

        # Non-origin merged cells have no independent value
        if is_merged and not is_origin:
            value = None

        # Get formatting
        font_size = None
        font_bold = False
        font_color = None
        fill_color = None
        border_top = False
        border_bottom = False
        border_left = False
        border_right = False
        number_format = None
        h_align = None
        v_align = None
        indent = 0

        try:
            xf_index = cell.xf_index
            if xf_index is not None:
                xf = self.wb.xf_list[xf_index]

                # Font
                font = self.wb.font_list[xf.font_index]
                font_size = font.height / 20.0  # twips to points
                font_bold = bool(font.bold)
                if font.colour_index not in (None, 0x7FFF, 64):
                    font_color = self._color_index_to_hex(font.colour_index)

                # Fill
                bg_color_idx = xf.background.pattern_colour_index
                if bg_color_idx not in (None, 0x7FFF, 64, 0):
                    fill_color = self._color_index_to_hex(bg_color_idx)

                # Borders
                border = xf.border
                border_top = border.top_line_style > 0
                border_bottom = border.bottom_line_style > 0
                border_left = border.left_line_style > 0
                border_right = border.right_line_style > 0

                # Number format
                fmt_key = xf.format_key
                fmt_map = self.wb.format_map
                if fmt_key in fmt_map:
                    number_format = fmt_map[fmt_key].format_str

                # Alignment
                align = xf.alignment
                h_align_map = {0: None, 1: "left", 2: "center", 3: "right"}
                v_align_map = {0: "top", 1: "center", 2: "bottom"}
                h_align = h_align_map.get(align.hor_align)
                v_align = v_align_map.get(align.vert_align)
                indent = align.indent
        except (AttributeError, IndexError, KeyError):
            pass

        return CellInfo(
            row=row + 1,
            col=col + 1,
            value=value,
            font_size=font_size,
            font_bold=font_bold,
            font_color=font_color,
            fill_color=fill_color,
            border_top=border_top,
            border_bottom=border_bottom,
            border_left=border_left,
            border_right=border_right,
            merge_width=merge_width,
            merge_height=merge_height,
            is_merged_origin=is_origin,
            is_merged_cell=is_merged,
            number_format=number_format,
            alignment_horizontal=h_align,
            alignment_vertical=v_align,
            indent=indent,
        )

    def _color_index_to_hex(self, color_index: int) -> str | None:
        """Convert xlrd color index to hex color string."""
        try:
            color_map = self.wb.colour_map
            if color_index in color_map:
                rgb = color_map[color_index]
                if rgb:
                    return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
        except (AttributeError, KeyError):
            pass
        return None

    def extract_images(self, output_dir: Path) -> list[ImageInfo]:
        """Extract images from .xls files (limited support via OLE)."""
        # xlrd doesn't provide a convenient image extraction API
        # Images in .xls are stored in OLE compound documents
        # For now, return empty list — full support would require olefile
        return []

    def get_column_widths(self, sheet_name: str) -> list[float]:
        """Get column widths."""
        ws = self.wb.sheet_by_name(sheet_name)
        widths: list[float] = []
        for col in range(ws.ncols):
            try:
                # xlrd returns width in 1/256th of character width
                width = ws.computed_column_width(col) / 256.0
                widths.append(width if width > 0 else 8.43)
            except (AttributeError, Exception):
                widths.append(8.43)
        return widths
