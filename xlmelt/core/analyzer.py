"""Structure analyzer - detect document structure from Excel metadata."""

from __future__ import annotations

import re
from pathlib import Path

from .model import (
    CellInfo,
    DocumentModel,
    ImageInfo,
    Region,
    Section,
    SectionType,
    SheetModel,
)
from .parser import ExcelParser

try:
    from .xls_parser import XlsParser
    HAS_XLS_SUPPORT = True
except ImportError:
    HAS_XLS_SUPPORT = False


# Threshold for "large font" heading detection
HEADING_FONT_SIZE_MIN = 14
SUBHEADING_FONT_SIZE_MIN = 12

# Threshold for uniform column width detection (Excel方眼紙)
HOUGANSHI_WIDTH_VARIANCE_MAX = 0.5
HOUGANSHI_MIN_COLS = 10

# List marker patterns for bullet/numbered list detection
_LIST_MARKER_RE = re.compile(
    r"^\s*("
    r"[・●■◆◇▶▷▸▹★☆○►]"           # Bullet markers
    r"|[-–—\u2022\u2023\u25E6]"       # Dash/unicode bullets
    r"|\d{1,3}[.)）]"                  # Numbered: 1. 1) 1）
    r"|[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]"  # Circled numbers
    r"|[ⅰⅱⅲⅳⅴⅵⅶⅷⅸⅹⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]"  # Roman numerals
    r"|[a-zA-Zａ-ｚＡ-Ｚ][.)）]"       # Lettered: a. a) A.
    r"|[(（]\d{1,3}[)）]"              # Parenthesized: (1) （1）
    r"|[(（][a-zA-Zａ-ｚＡ-Ｚ][)）]"   # Parenthesized letters: (a) （a）
    r"|※"                              # Note marker
    r")\s*"
)

# Ordered list marker patterns (numbered, circled, roman, lettered, parenthesized)
_ORDERED_MARKER_RE = re.compile(
    r"^\s*("
    r"\d{1,3}[.)）]"                   # Numbered: 1. 1) 1）
    r"|[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]"  # Circled numbers
    r"|[ⅰⅱⅲⅳⅴⅵⅶⅷⅸⅹⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]"  # Roman numerals
    r"|[a-zA-Zａ-ｚＡ-Ｚ][.)）]"       # Lettered: a. a) A.
    r"|[(（]\d{1,3}[)）]"              # Parenthesized: (1) （1）
    r"|[(（][a-zA-Zａ-ｚＡ-Ｚ][)）]"   # Parenthesized letters: (a) （a）
    r")\s*"
)


def _infer_format_type(number_format: str | None) -> str | None:
    """Infer semantic type from Excel number_format string."""
    if not number_format or number_format == "General":
        return None

    fmt = number_format.lower()

    # Date/time formats
    if any(p in fmt for p in ("yy", "mm/dd", "dd/mm", "yyyy", "m/d", "d/m", "ge.m.d")):
        # Distinguish date vs datetime
        if any(p in fmt for p in ("h:mm", "hh:", "ss")):
            return "datetime"
        return "date"
    if any(p in fmt for p in ("h:mm", "hh:", ":ss")):
        return "time"

    # Percentage
    if "%" in fmt:
        return "percentage"

    # Currency
    if any(p in fmt for p in ("¥", "$", "€", "£", "円", "usd", "jpy", "eur")):
        return "currency"
    # Japanese Yen format patterns like #,##0"円"
    if "円" in number_format:
        return "currency"

    # Number with decimal
    if "." in fmt and any(c.isdigit() or c == "#" or c == "0" for c in fmt):
        return "decimal"

    # Integer format with thousands separator
    if "#,##" in fmt or "0,0" in fmt:
        return "number"

    return None


class StructureAnalyzer:
    """Analyze Excel sheet structure using heuristics."""

    def analyze(self, file_path: str | Path, output_dir: Path | None = None) -> DocumentModel:
        """Analyze an Excel file and return a semantic model."""
        file_path = Path(file_path)

        # Choose parser based on file extension
        if file_path.suffix.lower() == ".xls":
            if not HAS_XLS_SUPPORT:
                raise ImportError(
                    "xlrd is required for .xls files. Install it with: pip install xlrd"
                )
            parser_cls = XlsParser
        else:
            parser_cls = ExcelParser

        with parser_cls(file_path) as parser:
            doc = DocumentModel(
                title=file_path.stem,
                source_file=file_path.name,
            )

            for sheet_name in parser.sheet_names:
                grid, row_count, col_count = parser.parse_sheet(sheet_name)
                if row_count == 0:
                    continue

                widths = parser.get_column_widths(sheet_name)
                is_houganshi = self._detect_houganshi(widths)

                sheet_model = SheetModel(
                    name=sheet_name,
                    row_count=row_count,
                    col_count=col_count,
                )

                if is_houganshi:
                    grid, row_count, col_count = self._preprocess_houganshi(
                        grid, row_count, col_count
                    )

                sections = self._analyze_sheet(grid, row_count, col_count, is_houganshi)
                sheet_model.sections = sections
                doc.sheets.append(sheet_model)

            # Extract images if output dir given
            if output_dir:
                doc.images = parser.extract_images(output_dir)
                # Insert image sections into sheets at anchor positions
                self._insert_image_sections(doc)

        return doc

    def _insert_image_sections(self, doc: DocumentModel) -> None:
        """Insert IMAGE sections into sheets based on anchor cell positions."""
        # Group images by sheet name
        sheet_images: dict[str, list[ImageInfo]] = {}
        unmatched_images: list[ImageInfo] = []
        for img in doc.images:
            if img.sheet_name and img.sheet_name in {s.name for s in doc.sheets}:
                sheet_images.setdefault(img.sheet_name, []).append(img)
            else:
                unmatched_images.append(img)

        for sheet in doc.sheets:
            imgs = sheet_images.get(sheet.name, [])
            if not imgs:
                continue
            self._insert_images_to_sheet(sheet, imgs)

        # Append unmatched images to the first sheet (or last if only one)
        if unmatched_images and doc.sheets:
            target_sheet = doc.sheets[0]
            self._insert_images_to_sheet(target_sheet, unmatched_images)

    def _insert_images_to_sheet(
        self, sheet: SheetModel, imgs: list[ImageInfo]
    ) -> None:
        """Insert image sections into a sheet's section list."""
        positioned: list[tuple[int, ImageInfo]] = []
        unpositioned: list[ImageInfo] = []
        for img in imgs:
            anchor_row = self._parse_anchor_row(img.anchor_cell)
            if anchor_row is not None:
                positioned.append((anchor_row, img))
            else:
                unpositioned.append(img)

        # Insert positioned images into sections list by anchor row
        positioned.sort(key=lambda x: x[0])
        for anchor_row, img in positioned:
            section = Section(
                type=SectionType.IMAGE,
                content={"path": img.path, "alt": img.alt_text or ""},
            )
            insert_idx = len(sheet.sections)
            for i, sec in enumerate(sheet.sections):
                if sec.source_region and sec.source_region.min_row > anchor_row:
                    insert_idx = i
                    break
            sheet.sections.insert(insert_idx, section)

        # Append unpositioned images at the end
        for img in unpositioned:
            sheet.sections.append(Section(
                type=SectionType.IMAGE,
                content={"path": img.path, "alt": img.alt_text or ""},
            ))

    @staticmethod
    def _parse_anchor_row(anchor_cell: str | None) -> int | None:
        """Extract row number from anchor cell reference like 'B15'."""
        if not anchor_cell:
            return None
        m = re.search(r"(\d+)$", anchor_cell)
        return int(m.group(1)) if m else None

    def _detect_houganshi(self, widths: list[float]) -> bool:
        """Detect Excel方眼紙 pattern (uniform narrow columns)."""
        if len(widths) < HOUGANSHI_MIN_COLS:
            return False
        if not widths:
            return False
        avg = sum(widths) / len(widths)
        if avg > 5:  # 方眼紙 typically uses very narrow columns
            return False
        variance = sum((w - avg) ** 2 for w in widths) / len(widths)
        return variance < HOUGANSHI_WIDTH_VARIANCE_MAX

    def _preprocess_houganshi(
        self,
        grid: list[list[CellInfo | None]],
        row_count: int,
        col_count: int,
    ) -> tuple[list[list[CellInfo | None]], int, int]:
        """Pre-process Excel方眼紙 grids by merging adjacent cells.

        In houganshi (grid paper) layout, narrow columns are used as a virtual
        grid, and content is placed by merging cells. Adjacent non-merged cells
        with the same value or empty cells next to valued cells on the same row
        are logically grouped to simulate wider cells.

        This method identifies "logical columns" by analyzing where values
        actually appear and creates virtual merge regions.
        """
        # Step 1: For each row, find contiguous runs of empty cells between
        # valued cells and mark them as belonging to the nearest value cell.
        # This effectively creates logical column boundaries.

        # Step 2: For rows without explicit merges, find value cells and
        # extend their effective width to cover adjacent empty cells until
        # the next value cell or edge.
        for row in range(1, row_count + 1):
            # Skip rows that already have merge info
            row_has_merges = any(
                grid[row][col] and grid[row][col].is_merged_cell
                for col in range(1, col_count + 1)
                if grid[row][col]
            )
            if row_has_merges:
                continue

            # Find value cells in this row
            value_cols: list[int] = []
            for col in range(1, col_count + 1):
                cell = grid[row][col]
                if cell and cell.value is not None and cell.value.strip():
                    value_cols.append(col)

            if not value_cols:
                continue

            # For each value cell, extend merge_width to cover empty cells
            # until the next value cell
            for i, vc in enumerate(value_cols):
                next_vc = value_cols[i + 1] if i + 1 < len(value_cols) else col_count + 1
                # Count contiguous empty cells after this value cell
                extend = 0
                for c in range(vc + 1, next_vc):
                    cell = grid[row][c]
                    if cell and (cell.value is None or not cell.value.strip()):
                        extend += 1
                    else:
                        break

                if extend > 0:
                    cell = grid[row][vc]
                    if cell:
                        cell.merge_width = extend + 1
                        cell.is_merged_origin = True
                        cell.is_merged_cell = True
                        # Mark extended cells as merged non-origin
                        for c in range(vc + 1, vc + extend + 1):
                            ext_cell = grid[row][c]
                            if ext_cell:
                                ext_cell.is_merged_cell = True
                                ext_cell.is_merged_origin = False

        return grid, row_count, col_count

    def _analyze_sheet(
        self,
        grid: list[list[CellInfo | None]],
        row_count: int,
        col_count: int,
        is_houganshi: bool,
    ) -> list[Section]:
        """Analyze a sheet grid and detect sections."""
        # Step 1: Find all non-empty rows and identify their characteristics
        row_infos = self._classify_rows(grid, row_count, col_count)

        # Step 2: Merge table footer rows into table_row
        row_infos = self._merge_table_footers(row_infos, grid, col_count)

        # Step 3: Detect contiguous regions of similar row types
        sections: list[Section] = []
        i = 0
        while i < len(row_infos):
            ri = row_infos[i]

            if ri["type"] == "empty":
                i += 1
                continue

            if ri["type"] == "heading":
                section = self._make_heading(grid, ri)
                sections.append(section)
                i += 1
                continue

            if ri["type"] == "table_row":
                # Scan forward for contiguous table rows
                end = i + 1
                while end < len(row_infos) and row_infos[end]["type"] in ("table_row", "empty_between_table"):
                    end += 1
                # Remove trailing empties
                while end > i + 1 and row_infos[end - 1]["type"] == "empty_between_table":
                    end -= 1
                table_rows = [row_infos[j] for j in range(i, end) if row_infos[j]["type"] == "table_row"]
                table_sections = self._make_tables(grid, table_rows, col_count)
                sections.extend(table_sections)
                i = end
                continue

            if ri["type"] == "kv_row":
                # Scan forward for contiguous key-value rows
                end = i + 1
                while end < len(row_infos) and row_infos[end]["type"] == "kv_row":
                    end += 1
                kv_rows = [row_infos[j] for j in range(i, end)]
                section = self._make_key_value(grid, kv_rows, col_count)
                sections.append(section)
                i = end
                continue

            if ri["type"] == "list_item":
                # Scan forward for contiguous list items
                end = i + 1
                while end < len(row_infos) and row_infos[end]["type"] == "list_item":
                    end += 1
                section = self._make_list(grid, row_infos[i:end], col_count)
                sections.append(section)
                i = end
                continue

            if ri["type"] == "text":
                # Scan forward for contiguous text rows
                end = i + 1
                while end < len(row_infos) and row_infos[end]["type"] == "text":
                    end += 1
                section = self._make_text(grid, row_infos[i:end], col_count)
                sections.append(section)
                i = end
                continue

            # Fallback: treat as text
            section = self._make_text(grid, [ri], col_count)
            sections.append(section)
            i += 1

        return sections

    def _classify_rows(
        self,
        grid: list[list[CellInfo | None]],
        row_count: int,
        col_count: int,
    ) -> list[dict]:
        """Classify each row by its structural role."""
        row_infos: list[dict] = []

        # Pre-compute: count of bordered cells per row for table detection
        for row in range(1, row_count + 1):
            cells = [grid[row][col] for col in range(1, col_count + 1) if grid[row][col]]
            non_empty = [c for c in cells if (c.value is not None and not c.is_merged_cell) or c.is_merged_origin]

            if not non_empty:
                row_infos.append({"type": "empty", "row": row})
                continue

            # Check for heading: single cell with large/bold font, possibly merged
            if self._is_heading_row(non_empty, col_count):
                row_infos.append({"type": "heading", "row": row, "cells": non_empty})
                continue

            # Check for bordered row (table candidate)
            # Only consider cells in the occupied column range (min_col to max_col with data/borders)
            # to avoid dilution by empty leading/trailing columns (common in houganshi)
            occupied_cols = set()
            for c in cells:
                if c and (c.value is not None or self._has_borders(c)):
                    if not c.is_merged_cell or c.is_merged_origin:
                        occupied_cols.add(c.col)
            if occupied_cols:
                min_occ = min(occupied_cols)
                max_occ = max(occupied_cols)
                non_merged_cells = [
                    grid[row][col]
                    for col in range(min_occ, max_occ + 1)
                    if grid[row][col]
                    and (not grid[row][col].is_merged_cell or grid[row][col].is_merged_origin)
                ]
            else:
                non_merged_cells = []
            bordered_count = sum(1 for c in non_merged_cells if self._has_borders(c))
            bordered_ratio = bordered_count / max(len(non_merged_cells), 1)

            if bordered_ratio > 0.5 and len(non_empty) >= 2:
                row_infos.append({"type": "table_row", "row": row, "cells": cells})
                continue

            # Check for key-value pattern: 2-4 cells where odd-indexed are labels
            if self._is_kv_row(non_empty, col_count):
                row_infos.append({"type": "kv_row", "row": row, "cells": non_empty})
                continue

            # Check for list item (bullet/numbered marker)
            if self._is_list_item(non_empty):
                row_infos.append({"type": "list_item", "row": row, "cells": non_empty})
                continue

            # Default: text
            row_infos.append({"type": "text", "row": row, "cells": non_empty})

        # Post-process: mark empty rows between table rows as "empty_between_table"
        for idx in range(len(row_infos)):
            if row_infos[idx]["type"] == "empty":
                before = idx > 0 and row_infos[idx - 1]["type"] == "table_row"
                after = idx < len(row_infos) - 1 and row_infos[idx + 1]["type"] == "table_row"
                if before and after:
                    row_infos[idx]["type"] = "empty_between_table"

        # Post-process: detect borderless tables in runs of "text" rows
        row_infos = self._detect_borderless_tables(row_infos, grid, col_count)

        return row_infos

    def _detect_borderless_tables(
        self,
        row_infos: list[dict],
        grid: list[list[CellInfo | None]],
        col_count: int,
    ) -> list[dict]:
        """Detect borderless tables: runs of text rows with consistent column alignment."""
        result = list(row_infos)  # shallow copy
        i = 0
        while i < len(result):
            if result[i]["type"] != "text":
                i += 1
                continue

            # Find run of consecutive text rows
            end = i + 1
            while end < len(result) and result[end]["type"] == "text":
                end += 1

            run_length = end - i
            if run_length >= 3:  # Need at least 3 rows for a borderless table
                if self._is_borderless_table(result[i:end], grid, col_count):
                    for j in range(i, end):
                        result[j]["type"] = "table_row"
            i = end

        return result

    def _is_borderless_table(
        self,
        text_rows: list[dict],
        grid: list[list[CellInfo | None]],
        col_count: int,
    ) -> bool:
        """Check if a group of text rows forms a borderless table.

        Criteria:
        - Each row has ≥2 value cells
        - Most rows have the same number of value cells (±1)
        - Value cells occupy consistent column positions across rows
        - At least some column has number/date data type consistency
        """
        if len(text_rows) < 3:
            return False

        # Collect column position sets and cell counts per row
        col_sets: list[set[int]] = []
        cell_counts: list[int] = []
        for ri in text_rows:
            row_num = ri["row"]
            cols_used: set[int] = set()
            count = 0
            for col in range(1, col_count + 1):
                cell = grid[row_num][col] if row_num < len(grid) and col < len(grid[row_num]) else None
                if cell and cell.value is not None and cell.value.strip():
                    if not cell.is_merged_cell or cell.is_merged_origin:
                        cols_used.add(col)
                        count += 1
            col_sets.append(cols_used)
            cell_counts.append(count)

        # All rows must have ≥2 cells
        if any(c < 2 for c in cell_counts):
            return False

        # Check column count consistency: most common count should cover ≥60% of rows
        from collections import Counter
        count_freq = Counter(cell_counts)
        most_common_count, most_common_freq = count_freq.most_common(1)[0]
        if most_common_freq / len(text_rows) < 0.6:
            return False

        # Check column position overlap: union/intersection ratio
        all_cols = set()
        for s in col_sets:
            all_cols |= s
        if not all_cols:
            return False

        # Count how many rows use each column
        col_usage: dict[int, int] = {}
        for s in col_sets:
            for c in s:
                col_usage[c] = col_usage.get(c, 0) + 1

        # At least half the columns should be used in ≥60% of rows
        consistent_cols = sum(1 for c, n in col_usage.items() if n >= len(text_rows) * 0.6)
        if consistent_cols < 2:
            return False

        # Check data type consistency in at least one column:
        # If a column has ≥50% numeric values (across non-header rows), it's likely tabular
        has_numeric_col = False
        for col in all_cols:
            numeric_count = 0
            total = 0
            for ri in text_rows[1:]:  # skip potential header row
                row_num = ri["row"]
                cell = grid[row_num][col] if row_num < len(grid) and col < len(grid[row_num]) else None
                if cell and cell.value is not None and cell.value.strip():
                    total += 1
                    val = cell.value.strip()
                    # Check if value looks numeric (including formatted numbers)
                    cleaned = val.replace(",", "").replace("¥", "").replace("$", "").replace("%", "").replace("円", "")
                    try:
                        float(cleaned)
                        numeric_count += 1
                    except ValueError:
                        pass
            if total >= 2 and numeric_count / total >= 0.5:
                has_numeric_col = True
                break

        # First row all text (potential header) + numeric column = strong signal
        if has_numeric_col:
            return True

        # Without numeric column, require very high column consistency (≥80%)
        if consistent_cols >= len(all_cols) * 0.8 and most_common_count >= 3:
            return True

        return False

    def _merge_table_footers(
        self,
        row_infos: list[dict],
        grid: list[list[CellInfo | None]],
        col_count: int,
    ) -> list[dict]:
        """Merge footer rows (合計, 小計, etc.) into the preceding table."""
        _FOOTER_KEYWORDS = {"合計", "計", "小計", "総計", "Total", "total", "TOTAL",
                            "Subtotal", "subtotal", "Sum", "sum", "平均", "Average"}

        result = list(row_infos)
        i = 0
        while i < len(result):
            if result[i]["type"] != "table_row":
                i += 1
                continue

            # Find end of table_row run
            end = i + 1
            while end < len(result) and result[end]["type"] in ("table_row", "empty_between_table"):
                end += 1

            # Check rows immediately after the table (skip one empty row at most)
            check_start = end
            if check_start < len(result) and result[check_start]["type"] == "empty":
                check_start += 1  # Allow one empty row gap

            # Check if the next non-empty row is a footer
            while check_start < len(result) and check_start <= end + 2:
                ri = result[check_start]
                if ri["type"] not in ("text", "kv_row"):
                    break
                # Check if any cell contains a footer keyword
                cells = ri.get("cells", [])
                has_footer_keyword = False
                for c in cells:
                    if c.value and any(kw in str(c.value) for kw in _FOOTER_KEYWORDS):
                        has_footer_keyword = True
                        break

                # Also check if the row has borders matching the table
                row_num = ri["row"]
                non_merged_cells = [
                    grid[row_num][col]
                    for col in range(1, col_count + 1)
                    if grid[row_num][col]
                    and (not grid[row_num][col].is_merged_cell or grid[row_num][col].is_merged_origin)
                ]
                bordered_count = sum(1 for c in non_merged_cells if self._has_borders(c))

                if has_footer_keyword or bordered_count > 0:
                    result[check_start]["type"] = "table_row"
                    # Also convert any skipped empty row
                    for skip_idx in range(end, check_start):
                        if result[skip_idx]["type"] == "empty":
                            result[skip_idx]["type"] = "empty_between_table"
                    check_start += 1
                else:
                    break

            i = max(end, check_start)

        return result

    def _is_heading_row(self, non_empty: list[CellInfo], col_count: int) -> bool:
        """Check if a row is likely a heading."""
        if not non_empty:
            return False

        # Single value cell (possibly merged across columns)
        value_cells = [c for c in non_empty if c.value and c.value.strip()]
        if len(value_cells) != 1:
            return False

        cell = value_cells[0]

        # Large font
        if cell.font_size and cell.font_size >= HEADING_FONT_SIZE_MIN:
            return True

        # Bold + merged across significant portion of the row
        if cell.font_bold and cell.merge_width >= max(col_count // 3, 2):
            return True

        # Bold + font slightly larger than default
        if cell.font_bold and cell.font_size and cell.font_size >= SUBHEADING_FONT_SIZE_MIN:
            return True

        # Centered bold text spanning the row
        if cell.font_bold and cell.alignment_horizontal == "center":
            return True

        # Bold-only single cell with short text (typical section heading)
        # Conservative: avoid false positives on bold data cells in tables
        if cell.font_bold and isinstance(cell.value, str) and len(cell.value.strip()) <= 40:
            return True

        return False

    def _has_borders(self, cell: CellInfo) -> bool:
        """Check if a cell has any borders."""
        return cell.border_top or cell.border_bottom or cell.border_left or cell.border_right

    @staticmethod
    def _cell_style(cell: CellInfo) -> dict[str, str] | None:
        """Extract visual style info from a cell (fill_color, font_color, font_bold)."""
        style: dict[str, str] = {}
        if cell.fill_color and cell.fill_color not in ("000000", "FFFFFF", "00000000"):
            style["bg"] = cell.fill_color
        if cell.font_color and cell.font_color not in ("000000", "00000000"):
            style["color"] = cell.font_color
        if cell.font_bold:
            style["bold"] = "true"
        return style if style else None

    def _is_kv_row(self, non_empty: list[CellInfo], col_count: int) -> bool:
        """Check if row looks like key-value pairs."""
        if len(non_empty) < 2 or len(non_empty) > 6:
            return False

        # Check for alternating bold/non-bold pattern (label: value)
        has_bold = any(c.font_bold for c in non_empty)
        has_non_bold = any(not c.font_bold for c in non_empty)
        if has_bold and has_non_bold:
            return True

        # Check for fill-color pattern (colored label cells)
        has_fill = any(c.fill_color for c in non_empty)
        has_no_fill = any(not c.fill_color for c in non_empty)
        if has_fill and has_no_fill and len(non_empty) <= 4:
            return True

        return False

    def _is_list_item(self, non_empty: list[CellInfo]) -> bool:
        """Check if a row is a list item (starts with a list marker)."""
        if not non_empty:
            return False
        # Get the first cell with text content
        value_cells = [c for c in non_empty if c.value and c.value.strip()]
        if not value_cells:
            return False
        # Only consider rows with 1-2 value cells (list items are typically single-cell)
        if len(value_cells) > 2:
            return False
        first_val = value_cells[0].value.strip()
        return bool(_LIST_MARKER_RE.match(first_val))

    def _make_list(
        self,
        grid: list[list[CellInfo | None]],
        list_rows: list[dict],
        col_count: int,
    ) -> Section:
        """Create a list section from contiguous list item rows.

        Strips list markers from text and determines if the list is ordered
        or unordered. Content is a dict: {"ordered": bool, "items": [str]}
        """
        items: list[str] = []
        ordered_count = 0
        unordered_count = 0

        for ri in list_rows:
            cells = ri["cells"]
            values = [c.value.strip() for c in cells if c.value and c.value.strip()]
            if values:
                full_text = "  ".join(values)
                # Split cell-internal newlines into separate lines
                lines = self._split_cell_list_lines(full_text)
                for line in lines:
                    if not line.strip():
                        continue
                    m = _LIST_MARKER_RE.match(line)
                    if m:
                        marker_text = m.group(1).strip()
                        if marker_text == "※":
                            items.append(line)
                            unordered_count += 1
                        else:
                            stripped = line[m.end():].strip()
                            if _ORDERED_MARKER_RE.match(line):
                                ordered_count += 1
                            else:
                                unordered_count += 1
                            items.append(stripped if stripped else line)
                    else:
                        # Non-marker line: append to previous item as continuation
                        if items:
                            items[-1] += "\n" + line.strip()
                        else:
                            items.append(line.strip())

        # Determine list type: ordered if majority of items are numbered
        is_ordered = ordered_count > unordered_count

        first_row = list_rows[0]["row"]
        last_row = list_rows[-1]["row"]

        return Section(
            type=SectionType.LIST,
            content={"ordered": is_ordered, "items": items},
            source_region=Region(
                min_row=first_row,
                min_col=1,
                max_row=last_row,
                max_col=col_count,
            ),
        )

    @staticmethod
    def _split_cell_list_lines(text: str) -> list[str]:
        """Split cell text on newlines when it contains list markers.

        If the text contains \\n and at least 2 lines start with a list marker,
        split into separate lines. Otherwise return as single item.
        """
        if "\n" not in text:
            return [text]
        lines = text.split("\n")
        marker_count = sum(1 for line in lines if _LIST_MARKER_RE.match(line.strip()))
        if marker_count >= 2:
            return lines
        return [text]

    def _make_heading(self, grid: list[list[CellInfo | None]], row_info: dict) -> Section:
        """Create a heading section."""
        cells = row_info["cells"]
        value_cells = [c for c in cells if c.value and c.value.strip()]
        cell = value_cells[0]

        # Determine heading level based on font size
        level = 2  # default
        if cell.font_size:
            if cell.font_size >= 18:
                level = 1
            elif cell.font_size >= 14:
                level = 2
            elif cell.font_size >= 12:
                level = 3
            else:
                level = 4
        elif cell.font_bold:
            level = 3

        return Section(
            type=SectionType.HEADING,
            level=level,
            title=cell.value.strip() if cell.value else "",
            source_region=Region(
                min_row=cell.row,
                min_col=cell.col,
                max_row=cell.row + cell.merge_height - 1,
                max_col=cell.col + cell.merge_width - 1,
            ),
        )

    def _make_tables(
        self,
        grid: list[list[CellInfo | None]],
        table_rows: list[dict],
        col_count: int,
    ) -> list[Section]:
        """Create table section(s), splitting on empty column gaps."""
        if not table_rows:
            return []

        # Find occupied columns across all table rows
        occupied: set[int] = set()
        for ri in table_rows:
            for c in ri["cells"]:
                if c and c.value is not None:
                    for cc in range(c.col, c.col + c.merge_width):
                        occupied.add(cc)

        if not occupied:
            return []

        # Identify contiguous column groups separated by empty columns
        sorted_cols = sorted(occupied)
        col_groups: list[tuple[int, int]] = []  # (min_col, max_col) per group
        group_start = sorted_cols[0]
        prev = sorted_cols[0]
        for c in sorted_cols[1:]:
            if c > prev + 1:
                col_groups.append((group_start, prev))
                group_start = c
            prev = c
        col_groups.append((group_start, prev))

        # Generate a layout_group id if there are multiple side-by-side tables
        first_row = table_rows[0]["row"]
        last_row = table_rows[-1]["row"]
        layout_group = f"R{first_row}_{last_row}" if len(col_groups) > 1 else None

        sections: list[Section] = []
        for min_col, max_col in col_groups:
            section = self._make_single_table(
                grid, table_rows, min_col, max_col, layout_group,
            )
            if section:
                sections.append(section)
        return sections

    def _detect_header_count(
        self,
        grid: list[list[CellInfo | None]],
        table_rows: list[dict],
        min_col: int,
        max_col: int,
    ) -> int:
        """Detect number of header rows using multiple heuristics."""
        header_count = 0
        for idx, ri in enumerate(table_rows):
            r = ri["row"]
            origin_cells = [
                grid[r][col]
                for col in range(min_col, max_col + 1)
                if grid[r][col]
                and (not grid[r][col].is_merged_cell or grid[r][col].is_merged_origin)
                and grid[r][col].value is not None
            ]
            if not origin_cells:
                break

            # Heuristic 1: All cells are bold or have fill color
            all_styled = all(c.font_bold or c.fill_color for c in origin_cells)

            # Heuristic 2: Row has colspan merges (typical of multi-level headers)
            has_header_merge = any(c.merge_width > 1 for c in origin_cells)

            # Heuristic 3: Next row has rowspan merges starting (typical of data,
            # not headers) — if this row doesn't but next does, this is last header
            next_has_rowspan = False
            if idx + 1 < len(table_rows):
                nr = table_rows[idx + 1]["row"]
                next_origin = [
                    grid[nr][col]
                    for col in range(min_col, max_col + 1)
                    if grid[nr][col]
                    and (not grid[nr][col].is_merged_cell or grid[nr][col].is_merged_origin)
                    and grid[nr][col].value is not None
                ]
                next_has_rowspan = any(c.merge_height > 1 for c in next_origin)

            # Heuristic 4: Data type consistency — header cells are typically text,
            # while data cells are numbers. Check if this row is all text while
            # next row has numbers.
            all_text = all(
                isinstance(c.value, str) or c.value is None
                for c in origin_cells
            )
            next_has_numbers = False
            if idx + 1 < len(table_rows):
                nr = table_rows[idx + 1]["row"]
                for col in range(min_col, max_col + 1):
                    nc = grid[nr][col]
                    if nc and nc.value is not None:
                        try:
                            float(nc.value)
                            next_has_numbers = True
                            break
                        except (ValueError, TypeError):
                            pass

            if all_styled:
                header_count += 1
            elif has_header_merge and all_text:
                header_count += 1
            elif all_text and next_has_numbers and idx == 0:
                # First row is all text and next has numbers — likely header
                header_count += 1
            else:
                break

            # If next row starts rowspan, header section is complete
            if next_has_rowspan and not has_header_merge:
                break

        return header_count

    def _make_single_table(
        self,
        grid: list[list[CellInfo | None]],
        table_rows: list[dict],
        min_col: int,
        max_col: int,
        layout_group: str | None,
    ) -> Section | None:
        """Create a single table section for a specific column range."""
        first_row = table_rows[0]["row"]
        last_row = table_rows[-1]["row"]

        # Filter table_rows to only rows that have data in this column range
        relevant_rows = []
        for ri in table_rows:
            r = ri["row"]
            has_data = any(
                grid[r][col] and grid[r][col].value is not None
                and (not grid[r][col].is_merged_cell or grid[r][col].is_merged_origin)
                for col in range(min_col, max_col + 1)
            )
            if has_data:
                relevant_rows.append(ri)

        if not relevant_rows:
            return None

        first_row = relevant_rows[0]["row"]
        last_row = relevant_rows[-1]["row"]

        header_count = self._detect_header_count(
            grid, relevant_rows, min_col, max_col,
        )

        # Build header_rows with colspan/rowspan info
        header_rows: list[list] = []
        for i in range(header_count):
            r = relevant_rows[i]["row"]
            header_cells: list = []
            col = min_col
            while col <= max_col:
                cell = grid[r][col] if r < len(grid) and col < len(grid[r]) else None
                if cell is None:
                    header_cells.append("")
                    col += 1
                    continue
                if cell.is_merged_cell and not cell.is_merged_origin:
                    col += 1
                    continue
                val = str(cell.value).strip() if cell.value else ""
                has_colspan = cell.merge_width > 1 and cell.col + cell.merge_width - 1 <= max_col
                has_rowspan = cell.merge_height > 1
                style = self._cell_style(cell)
                if has_colspan or has_rowspan or style:
                    cell_data: dict = {"value": val}
                    if has_colspan:
                        cell_data["colspan"] = min(cell.merge_width, max_col - cell.col + 1)
                    if has_rowspan:
                        cell_data["rowspan"] = cell.merge_height
                    if style:
                        cell_data["style"] = style
                    header_cells.append(cell_data)
                else:
                    header_cells.append(val)
                col += cell.merge_width if cell.merge_width > 1 else 1
            header_rows.append(header_cells)

        # Build flat headers for backward compat (from last header row)
        headers: list[str] | None = None
        if header_count > 0:
            last_hr = relevant_rows[header_count - 1]["row"]
            headers = []
            for col in range(min_col, max_col + 1):
                cell = grid[last_hr][col] if last_hr < len(grid) and col < len(grid[last_hr]) else None
                if cell and cell.is_merged_cell and not cell.is_merged_origin:
                    continue
                val = str(cell.value).strip() if cell and cell.value else ""
                headers.append(val)

        # Build data rows with rowspan/colspan info for merged cells
        rows: list[list] = []
        for ri in relevant_rows[header_count:]:
            r = ri["row"]
            row_data: list = []
            col = min_col
            while col <= max_col:
                cell = grid[r][col] if r < len(grid) and col < len(grid[r]) else None
                if cell is None:
                    row_data.append("")
                    col += 1
                    continue
                if cell.is_merged_cell and not cell.is_merged_origin:
                    row_data.append(None)
                    col += 1
                    continue
                val = str(cell.value).strip() if cell.value else ""
                has_colspan = cell.merge_width > 1 and cell.col + cell.merge_width - 1 <= max_col
                has_rowspan = cell.merge_height > 1
                fmt = _infer_format_type(cell.number_format)
                style = self._cell_style(cell)
                if has_colspan or has_rowspan or fmt or style:
                    cell_data = {"value": val}
                    if has_colspan:
                        cell_data["colspan"] = min(cell.merge_width, max_col - cell.col + 1)
                    if has_rowspan:
                        cell_data["rowspan"] = cell.merge_height
                    if fmt:
                        cell_data["format"] = fmt
                    if style:
                        cell_data["style"] = style
                    row_data.append(cell_data)
                else:
                    row_data.append(val)
                col += max(cell.merge_width, 1)
            rows.append(row_data)

        # Fill implicit row indices (blank cells that repeat the value above)
        rows = self._fill_implicit_row_indices(rows)

        content: dict = {}
        if header_rows:
            content["header_rows"] = header_rows
        if headers:
            content["headers"] = headers
        content["rows"] = rows
        if layout_group:
            content["layout_group"] = layout_group

        return Section(
            type=SectionType.TABLE,
            content=content,
            source_region=Region(
                min_row=first_row,
                min_col=min_col,
                max_row=last_row,
                max_col=max_col,
            ),
        )

    def _fill_implicit_row_indices(self, rows: list[list]) -> list[list]:
        """Fill blank cells in the first column(s) by repeating the value from above.

        Common pattern in Japanese tables: grouped rows share the first column value,
        which is only written once in the topmost row of the group.
        Only fills columns where:
        - The column is one of the first 2 columns
        - The cell is empty string (not a merge placeholder None)
        - There's a non-empty value above in the same column
        - The row has data in other columns (not entirely empty)
        """
        if not rows or len(rows) < 2:
            return rows

        # Determine how many leading columns might be implicit indices
        # (typically 1-2 columns: category, subcategory)
        max_index_cols = min(2, len(rows[0]) if rows[0] else 0)

        for col_idx in range(max_index_cols):
            last_value = None
            for row in rows:
                if col_idx >= len(row):
                    continue
                cell = row[col_idx]
                # Only fill empty string cells (not None which means merge placeholder)
                if cell == "":
                    # Check that row has data in other columns
                    has_other_data = any(
                        c not in ("", None) and not (isinstance(c, dict) and c.get("value", "") == "")
                        for j, c in enumerate(row) if j != col_idx
                    )
                    if has_other_data and last_value is not None:
                        row[col_idx] = last_value
                elif cell is not None:
                    # Update last_value (skip None merge placeholders)
                    val = cell if isinstance(cell, str) else (cell.get("value", "") if isinstance(cell, dict) else str(cell))
                    if val:
                        last_value = cell

        return rows

    def _make_key_value(
        self,
        grid: list[list[CellInfo | None]],
        kv_rows: list[dict],
        col_count: int,
    ) -> Section:
        """Create a key-value section."""
        pairs: dict[str, str] = {}
        for ri in kv_rows:
            cells = [c for c in ri["cells"] if c.value and c.value.strip()]
            # Try to pair them: bold=key, non-bold=value
            keys = [c for c in cells if c.font_bold or c.fill_color]
            vals = [c for c in cells if not c.font_bold and not c.fill_color]

            if keys and vals:
                for k, v in zip(keys, vals):
                    pairs[k.value.strip()] = v.value.strip()
            elif len(cells) == 2:
                # Just pair first two cells
                pairs[cells[0].value.strip()] = cells[1].value.strip()
            elif len(cells) >= 2:
                # Pair consecutive cells
                for j in range(0, len(cells) - 1, 2):
                    k = cells[j].value.strip()
                    v = cells[j + 1].value.strip()
                    pairs[k] = v

        first_row = kv_rows[0]["row"]
        last_row = kv_rows[-1]["row"]

        return Section(
            type=SectionType.KEY_VALUE,
            content=pairs,
            source_region=Region(
                min_row=first_row,
                min_col=1,
                max_row=last_row,
                max_col=col_count,
            ),
        )

    def _make_text(
        self,
        grid: list[list[CellInfo | None]],
        text_rows: list[dict],
        col_count: int,
    ) -> Section:
        """Create a text section from contiguous text rows."""
        lines: list[str] = []
        for ri in text_rows:
            cells = ri["cells"]
            values = [c.value.strip() for c in cells if c.value and c.value.strip()]
            if values:
                lines.append("  ".join(values))

        first_row = text_rows[0]["row"]
        last_row = text_rows[-1]["row"]

        return Section(
            type=SectionType.TEXT,
            content="\n".join(lines),
            source_region=Region(
                min_row=first_row,
                min_col=1,
                max_row=last_row,
                max_col=col_count,
            ),
        )
