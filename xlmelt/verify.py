"""Verify JSON↔HTML consistency.

Reads a document.json, re-renders it to HTML, and compares
structurally against the original document.html to detect
information loss or mapping discrepancies.
"""

from __future__ import annotations

import json
from html import escape
from pathlib import Path
from typing import Any


# ─── JSON → HTML re-renderer ────────────────────────────────────────────
# Mirrors HtmlWriter logic but reads from raw JSON dicts instead of model objects.

def render_html_from_json(data: dict, include_style: bool = True) -> str:
    """Re-render HTML from a parsed document.json dict."""
    doc = data.get("document", data)
    title = doc.get("title", "")
    source = doc.get("source", "")
    sheets = doc.get("sheets", [])

    parts: list[str] = []
    parts.append("<!DOCTYPE html>")
    parts.append('<html lang="ja">')
    parts.append("<head>")
    parts.append('<meta charset="utf-8">')
    parts.append(f"<title>{escape(title)}</title>")
    if include_style:
        parts.append(_default_style())
    parts.append("</head>")
    parts.append("<body>")
    parts.append(f"<h1>{escape(title)}</h1>")
    parts.append(f'<p class="source">Source: {escape(source)}</p>')

    multi_sheet = len(sheets) > 1
    for sheet in sheets:
        if multi_sheet:
            parts.append('<section class="sheet">')
            parts.append(f"<h2>{escape(sheet.get('name', ''))}</h2>")

        sections = sheet.get("sections", [])
        idx = 0
        while idx < len(sections):
            sec = sections[idx]
            content = sec.get("content")
            lg = None
            if sec.get("type") == "table" and isinstance(content, dict):
                lg = content.get("layout_group")
            if lg:
                group = [sec]
                while (idx + 1 < len(sections)
                       and sections[idx + 1].get("type") == "table"
                       and isinstance(sections[idx + 1].get("content"), dict)
                       and sections[idx + 1]["content"].get("layout_group") == lg):
                    idx += 1
                    group.append(sections[idx])
                parts.append('<div class="table-group">')
                for tbl in group:
                    parts.append(_render_table(tbl))
                parts.append("</div>")
            else:
                parts.append(_render_section(sec))
            idx += 1

        if multi_sheet:
            parts.append("</section>")

    parts.append("</body>")
    parts.append("</html>")
    return "\n".join(parts)


def _render_section(sec: dict) -> str:
    t = sec.get("type", "")
    if t == "heading":
        return _render_heading(sec)
    elif t == "table":
        return _render_table(sec)
    elif t == "key_value":
        return _render_key_value(sec)
    elif t == "text":
        return _render_text(sec)
    elif t == "list":
        return _render_list(sec)
    elif t == "image":
        return _render_image(sec)
    else:
        return _render_text(sec)


def _render_heading(sec: dict) -> str:
    level = min(max(sec.get("level", 2), 1), 6)
    text = escape(sec.get("title", "") or "")
    return f"<h{level}>{text}</h{level}>"


def _render_table(sec: dict) -> str:
    content = sec.get("content")
    if not content:
        return ""
    parts: list[str] = []
    title = sec.get("title")
    if title:
        parts.append(f'<p class="table-title">{escape(title)}</p>')
    parts.append("<table>")

    header_rows = content.get("header_rows")
    headers = content.get("headers")
    rows = content.get("rows", [])

    if header_rows:
        parts.append("<thead>")
        for hr in header_rows:
            parts.append("<tr>")
            for cell in hr:
                attrs = _merge_attrs(cell)
                val = _cell_value(cell)
                parts.append(f"<th{attrs}>{_escape_cell(val)}</th>")
            parts.append("</tr>")
        parts.append("</thead>")
    elif headers:
        parts.append("<thead><tr>")
        for h in headers:
            parts.append(f"<th>{escape(str(h))}</th>")
        parts.append("</tr></thead>")

    parts.append("<tbody>")
    for row in rows:
        parts.append("<tr>")
        for cell in row:
            if cell is None:
                continue
            attrs = _merge_attrs(cell)
            val = _cell_value(cell)
            parts.append(f"<td{attrs}>{_escape_cell(val)}</td>")
        parts.append("</tr>")
    parts.append("</tbody>")
    parts.append("</table>")
    return "\n".join(parts)


def _merge_attrs(cell) -> str:
    if not isinstance(cell, dict):
        return ""
    attrs = ""
    if cell.get("colspan", 1) > 1:
        attrs += f' colspan="{cell["colspan"]}"'
    if cell.get("rowspan", 1) > 1:
        attrs += f' rowspan="{cell["rowspan"]}"'
    style = _cell_style_css(cell)
    if style:
        attrs += f' style="{escape(style)}"'
    return attrs


def _cell_style_css(cell) -> str:
    if not isinstance(cell, dict):
        return ""
    style = cell.get("style")
    if not style or not isinstance(style, dict):
        return ""
    parts: list[str] = []
    bg = style.get("bg")
    if bg:
        color = bg if bg.startswith("#") else f"#{bg}"
        parts.append(f"background-color: {color}")
    color = style.get("color")
    if color:
        color = color if color.startswith("#") else f"#{color}"
        parts.append(f"color: {color}")
    if style.get("bold") == "true":
        parts.append("font-weight: bold")
    return "; ".join(parts)


def _cell_value(cell) -> str:
    if isinstance(cell, dict):
        return str(cell.get("value", ""))
    return str(cell)


def _escape_cell(val: str) -> str:
    escaped = escape(val)
    return escaped.replace("\n", "<br>")


def _render_key_value(sec: dict) -> str:
    content = sec.get("content")
    if not content or not isinstance(content, dict):
        return ""
    parts: list[str] = ["<dl>"]
    for key, value in content.items():
        parts.append(f"<dt>{escape(str(key))}</dt>")
        parts.append(f"<dd>{_escape_cell(str(value))}</dd>")
    parts.append("</dl>")
    return "\n".join(parts)


def _render_text(sec: dict) -> str:
    content = sec.get("content")
    if not content:
        return ""
    text = str(content)
    paragraphs = text.split("\n")
    parts = [f"<p>{escape(p)}</p>" for p in paragraphs if p.strip()]
    return "\n".join(parts)


def _render_list(sec: dict) -> str:
    content = sec.get("content")
    if not content:
        return ""
    if isinstance(content, dict):
        is_ordered = content.get("ordered", False)
        items = content.get("items", [])
    elif isinstance(content, list):
        is_ordered = False
        items = content
    else:
        items = [content]
        is_ordered = False

    tag = "ol" if is_ordered else "ul"
    parts = [f"<{tag}>"]
    for item in items:
        parts.append(f"<li>{escape(str(item))}</li>")
    parts.append(f"</{tag}>")
    return "\n".join(parts)


def _render_image(sec: dict) -> str:
    content = sec.get("content")
    if not content:
        return ""
    src = content.get("path", "")
    alt = content.get("alt", "")
    if not src:
        return f'<figure class="chart-placeholder"><p>[{escape(alt)}]</p></figure>'
    return f'<figure><img src="{escape(src)}" alt="{escape(alt)}"></figure>'


def _default_style() -> str:
    return """<style>
body { font-family: 'Hiragino Kaku Gothic ProN', 'Meiryo', sans-serif; max-width: 960px; margin: 2em auto; padding: 0 1em; color: #333; }
h1 { border-bottom: 2px solid #333; padding-bottom: 0.3em; }
h2 { color: #555; border-bottom: 1px solid #ccc; padding-bottom: 0.2em; }
.source { color: #888; font-size: 0.9em; }
table { border-collapse: collapse; width: 100%; margin: 1em 0; }
th, td { border: 1px solid #ccc; padding: 0.5em; text-align: left; }
th { background-color: #f5f5f5; font-weight: bold; }
dl { margin: 1em 0; }
dt { font-weight: bold; margin-top: 0.5em; }
dd { margin-left: 2em; }
.table-title { font-weight: bold; margin-bottom: 0.3em; }
.sheet { margin: 2em 0; padding: 1em; border: 1px solid #eee; border-radius: 4px; }
.table-group { display: flex; gap: 2em; align-items: flex-start; flex-wrap: wrap; margin: 1em 0; }
.table-group table { width: auto; }
.chart-placeholder { border: 2px dashed #ccc; padding: 2em; text-align: center; color: #888; margin: 1em 0; }
</style>"""


# ─── Structural comparator ──────────────────────────────────────────────

class VerifyResult:
    """Verification result with pass/fail items."""

    def __init__(self, name: str = "") -> None:
        self.name = name
        self.passed: list[str] = []
        self.failed: list[str] = []
        self.warnings: list[str] = []

    @property
    def ok(self) -> bool:
        return len(self.failed) == 0

    @property
    def total(self) -> int:
        return len(self.passed) + len(self.failed)

    def add_pass(self, msg: str) -> None:
        self.passed.append(msg)

    def add_fail(self, msg: str) -> None:
        self.failed.append(msg)

    def add_warning(self, msg: str) -> None:
        self.warnings.append(msg)

    def summary(self) -> str:
        lines: list[str] = []
        if self.failed:
            lines.append(f"FAIL: {len(self.failed)} issue(s) found")
            for f in self.failed:
                lines.append(f"  [x] {f}")
        if self.warnings:
            lines.append(f"WARN: {len(self.warnings)} warning(s)")
            for w in self.warnings:
                lines.append(f"  [!] {w}")
        if self.passed:
            lines.append(f"PASS: {len(self.passed)} check(s) passed")
        if self.ok:
            lines.append("Result: ALL CHECKS PASSED")
        else:
            lines.append(f"Result: {len(self.failed)} FAILURE(S)")
        return "\n".join(lines)


def verify_json_html(
    json_data: dict,
    original_html: str,
    name: str = "",
    xlsx_path: Path | None = None,
) -> VerifyResult:
    """Compare JSON content against original HTML structurally.

    Three-phase verification:
      1. Re-render JSON → HTML and compare line-by-line
      2. Structural checks (section counts, content matching)
      3. xlsx coverage check (if xlsx_path provided)
    """
    result = VerifyResult(name=name)

    # Phase 1: Re-render and compare
    # Detect if original HTML was generated with --no-style
    include_style = "<style>" in original_html
    rerendered = render_html_from_json(json_data, include_style=include_style)
    orig_lines = [l.strip() for l in original_html.strip().splitlines() if l.strip()]
    new_lines = [l.strip() for l in rerendered.strip().splitlines() if l.strip()]

    if orig_lines == new_lines:
        result.add_pass("HTML re-render: exact match (JSON contains all information)")
    else:
        # Find first difference
        diff_details = _find_html_diffs(orig_lines, new_lines)
        if diff_details:
            result.add_fail(f"HTML re-render: {len(diff_details)} line(s) differ")
            for detail in diff_details[:10]:
                result.add_fail(f"  {detail}")
            if len(diff_details) > 10:
                result.add_fail(f"  ... and {len(diff_details) - 10} more")
        else:
            result.add_pass("HTML re-render: match (minor whitespace differences only)")

    # Phase 2: Structural checks
    doc = json_data.get("document", json_data)
    sheets = doc.get("sheets", [])

    for sheet in sheets:
        sheet_name = sheet.get("name", "?")
        sections = sheet.get("sections", [])

        for i, sec in enumerate(sections):
            sec_type = sec.get("type", "unknown")
            loc = f"Sheet[{sheet_name}].Section[{i}]({sec_type})"

            # Check: every section has a type
            if not sec.get("type"):
                result.add_fail(f"{loc}: missing 'type' field")
                continue

            result.add_pass(f"{loc}: type present")

            content = sec.get("content")

            if sec_type == "heading":
                if sec.get("title"):
                    result.add_pass(f"{loc}: title='{sec['title']}'")
                else:
                    result.add_fail(f"{loc}: heading has no title")
                if not sec.get("level"):
                    result.add_warning(f"{loc}: heading has no level (defaults to 0)")

            elif sec_type == "table":
                if not isinstance(content, dict):
                    result.add_fail(f"{loc}: content is not a dict")
                    continue
                rows = content.get("rows", [])
                headers = content.get("headers", [])
                header_rows = content.get("header_rows", [])
                if not headers and not header_rows:
                    result.add_warning(f"{loc}: no headers defined")
                if not rows:
                    result.add_warning(f"{loc}: no data rows")
                else:
                    result.add_pass(f"{loc}: {len(rows)} rows")
                # Check cell integrity
                _verify_table_cells(content, loc, result)

            elif sec_type == "key_value":
                if not isinstance(content, dict):
                    result.add_fail(f"{loc}: content is not a dict")
                elif not content:
                    result.add_warning(f"{loc}: empty key-value pairs")
                else:
                    result.add_pass(f"{loc}: {len(content)} pair(s)")

            elif sec_type == "list":
                if isinstance(content, dict):
                    items = content.get("items", [])
                    if not items:
                        result.add_warning(f"{loc}: empty list")
                    else:
                        result.add_pass(f"{loc}: {len(items)} item(s)")
                elif isinstance(content, list):
                    result.add_pass(f"{loc}: {len(content)} item(s) (legacy format)")
                else:
                    result.add_fail(f"{loc}: invalid list content type")

            elif sec_type == "text":
                if content:
                    result.add_pass(f"{loc}: len={len(str(content))}")
                else:
                    result.add_warning(f"{loc}: empty text")

            elif sec_type == "image":
                if isinstance(content, dict):
                    if content.get("path") or content.get("alt"):
                        result.add_pass(f"{loc}: path='{content.get('path', '')}' alt='{content.get('alt', '')}'")
                    else:
                        result.add_fail(f"{loc}: image has no path or alt")
                else:
                    result.add_fail(f"{loc}: invalid image content")

    # Phase 3: xlsx coverage check
    if xlsx_path and xlsx_path.exists():
        _verify_xlsx_coverage(json_data, xlsx_path, result)

    return result


def _verify_table_cells(content: dict, loc: str, result: VerifyResult) -> None:
    """Verify table cell integrity — check for information loss."""
    rows = content.get("rows", [])
    header_rows = content.get("header_rows", [])

    # Check that cell dicts have 'value' key
    all_cells = []
    for hr in header_rows:
        all_cells.extend(hr)
    for row in rows:
        all_cells.extend(row)

    dict_cells = [c for c in all_cells if isinstance(c, dict)]
    for cell in dict_cells:
        if "value" not in cell:
            result.add_fail(f"{loc}: cell dict missing 'value' key: {cell}")
            return

    # Check colspan/rowspan values are positive
    for cell in dict_cells:
        cs = cell.get("colspan", 1)
        rs = cell.get("rowspan", 1)
        if not isinstance(cs, (int, float)) or not isinstance(rs, (int, float)):
            result.add_fail(f"{loc}: non-numeric span: colspan={cs} rowspan={rs}")
            return
        if cs < 1 or rs < 1:
            result.add_fail(f"{loc}: invalid span: colspan={cs} rowspan={rs}")
            return


def _verify_xlsx_coverage(json_data: dict, xlsx_path: Path, result: VerifyResult) -> None:
    """Check which xlsx cells are covered by JSON sections (via source_range)."""
    import re

    try:
        from ..core.parser import ExcelParser
        from ..core.xls_parser import XlsParser
    except ImportError:
        from xlmelt.core.parser import ExcelParser
        from xlmelt.core.xls_parser import XlsParser

    range_re = re.compile(r"R(\d+)C(\d+):R(\d+)C(\d+)")

    doc = json_data.get("document", json_data)
    json_sheets = {s.get("name", ""): s for s in doc.get("sheets", [])}

    parser_cls = XlsParser if xlsx_path.suffix.lower() == ".xls" else ExcelParser
    try:
        with parser_cls(xlsx_path) as parser:
            for sheet_name in parser.sheet_names:
                grid, row_count, col_count = parser.parse_sheet(sheet_name)

                # Collect all non-empty cells
                nonempty: set[tuple[int, int]] = set()
                for r in range(1, row_count + 1):
                    for c in range(1, col_count + 1):
                        cell = grid[r][c] if r < len(grid) and c < len(grid[r]) else None
                        if cell and cell.value is not None and str(cell.value).strip():
                            nonempty.add((r, c))

                if not nonempty:
                    continue

                # Collect covered cells from source_range (including children)
                covered: set[tuple[int, int]] = set()
                json_sheet = json_sheets.get(sheet_name)
                if json_sheet:
                    def _collect_ranges(sections: list) -> None:
                        for sec in sections:
                            sr = sec.get("source_range", "")
                            m = range_re.match(sr)
                            if m:
                                r1, c1, r2, c2 = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
                                for r in range(r1, r2 + 1):
                                    for c in range(c1, c2 + 1):
                                        covered.add((r, c))
                            if sec.get("children"):
                                _collect_ranges(sec["children"])
                    _collect_ranges(json_sheet.get("sections", []))

                missed = nonempty - covered
                total = len(nonempty)
                covered_count = total - len(missed)
                pct = (covered_count / total * 100) if total > 0 else 100.0

                result.add_pass(
                    f"Coverage Sheet[{sheet_name}]: {covered_count}/{total} cells ({pct:.1f}%)"
                )

                if missed:
                    # Group missed cells into contiguous row ranges for readability
                    missed_by_row: dict[int, list[int]] = {}
                    for r, c in sorted(missed):
                        missed_by_row.setdefault(r, []).append(c)

                    missed_samples: list[str] = []
                    for r in sorted(missed_by_row)[:10]:
                        cols = missed_by_row[r]
                        # Show cell values for context
                        cell_previews: list[str] = []
                        for c in cols[:5]:
                            cell = grid[r][c] if r < len(grid) and c < len(grid[r]) else None
                            val = str(cell.value)[:30] if cell and cell.value is not None else ""
                            cell_previews.append(f"R{r}C{c}='{val}'")
                        line = ", ".join(cell_previews)
                        if len(cols) > 5:
                            line += f" ... (+{len(cols)-5} more cols)"
                        missed_samples.append(line)

                    result.add_warning(
                        f"Coverage Sheet[{sheet_name}]: {len(missed)} cell(s) not covered by any section"
                    )
                    for sample in missed_samples:
                        result.add_warning(f"  Uncovered: {sample}")
                    if len(missed_by_row) > 10:
                        result.add_warning(
                            f"  ... and {len(missed_by_row) - 10} more row(s)"
                        )

    except Exception as e:
        result.add_warning(f"Coverage check failed: {e}")


def _find_html_diffs(orig: list[str], new: list[str]) -> list[str]:
    """Find specific line differences between two HTML line lists."""
    diffs: list[str] = []
    max_len = max(len(orig), len(new))
    for i in range(max_len):
        o = orig[i] if i < len(orig) else "<missing>"
        n = new[i] if i < len(new) else "<missing>"
        if o != n:
            # Truncate long lines for readability
            o_display = o[:120] + "..." if len(o) > 120 else o
            n_display = n[:120] + "..." if len(n) > 120 else n
            diffs.append(f"Line {i+1}:")
            diffs.append(f"  orig: {o_display}")
            diffs.append(f"  json: {n_display}")
    return diffs


def verify_file(output_dir: Path, xlsx_path: Path | None = None) -> VerifyResult:
    """Verify JSON↔HTML consistency for a converted file's output directory.

    If xlsx_path is provided, also checks cell coverage against the original file.
    """
    json_path = output_dir / "document.json"
    html_path = output_dir / "document.html"

    result = VerifyResult()

    if not json_path.exists():
        result.add_fail(f"document.json not found in {output_dir}")
        return result
    if not html_path.exists():
        result.add_fail(f"document.html not found in {output_dir}")
        return result

    with open(json_path, encoding="utf-8") as f:
        json_data = json.load(f)
    with open(html_path, encoding="utf-8") as f:
        original_html = f.read()

    return verify_json_html(json_data, original_html, name=output_dir.name, xlsx_path=xlsx_path)


def _categorize_items(items: list[str]) -> tuple[list[str], list[str]]:
    """Split items into coverage-related and other items."""
    coverage: list[str] = []
    other: list[str] = []
    for item in items:
        if item.startswith("Coverage ") or item.strip().startswith("Uncovered:"):
            coverage.append(item)
        else:
            other.append(item)
    return coverage, other


def generate_report(results: list[VerifyResult], report_path: Path) -> None:
    """Write a detailed verification report to a file.

    Supports .md and .txt output based on file extension.
    """
    is_md = report_path.suffix.lower() == ".md"
    lines: list[str] = []

    # Header
    if is_md:
        lines.append("# xlmelt Verification Report")
        lines.append("")
        lines.append(f"Verified {len(results)} file(s).")
        lines.append("")
    else:
        lines.append("xlmelt Verification Report")
        lines.append("=" * 40)
        lines.append(f"Verified {len(results)} file(s).")
        lines.append("")

    # Summary table
    total_pass = sum(len(r.passed) for r in results)
    total_fail = sum(len(r.failed) for r in results)
    total_warn = sum(len(r.warnings) for r in results)
    files_ok = sum(1 for r in results if r.ok)
    files_fail = len(results) - files_ok

    if is_md:
        lines.append("## Summary")
        lines.append("")
        lines.append(f"| | Count |")
        lines.append(f"|---|---|")
        lines.append(f"| Files verified | {len(results)} |")
        lines.append(f"| Files PASS | {files_ok} |")
        lines.append(f"| Files FAIL | {files_fail} |")
        lines.append(f"| Total checks | {total_pass + total_fail} |")
        lines.append(f"| Passed | {total_pass} |")
        lines.append(f"| Failed | {total_fail} |")
        lines.append(f"| Warnings | {total_warn} |")
        lines.append("")
    else:
        lines.append(f"Files verified: {len(results)}")
        lines.append(f"Files PASS: {files_ok}  /  Files FAIL: {files_fail}")
        lines.append(f"Total checks: {total_pass + total_fail}  (pass: {total_pass}, fail: {total_fail}, warn: {total_warn})")
        lines.append("")

    # xlsx Coverage Summary (Phase 3) — dedicated section
    has_coverage = any(
        any(p.startswith("Coverage ") for p in r.passed)
        or any(w.startswith("Coverage ") for w in r.warnings)
        for r in results
    )
    if has_coverage:
        if is_md:
            lines.append("## xlsx Cell Coverage (Phase 3)")
            lines.append("")
            lines.append("| File | Sheet | Covered | Total | Rate |")
            lines.append("|---|---|---|---|---|")
        else:
            lines.append("xlsx Cell Coverage (Phase 3)")
            lines.append("-" * 40)

        import re
        cov_re = re.compile(r"Coverage Sheet\[(.+?)\]: (\d+)/(\d+) cells \((.+?)%\)")

        for r in results:
            name = r.name or "(unknown)"
            for p in r.passed:
                m = cov_re.match(p)
                if m:
                    sheet, covered, total, pct = m.group(1), m.group(2), m.group(3), m.group(4)
                    if is_md:
                        lines.append(f"| {name} | {sheet} | {covered} | {total} | {pct}% |")
                    else:
                        lines.append(f"  {name} / {sheet}: {covered}/{total} ({pct}%)")

        lines.append("")

        # Show uncovered cell details after the table
        for r in results:
            name = r.name or "(unknown)"
            cov_warnings, _ = _categorize_items(r.warnings)
            if cov_warnings:
                if is_md:
                    lines.append(f"**{name} — uncovered cells:**")
                    lines.append("")
                    for w in cov_warnings:
                        lines.append(f"- {w}")
                    lines.append("")
                else:
                    lines.append(f"  {name} — uncovered cells:")
                    for w in cov_warnings:
                        lines.append(f"    {w}")

    # Per-file details
    if is_md:
        lines.append("## Details (Phase 1 & 2)")
        lines.append("")
    else:
        lines.append("Details (Phase 1 & 2)")
        lines.append("-" * 40)

    for r in results:
        name = r.name or "(unknown)"
        status = "PASS" if r.ok else "FAIL"

        # Separate coverage items from other items
        _, other_passed = _categorize_items(r.passed)
        _, other_warnings = _categorize_items(r.warnings)

        if is_md:
            icon = "+" if r.ok else "x"
            lines.append(f"### [{icon}] {name} — {status}")
            lines.append("")
            if r.failed:
                lines.append(f"**Failures ({len(r.failed)}):**")
                lines.append("")
                for f in r.failed:
                    lines.append(f"- {f}")
                lines.append("")
            if other_warnings:
                lines.append(f"**Warnings ({len(other_warnings)}):**")
                lines.append("")
                for w in other_warnings:
                    lines.append(f"- {w}")
                lines.append("")
            if other_passed:
                lines.append(f"**Passed ({len(other_passed)}):**")
                lines.append("")
                for p in other_passed:
                    lines.append(f"- {p}")
                lines.append("")
        else:
            lines.append(f"--- {name} [{status}] ---")
            if r.failed:
                for f in r.failed:
                    lines.append(f"  [FAIL] {f}")
            if other_warnings:
                for w in other_warnings:
                    lines.append(f"  [WARN] {w}")
            if other_passed:
                for p in other_passed:
                    lines.append(f"  [PASS] {p}")
            lines.append("")

    # Final verdict
    if files_fail == 0:
        verdict = "ALL FILES PASSED"
    else:
        verdict = f"{files_fail} FILE(S) FAILED"

    if is_md:
        lines.append("---")
        lines.append(f"**Result: {verdict}**")
    else:
        lines.append("=" * 40)
        lines.append(f"Result: {verdict}")

    report_path.parent.mkdir(parents=True, exist_ok=True)
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
