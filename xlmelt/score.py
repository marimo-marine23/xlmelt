"""AI-readability scoring: Excel vs JSON/HTML Before/After comparison.

Computes quantitative metrics comparing raw Excel content against
xlmelt's structured output across three categories:
  1. Readability  — how well an AI can understand the content
  2. Efficiency   — context window consumption (tokens, noise)
  3. Accuracy     — how faithfully the original data is preserved
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


# ─── Token estimation ─────────────────────────────────────────────────
# No external dependencies (tiktoken etc.) — approximate from char counts.
# Japanese business documents: ~1 token per 1.5 chars (CJK heavy)
# English/JSON structure: ~1 token per 4 chars
# Mixed default: chars ÷ 2.5

_CJK_RE = re.compile(r"[\u3000-\u9fff\uf900-\ufaff\U00020000-\U0002a6df]")


def estimate_tokens(text: str) -> int:
    """Estimate token count without external tokenizer.

    Uses character-based heuristic tuned for Claude/GPT tokenizers:
    - CJK characters: ~1 token per 1.5 chars
    - ASCII/Latin: ~1 token per 4 chars
    """
    if not text:
        return 0
    cjk_chars = len(_CJK_RE.findall(text))
    other_chars = len(text) - cjk_chars
    return int(cjk_chars / 1.5 + other_chars / 4)


# ─── Data classes ──────────────────────────────────────────────────────

@dataclass
class SheetMetrics:
    """Per-sheet raw metrics."""

    name: str
    total_cells: int = 0
    nonempty_cells: int = 0
    empty_cells: int = 0
    raw_text: str = ""  # concatenated cell values for token estimation

    # JSON analysis
    sections: int = 0
    headings: int = 0
    tables: int = 0
    key_values: int = 0
    lists: int = 0
    texts: int = 0
    images: int = 0
    max_heading_depth: int = 0
    json_chars: int = 0
    html_chars: int = 0

    # Coverage (from source_range)
    covered_cells: int = 0


@dataclass
class FileScore:
    """Complete Before/After score for a file."""

    name: str
    source: str
    sheets: list[SheetMetrics] = field(default_factory=list)

    # ── Readability ──
    readability_raw: float = 0.0      # Raw Excel score (0-100)
    readability_json: float = 0.0     # JSON output score (0-100)
    readability_improvement: float = 0.0

    # ── Efficiency ──
    raw_tokens: int = 0
    json_tokens: int = 0
    html_tokens: int = 0
    raw_chars: int = 0
    json_chars: int = 0
    html_chars: int = 0
    context_usage_raw_pct: float = 0.0   # % of 200K context window
    context_usage_json_pct: float = 0.0
    noise_tokens: int = 0
    info_per_1k_tokens: float = 0.0      # sections per 1K tokens
    token_saved_pct: float = 0.0         # negative = savings

    # ── Accuracy ──
    cell_coverage_pct: float = 0.0
    structure_ratio_pct: float = 0.0

    # ── Overall ──
    overall: float = 0.0

    def finalize(self) -> None:
        """Compute all scores from per-sheet metrics."""
        # Aggregate raw metrics
        total_cells = 0
        nonempty_cells = 0
        empty_cells = 0
        all_raw_text_parts: list[str] = []
        total_sections = 0
        total_structured = 0
        total_headings = 0
        max_depth = 0
        total_json_chars = 0
        total_html_chars = 0
        total_covered = 0
        type_set: set[str] = set()

        for s in self.sheets:
            total_cells += s.total_cells
            nonempty_cells += s.nonempty_cells
            empty_cells += s.empty_cells
            all_raw_text_parts.append(s.raw_text)
            total_sections += s.sections
            structured = s.headings + s.tables + s.key_values + s.lists + s.images
            total_structured += structured
            total_headings += s.headings
            max_depth = max(max_depth, s.max_heading_depth)
            total_json_chars += s.json_chars
            total_html_chars += s.html_chars
            total_covered += s.covered_cells
            if s.headings:
                type_set.add("heading")
            if s.tables:
                type_set.add("table")
            if s.key_values:
                type_set.add("kv")
            if s.lists:
                type_set.add("list")
            if s.texts:
                type_set.add("text")
            if s.images:
                type_set.add("image")

        # ── Build raw Excel representation for token estimation ──
        # Simulate what an AI sees when reading raw Excel: "Row 3, Col 2: value\n" per cell
        raw_cell_dump = "\n".join(all_raw_text_parts)
        self.raw_chars = len(raw_cell_dump)
        self.raw_tokens = estimate_tokens(raw_cell_dump)

        # JSON/HTML token estimation
        self.json_chars = total_json_chars
        self.html_chars = total_html_chars

        # For JSON tokens, estimate from actual JSON string
        # (will be set externally from the full JSON string)
        # For now use chars-based estimate if not set
        if self.json_tokens == 0 and self.json_chars > 0:
            self.json_tokens = int(self.json_chars / 4)  # JSON is mostly ASCII
        if self.html_tokens == 0 and self.html_chars > 0:
            self.html_tokens = int(self.html_chars / 4)

        # ── 1. Readability: Raw Excel ──
        # Raw Excel has no structure, no semantic types, just cell values
        # Score based on: data density, formatting cues, implicit structure
        raw_score = 0.0
        if total_cells > 0:
            # Data density: higher density = slightly more readable (less noise)
            density = nonempty_cells / total_cells
            raw_score += density * 20  # 0-20 pts

            # Implicit structure penalty: cells are just a grid with no semantics
            # Give a small base score for having data at all
            raw_score += 10  # base

            # Large files are harder to navigate without structure
            if nonempty_cells > 100:
                raw_score -= min(5, (nonempty_cells - 100) / 100)

        self.readability_raw = max(0, min(100, raw_score))

        # ── 1. Readability: JSON Output ──
        json_score = 0.0

        # Semantic typing (0-30): sections have explicit types
        if total_sections > 0:
            typed_ratio = total_structured / total_sections
            json_score += typed_ratio * 30

        # Heading hierarchy (0-25): navigable structure
        if total_headings > 0:
            json_score += min(15, total_headings * 2.5)  # up to 15 for headings
            json_score += min(10, max_depth * 5)          # up to 10 for depth

        # Section diversity (0-25): variety of semantic types
        json_score += min(25, len(type_set) * 5)

        # Compactness bonus (0-20): less noise than raw
        if self.raw_tokens > 0:
            reduction = 1.0 - (self.json_tokens / self.raw_tokens)
            json_score += max(0, min(20, reduction * 40))

        self.readability_json = max(0, min(100, json_score))
        self.readability_improvement = self.readability_json - self.readability_raw

        # ── 2. Efficiency ──
        CONTEXT_WINDOW = 200_000  # Claude context window
        self.context_usage_raw_pct = round((self.raw_tokens / CONTEXT_WINDOW) * 100, 2)
        self.context_usage_json_pct = round((self.json_tokens / CONTEXT_WINDOW) * 100, 2)

        # Noise tokens = tokens from empty/decoration cells in raw
        if total_cells > 0 and nonempty_cells > 0:
            noise_ratio = empty_cells / total_cells
            self.noise_tokens = int(self.raw_tokens * noise_ratio)
        else:
            self.noise_tokens = 0

        # Info density: sections per 1K tokens
        if self.json_tokens > 0:
            self.info_per_1k_tokens = round(total_sections / (self.json_tokens / 1000), 1)
        else:
            self.info_per_1k_tokens = 0.0

        # Token savings
        if self.raw_tokens > 0:
            self.token_saved_pct = round(
                ((self.json_tokens - self.raw_tokens) / self.raw_tokens) * 100, 1
            )
        else:
            self.token_saved_pct = 0.0

        # ── 3. Accuracy ──
        if nonempty_cells > 0:
            self.cell_coverage_pct = round((total_covered / nonempty_cells) * 100, 1)
        else:
            self.cell_coverage_pct = 100.0

        if total_sections > 0:
            self.structure_ratio_pct = round((total_structured / total_sections) * 100, 1)
        else:
            self.structure_ratio_pct = 0.0

        # ── Overall ──
        # Weighted: Readability improvement 35%, Efficiency 35%, Accuracy 30%
        # Normalize improvement to 0-100 scale (cap at +60 pts improvement → 100)
        readability_norm = max(0, min(100, self.readability_improvement * (100 / 60)))
        # Efficiency: token savings only count when JSON is smaller (negative %)
        # If JSON is larger (positive %), efficiency score is 0
        savings = max(0, -self.token_saved_pct)  # positive when JSON is smaller
        efficiency_norm = max(0, min(100, savings * (100 / 80)))
        # Accuracy: average of coverage and structure
        accuracy_norm = (self.cell_coverage_pct + self.structure_ratio_pct) / 2

        self.overall = round(
            readability_norm * 0.35
            + efficiency_norm * 0.35
            + accuracy_norm * 0.30,
            1,
        )

    def to_dict(self) -> dict:
        return {
            "file": self.name,
            "source": self.source,
            "readability": {
                "raw_excel": round(self.readability_raw, 1),
                "json_output": round(self.readability_json, 1),
                "improvement": round(self.readability_improvement, 1),
            },
            "efficiency": {
                "raw_tokens": self.raw_tokens,
                "json_tokens": self.json_tokens,
                "raw_chars": self.raw_chars,
                "json_chars": self.json_chars,
                "context_usage_raw_pct": self.context_usage_raw_pct,
                "context_usage_json_pct": self.context_usage_json_pct,
                "noise_tokens": self.noise_tokens,
                "info_per_1k_tokens": self.info_per_1k_tokens,
                "token_saved_pct": self.token_saved_pct,
            },
            "accuracy": {
                "cell_coverage_pct": self.cell_coverage_pct,
                "structure_ratio_pct": self.structure_ratio_pct,
            },
            "overall": self.overall,
            "sheets": [
                {
                    "name": s.name,
                    "cells": s.nonempty_cells,
                    "total_cells": s.total_cells,
                    "sections": s.sections,
                    "types": {
                        "heading": s.headings,
                        "table": s.tables,
                        "key_value": s.key_values,
                        "list": s.lists,
                        "text": s.texts,
                        "image": s.images,
                    },
                }
                for s in self.sheets
            ],
        }

    def summary(self) -> str:
        """Human-readable Before/After summary."""
        total_cells = sum(s.total_cells for s in self.sheets)
        nonempty = sum(s.nonempty_cells for s in self.sheets)
        n_sheets = len(self.sheets)

        lines = [
            f"Score: {self.name}",
            f"  Source: {self.source}"
            f" ({n_sheets} sheet{'s' if n_sheets != 1 else ''},"
            f" {total_cells} cells → {nonempty} content cells)",
            "",
            f"  ┌─ Readability ─────────────────────────────────────┐",
            f"  │ Raw Excel:   {self.readability_raw:5.1f} / 100"
            f"  (no structure, grid of cells)",
            f"  │ JSON Output: {self.readability_json:5.1f} / 100"
            f"  ({sum(s.headings for s in self.sheets)} headings,"
            f" {sum(s.tables for s in self.sheets)} tables,"
            f" semantic types)",
            f"  │ Improvement: {'+' if self.readability_improvement >= 0 else ''}"
            f"{self.readability_improvement:.0f} pts",
            f"  └─────────────────────────────────────────────────────┘",
            "",
            f"  ┌─ Efficiency (Context Window) ─────────────────────┐",
            f"  │                    Raw Excel    JSON Output",
            f"  │ Est. Tokens:    {self.raw_tokens:>10,}    {self.json_tokens:>10,}",
            f"  │ Context Usage:     {self.context_usage_raw_pct:>5.1f}%"
            f"        {self.context_usage_json_pct:>5.1f}%"
            f"    (of 200K)",
            f"  │ Noise Tokens:  {self.noise_tokens:>10,}             0",
            f"  │ Info/1K Tokens:                  {self.info_per_1k_tokens:>6.1f}"
            f"   (sections)",
            f"  │ Token Saved:                  {self.token_saved_pct:>+6.0f}%",
            f"  └─────────────────────────────────────────────────────┘",
            "",
            f"  ┌─ Accuracy ────────────────────────────────────────┐",
            f"  │ Cell Coverage:    {self.cell_coverage_pct:>6.1f}%",
            f"  │ Structure Ratio:  {self.structure_ratio_pct:>6.1f}%",
            f"  └─────────────────────────────────────────────────────┘",
        ]
        return "\n".join(lines)


# ─── Scoring function ──────────────────────────────────────────────────

def score_file(xlsx_path: Path) -> FileScore:
    """Score a single Excel file's AI-readability (Before/After comparison).

    Parses the raw Excel, converts to JSON/HTML, then computes
    readability, efficiency, and accuracy metrics.
    """
    from .core.analyzer import StructureAnalyzer
    from .core.parser import ExcelParser
    from .output.html_writer import HtmlWriter
    from .output.json_writer import JsonWriter

    result = FileScore(name=xlsx_path.stem, source=xlsx_path.name)

    range_re = re.compile(r"R(\d+)C(\d+):R(\d+)C(\d+)")

    # ── Parse raw Excel ──
    # Select parser based on file extension
    if xlsx_path.suffix.lower() == ".xls":
        from .core.xls_parser import XlsParser
        parser_cls = XlsParser
    else:
        parser_cls = ExcelParser
    with parser_cls(xlsx_path) as parser:
        for sheet_name in parser.sheet_names:
            grid, row_count, col_count = parser.parse_sheet(sheet_name)

            sm = SheetMetrics(name=sheet_name)
            raw_lines: list[str] = []

            for r in range(1, row_count + 1):
                for c in range(1, col_count + 1):
                    sm.total_cells += 1
                    cell = grid[r][c] if r < len(grid) and c < len(grid[r]) else None
                    if cell and cell.value is not None and str(cell.value).strip():
                        sm.nonempty_cells += 1
                        val = str(cell.value)
                        # Simulate raw cell dump format: "Row N, Col M: value"
                        raw_lines.append(f"Row {r}, Col {c}: {val}")
                    else:
                        sm.empty_cells += 1
                        # Empty cells still appear in raw dumps as noise
                        raw_lines.append(f"Row {r}, Col {c}: ")

            sm.raw_text = "\n".join(raw_lines)
            result.sheets.append(sm)

    # ── Convert to JSON/HTML ──
    analyzer = StructureAnalyzer()
    doc = analyzer.analyze(xlsx_path)
    json_writer = JsonWriter()
    json_str = json_writer.to_string(doc)
    json_data = json.loads(json_str)

    html_writer = HtmlWriter(include_style=False)
    html_str = html_writer.to_string(doc)

    # ── Analyze JSON structure ──
    doc_data = json_data.get("document", json_data)
    json_sheets = {s.get("name", ""): s for s in doc_data.get("sheets", [])}

    for sm in result.sheets:
        sheet_data = json_sheets.get(sm.name, {})
        sections = sheet_data.get("sections", [])
        sm.sections = len(sections)
        sm.json_chars = len(json.dumps(sheet_data, ensure_ascii=False))

        # Collect section types and coverage
        nonempty_set: set[tuple[int, int]] = set()
        # Build nonempty set for coverage
        # (re-parse would be expensive; use raw_text line count as proxy)
        # Actually, use source_range to determine covered cells
        def _analyze_sections(secs: list[dict]) -> None:
            for sec in secs:
                sec_type = sec.get("type", "")
                if sec_type == "heading":
                    sm.headings += 1
                    level = sec.get("level", 0)
                    sm.max_heading_depth = max(sm.max_heading_depth, level)
                elif sec_type == "table":
                    sm.tables += 1
                elif sec_type == "key_value":
                    sm.key_values += 1
                elif sec_type == "list":
                    sm.lists += 1
                elif sec_type == "text":
                    sm.texts += 1
                elif sec_type == "image":
                    sm.images += 1

                # Count covered cells from source_range
                sr = sec.get("source_range", "")
                m = range_re.match(sr)
                if m:
                    r1, c1, r2, c2 = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
                    for r in range(r1, r2 + 1):
                        for c in range(c1, c2 + 1):
                            nonempty_set.add((r, c))

                if sec.get("children"):
                    _analyze_sections(sec["children"])

        _analyze_sections(sections)
        sm.covered_cells = min(len(nonempty_set), sm.nonempty_cells)

    # HTML chars proportional split
    for sm in result.sheets:
        if len(result.sheets) == 1:
            sm.html_chars = len(html_str)
        else:
            total_json = sum(s.json_chars for s in result.sheets)
            if total_json > 0:
                sm.html_chars = int(len(html_str) * sm.json_chars / total_json)

    # Set token counts from actual strings
    result.json_tokens = estimate_tokens(json_str)
    result.html_tokens = estimate_tokens(html_str)

    result.finalize()
    return result


# ─── Report generation ─────────────────────────────────────────────────

def generate_score_report(scores: list[FileScore], report_path: Path) -> None:
    """Write a score report to file."""
    is_md = report_path.suffix.lower() == ".md"
    lines: list[str] = []

    if is_md:
        lines.append("# xlmelt AI-Readability Score Report")
        lines.append("")
        lines.append(f"Scored {len(scores)} file(s).")
        lines.append("")

        # Summary table
        lines.append("## Summary")
        lines.append("")
        lines.append(
            "| File | Raw | JSON | Improve | Raw Tokens | JSON Tokens"
            " | Token Saved | Coverage | **Overall** |"
        )
        lines.append("|---|---|---|---|---|---|---|---|---|")
        for s in scores:
            lines.append(
                f"| {s.name}"
                f" | {s.readability_raw:.0f}"
                f" | {s.readability_json:.0f}"
                f" | {'+' if s.readability_improvement >= 0 else ''}{s.readability_improvement:.0f}"
                f" | {s.raw_tokens:,}"
                f" | {s.json_tokens:,}"
                f" | {s.token_saved_pct:+.0f}%"
                f" | {s.cell_coverage_pct:.0f}%"
                f" | **{s.overall:.0f}** |"
            )
        lines.append("")

        # Averages
        if scores:
            avg_overall = sum(s.overall for s in scores) / len(scores)
            avg_raw = sum(s.readability_raw for s in scores) / len(scores)
            avg_json = sum(s.readability_json for s in scores) / len(scores)
            avg_saved = sum(s.token_saved_pct for s in scores) / len(scores)
            lines.append(f"**Average Overall: {avg_overall:.1f} / 100**")
            lines.append(
                f"(Readability: {avg_raw:.0f} → {avg_json:.0f},"
                f" Token Saved: {avg_saved:+.0f}%)"
            )
            lines.append("")

        # Score explanation
        lines.append("## Score Definitions")
        lines.append("")
        lines.append("### Readability (Before/After)")
        lines.append("")
        lines.append("| Component | Raw Excel | JSON Output |")
        lines.append("|---|---|---|")
        lines.append("| Semantic Typing | None (cells only) | heading/table/kv/list types |")
        lines.append("| Heading Hierarchy | None | H1-H4 levels |")
        lines.append("| Section Diversity | None | Multiple section types |")
        lines.append("| Table Clarity | Ambiguous grid | headers + rows structure |")
        lines.append("")

        lines.append("### Efficiency (Context Window)")
        lines.append("")
        lines.append("| Metric | Description |")
        lines.append("|---|---|")
        lines.append("| Est. Tokens | Estimated token count (CJK ÷ 1.5, ASCII ÷ 4) |")
        lines.append("| Context Usage | % of Claude's 200K context window |")
        lines.append("| Noise Tokens | Tokens wasted on empty/decorative cells |")
        lines.append("| Info/1K Tokens | Semantic sections per 1K tokens (information density) |")
        lines.append("| Token Saved | % reduction from raw → JSON |")
        lines.append("")

        lines.append("### Accuracy")
        lines.append("")
        lines.append("| Metric | Description |")
        lines.append("|---|---|")
        lines.append("| Cell Coverage | % of non-empty cells covered by source_range |")
        lines.append("| Structure Ratio | % of sections with semantic types (not plain text) |")
        lines.append("")

        lines.append("### Overall Score Weight")
        lines.append("")
        lines.append("| Category | Weight |")
        lines.append("|---|---|")
        lines.append("| Readability Improvement | 35% |")
        lines.append("| Token Efficiency | 35% |")
        lines.append("| Accuracy | 30% |")
        lines.append("")

        # Per-file details
        lines.append("## Details")
        lines.append("")
        for s in scores:
            lines.append(f"### {s.name}")
            lines.append("")
            lines.append(s.summary())
            lines.append("")
    else:
        lines.append("xlmelt AI-Readability Score Report")
        lines.append("=" * 55)
        lines.append(f"Scored {len(scores)} file(s).")
        lines.append("")
        for s in scores:
            lines.append(s.summary())
            lines.append("")

    report_path.parent.mkdir(parents=True, exist_ok=True)
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


# ─── Directory-level efficiency scoring ────────────────────────────────

CONTEXT_WINDOW = 200_000


@dataclass
class DirectoryScore:
    """Aggregate efficiency score for a directory of files.

    Compares reading strategies:
      A) Read all xlsx raw (cell dumps)
      B) Read all document.json
      C) Read manifest.json only (structure overview)
      D) Read manifest + selectively read N files
    """

    file_count: int = 0
    total_raw_tokens: int = 0
    total_json_tokens: int = 0
    manifest_tokens: int = 0
    avg_json_tokens_per_file: int = 0

    # Derived
    raw_context_pct: float = 0.0
    json_context_pct: float = 0.0
    manifest_context_pct: float = 0.0
    manifest_plus_1_tokens: int = 0
    manifest_plus_3_tokens: int = 0
    manifest_plus_5_tokens: int = 0
    manifest_plus_1_pct: float = 0.0
    manifest_plus_3_pct: float = 0.0
    manifest_plus_5_pct: float = 0.0

    # Savings vs raw
    json_vs_raw_pct: float = 0.0        # All JSON vs All raw
    manifest_vs_raw_pct: float = 0.0    # Manifest-only vs All raw
    manifest_3_vs_raw_pct: float = 0.0  # Manifest+3 vs All raw

    def finalize(self) -> None:
        """Compute derived metrics."""
        if self.file_count > 0:
            self.avg_json_tokens_per_file = self.total_json_tokens // self.file_count

        self.raw_context_pct = round((self.total_raw_tokens / CONTEXT_WINDOW) * 100, 1)
        self.json_context_pct = round((self.total_json_tokens / CONTEXT_WINDOW) * 100, 1)
        self.manifest_context_pct = round((self.manifest_tokens / CONTEXT_WINDOW) * 100, 1)

        avg = self.avg_json_tokens_per_file
        self.manifest_plus_1_tokens = self.manifest_tokens + avg * 1
        self.manifest_plus_3_tokens = self.manifest_tokens + avg * min(3, self.file_count)
        self.manifest_plus_5_tokens = self.manifest_tokens + avg * min(5, self.file_count)
        self.manifest_plus_1_pct = round((self.manifest_plus_1_tokens / CONTEXT_WINDOW) * 100, 1)
        self.manifest_plus_3_pct = round((self.manifest_plus_3_tokens / CONTEXT_WINDOW) * 100, 1)
        self.manifest_plus_5_pct = round((self.manifest_plus_5_tokens / CONTEXT_WINDOW) * 100, 1)

        if self.total_raw_tokens > 0:
            self.json_vs_raw_pct = round(
                ((self.total_json_tokens - self.total_raw_tokens) / self.total_raw_tokens) * 100, 1
            )
            self.manifest_vs_raw_pct = round(
                ((self.manifest_tokens - self.total_raw_tokens) / self.total_raw_tokens) * 100, 1
            )
            self.manifest_3_vs_raw_pct = round(
                ((self.manifest_plus_3_tokens - self.total_raw_tokens) / self.total_raw_tokens) * 100, 1
            )

    def summary(self) -> str:
        """Human-readable directory efficiency summary."""
        lines = [
            f"Directory Efficiency: {self.file_count} files",
            "",
            f"  ┌─ Context Window Comparison ─────────────────────────────┐",
            f"  │ Strategy                  Tokens    Context%    vs Raw  │",
            f"  │ ───────────────────────── ───────── ───────── ──────── │",
            f"  │ A) All xlsx raw        {self.total_raw_tokens:>10,}    {self.raw_context_pct:>5.1f}%    (base)  │",
            f"  │ B) All JSON            {self.total_json_tokens:>10,}    {self.json_context_pct:>5.1f}%   {self.json_vs_raw_pct:>+5.0f}%  │",
            f"  │ C) Manifest only       {self.manifest_tokens:>10,}    {self.manifest_context_pct:>5.1f}%   {self.manifest_vs_raw_pct:>+5.0f}%  │",
            f"  │ D) Manifest + 1 file   {self.manifest_plus_1_tokens:>10,}    {self.manifest_plus_1_pct:>5.1f}%           │",
            f"  │ E) Manifest + 3 files  {self.manifest_plus_3_tokens:>10,}    {self.manifest_plus_3_pct:>5.1f}%   {self.manifest_3_vs_raw_pct:>+5.0f}%  │",
            f"  │ F) Manifest + 5 files  {self.manifest_plus_5_tokens:>10,}    {self.manifest_plus_5_pct:>5.1f}%           │",
            f"  └─────────────────────────────────────────────────────────┘",
            "",
            f"  Manifest alone provides structure overview of all {self.file_count} files",
            f"  at {self.manifest_context_pct}% context cost ({self.manifest_vs_raw_pct:+.0f}% vs reading all raw).",
        ]
        return "\n".join(lines)

    def to_dict(self) -> dict:
        return {
            "file_count": self.file_count,
            "strategies": {
                "all_raw": {
                    "tokens": self.total_raw_tokens,
                    "context_pct": self.raw_context_pct,
                },
                "all_json": {
                    "tokens": self.total_json_tokens,
                    "context_pct": self.json_context_pct,
                    "vs_raw_pct": self.json_vs_raw_pct,
                },
                "manifest_only": {
                    "tokens": self.manifest_tokens,
                    "context_pct": self.manifest_context_pct,
                    "vs_raw_pct": self.manifest_vs_raw_pct,
                },
                "manifest_plus_1": {
                    "tokens": self.manifest_plus_1_tokens,
                    "context_pct": self.manifest_plus_1_pct,
                },
                "manifest_plus_3": {
                    "tokens": self.manifest_plus_3_tokens,
                    "context_pct": self.manifest_plus_3_pct,
                    "vs_raw_pct": self.manifest_3_vs_raw_pct,
                },
                "manifest_plus_5": {
                    "tokens": self.manifest_plus_5_tokens,
                    "context_pct": self.manifest_plus_5_pct,
                },
            },
            "avg_json_tokens_per_file": self.avg_json_tokens_per_file,
        }


def score_directory(scores: list[FileScore], manifest_path: Path | None = None) -> DirectoryScore:
    """Compute directory-level efficiency metrics from individual file scores.

    If manifest_path is provided, estimates manifest token count from the file.
    Otherwise, estimates from aggregate data.
    """
    ds = DirectoryScore(file_count=len(scores))
    ds.total_raw_tokens = sum(s.raw_tokens for s in scores)
    ds.total_json_tokens = sum(s.json_tokens for s in scores)

    if manifest_path and manifest_path.exists():
        with open(manifest_path, encoding="utf-8") as f:
            manifest_text = f.read()
        ds.manifest_tokens = estimate_tokens(manifest_text)
    else:
        # Estimate: ~50 tokens per file for outline metadata
        # (section type + title + basic stats per section)
        total_sections = sum(
            sum(s.headings + s.tables + s.key_values + s.lists + s.texts + s.images
                for s in score.sheets)
            for score in scores
        )
        ds.manifest_tokens = len(scores) * 50 + total_sections * 20

    ds.finalize()
    return ds
