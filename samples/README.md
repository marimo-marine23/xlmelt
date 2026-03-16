# Sample Excel Files for Converter Testing

This directory contains a generator script and sample Excel files that represent common Japanese enterprise document patterns. These files are designed to test an Excel-to-JSON/HTML converter tool with realistic, complex layouts.

## Prerequisites

```
pip install openpyxl
pip install Pillow  # optional, for sample_with_images.xlsx placeholder image generation
```

## Usage

```
python generate_samples.py
```

Files are generated into the `output/` subdirectory.

## Sample Files

| File | Japanese Name | Description |
|------|--------------|-------------|
| `sample_houganshi.xlsx` | Excel方眼紙（稟議書） | Grid-paper-style approval document with very narrow columns used as a layout grid. Contains merged cells forming title, key-value header, body text, and a cost breakdown table. Challenges parsers with the fine-grained column grid and extensive merging. |
| `sample_review.xlsx` | レビュー資料 | Multi-sheet design review document (overview, issue list, revision history). Features color-coded severity/status columns, key-value project info header, and merged note cells. |
| `sample_schedule.xlsx` | 工程表 | Gantt-chart-style project schedule with multi-level date headers (month, week). Task hierarchy with colored bars in narrow date columns. Mixed fixed-width and narrow-width columns. |
| `sample_spec.xlsx` | 仕様書 | Functional specification with deep heading hierarchy (chapter/section/subsection), version history table, document metadata, and long merged text blocks for specification body. |
| `sample_ledger.xlsx` | 管理台帳 | Issue management ledger with frozen panes, auto-filter, data validation dropdowns, two-level column headers (category + subcategory), and color-coded status/priority cells. |
| `sample_minutes.xlsx` | 議事録 | Meeting minutes with key-value header block, numbered agenda, discussion sections with large merged text cells, action item table, and next meeting info. |
| `sample_test_spec.xlsx` | テスト仕様書 | Two-sheet test specification (unit tests and integration tests). Features hierarchical category merging in left columns, pass/fail/pending color coding, and realistic test case content. |
| `sample_budget.xlsx` | 予算管理表 | Budget management table with multi-level row headers (merged category cells), quarterly columns, subtotals, grand total, percentage formatting, currency formatting, and variance coloring. |
| `sample_freetext.xlsx` | ただ文章を書いただけ（社内通知） | Free-text document written line by line in cells, no merged cells or table formatting. Headings use bold/larger fonts, indentation uses columns B/C/D, empty rows as paragraph separators. Content: office relocation notice. |
| `sample_freeform_report.xlsx` | お絵描き帳スタイルの報告書 | Sketchbook-style freeform report with large title, subtitle at offset column, free text paragraphs, two small tables at different positions, a bar chart, and key-value pairs scattered at arbitrary positions. Content: monthly sales report. |
| `sample_sketchpad.xlsx` | 完全自由配置（オリエンテーション資料） | Multiple "islands" of content placed at arbitrary positions with lots of empty space between them. Each island has different fonts, colors, and background fills. Includes a title block, contact info, numbered steps, a data table, and a caution box. Content: new employee orientation materials. |
| `sample_with_images.xlsx` | 画像入りドキュメント（作業手順書） | Document with programmatically generated placeholder images (requires Pillow/PIL), surrounding descriptive text, and a reference table. Gracefully degrades if Pillow is unavailable. Content: server restart procedure with screenshots. |
| `sample_mixed_document.xlsx` | 文章+テーブル+グラフ（営業実績分析） | Word-document-style Excel file with alternating text paragraphs, data tables, and charts (bar chart and line chart). Reads like a narrative report with embedded analytics. Content: annual sales performance analysis. |
| `sample_hierarchical_text.xlsx` | ワード風階層構造テキスト（設計方針書） | Structured document using Excel column positioning for hierarchy: chapters in column A (font 16), sections in column B (font 13), subsections in column C (font 11), body text in column D (font 10). Includes bullet points, numbered items, and inline bordered tables. Content: system design guidelines. |

## Patterns Covered

These samples collectively exercise the following parsing challenges:

- **Merged cells**: Horizontal spans, vertical spans, and block merges
- **Multi-sheet workbooks**: Sheet-level navigation and context
- **Grid-paper layout (方眼紙)**: Very narrow columns used as a pixel-like grid
- **Multi-level headers**: Both row and column header hierarchies
- **Color-coded data**: Background fills conveying semantic meaning (status, severity, Gantt bars)
- **Mixed content**: Key-value pairs, free-text blocks, and tabular data in one sheet
- **Japanese content**: Realistic names, dates, terminology, and formatting conventions
- **Data validation**: Dropdown lists for constrained input
- **Frozen panes and filters**: Structural metadata beyond cell content
- **Number formatting**: Currency, percentage, and custom formats
- **Border patterns**: Thin, medium, and thick borders delineating sections
- **Free-text documents**: Cell-by-cell line writing with no table structure, simulating Word-style documents in Excel
- **Freeform layout**: Content "islands" placed at arbitrary positions with empty space between them
- **Column-based indentation**: Using columns B, C, D for hierarchical indentation instead of actual indent
- **Embedded charts**: Bar charts and line charts generated from inline table data (openpyxl.chart)
- **Embedded images**: Programmatically generated placeholder images placed in worksheets (requires Pillow)
- **Mixed content flow**: Alternating paragraphs, tables, and charts in a single narrative document
- **Hierarchical text structure**: Chapter/section/subsection hierarchy expressed through column positioning and font sizing
