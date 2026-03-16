"""Index and manifest generation for batch output directories.

Generates:
  - index.html: human-navigable page linking to all converted documents
  - manifest.json: AI-optimized catalog with section outlines (no full content)
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import datetime, timezone
from html import escape
from pathlib import Path
from typing import Any

from .. import __version__
from ..core.model import DocumentModel, Section, _region_to_a1


# ─── Section outline extraction ────────────────────────────────────────

def section_outline(section: Section) -> dict:
    """Create a compact outline dict for a section (type + key metadata only).

    Used by both manifest.json generation and `inspect --json`.
    """
    result: dict = {"type": section.type.value}
    if section.type.value == "heading":
        result["level"] = section.level
        result["title"] = section.title
    elif section.type.value == "table":
        content = section.content if isinstance(section.content, dict) else {}
        rows = content.get("rows", [])
        headers = content.get("headers", [])
        result["rows"] = len(rows)
        result["headers"] = headers
    elif section.type.value == "key_value":
        pairs = section.content if isinstance(section.content, dict) else {}
        result["keys"] = list(pairs.keys())
    elif section.type.value == "list":
        if isinstance(section.content, dict):
            items = section.content.get("items", [])
            result["ordered"] = section.content.get("ordered", False)
        elif isinstance(section.content, list):
            items = section.content
            result["ordered"] = False
        else:
            items = []
            result["ordered"] = False
        result["count"] = len(items)
    elif section.type.value == "text":
        text = str(section.content or "")
        result["length"] = len(text)
        result["preview"] = text[:80]
    elif section.type.value == "image":
        if isinstance(section.content, dict):
            result["path"] = section.content.get("path", "")
            result["alt"] = section.content.get("alt", "")
    if section.source_region:
        result["range"] = _region_to_a1(section.source_region)
    if section.children:
        result["children"] = [section_outline(c) for c in section.children]
    return result


# ─── File entry ────────────────────────────────────────────────────────

@dataclass
class FileEntry:
    """Manifest entry for one converted file."""

    name: str           # stem (e.g., "sample_spec")
    source: str         # original filename (e.g., "sample_spec.xlsx")
    sheets: list[dict] = field(default_factory=list)
    total_sections: int = 0
    total_images: int = 0


def build_entry_from_doc(stem: str, source: str, doc: DocumentModel) -> FileEntry:
    """Build a FileEntry from an in-memory DocumentModel (during conversion)."""
    entry = FileEntry(name=stem, source=source)
    entry.total_images = len(doc.images)

    for sheet in doc.sheets:
        type_counts: dict[str, int] = {}
        for s in sheet.sections:
            t = s.type.value
            type_counts[t] = type_counts.get(t, 0) + 1

        sheet_data = {
            "name": sheet.name,
            "section_summary": type_counts,
            "outline": [section_outline(s) for s in sheet.sections],
        }
        entry.sheets.append(sheet_data)
        entry.total_sections += len(sheet.sections)

    return entry


def build_entry_from_output(file_output_dir: Path) -> FileEntry | None:
    """Reconstruct a FileEntry from existing output files on disk.

    Reads metadata.json for basic info and document.json for section outlines.
    """
    meta_path = file_output_dir / "metadata.json"
    json_path = file_output_dir / "document.json"

    if not meta_path.exists():
        return None

    try:
        with open(meta_path, encoding="utf-8") as f:
            meta = json.load(f)
    except Exception:
        return None

    stem = file_output_dir.name
    source = meta.get("source", f"{stem}.xlsx")
    entry = FileEntry(name=stem, source=source)
    entry.total_sections = meta.get("total_sections", 0)
    entry.total_images = meta.get("total_images", 0)

    # Try to read document.json for detailed outlines
    if json_path.exists():
        try:
            with open(json_path, encoding="utf-8") as f:
                doc_data = json.load(f)
            doc_root = doc_data.get("document", doc_data)
            for sheet in doc_root.get("sheets", []):
                sheet_name = sheet.get("name", "")
                sections = sheet.get("sections", [])
                type_counts: dict[str, int] = {}
                for sec in sections:
                    t = sec.get("type", "unknown")
                    type_counts[t] = type_counts.get(t, 0) + 1
                sheet_data = {
                    "name": sheet_name,
                    "section_summary": type_counts,
                    "outline": [_outline_from_json(sec) for sec in sections],
                }
                entry.sheets.append(sheet_data)
        except Exception:
            # Fall back to metadata-only info
            for sheet_meta in meta.get("sheets", []):
                entry.sheets.append({
                    "name": sheet_meta.get("name", ""),
                    "section_summary": sheet_meta.get("section_types", {}),
                    "outline": [],
                })
    else:
        # No document.json — use metadata only
        for sheet_meta in meta.get("sheets", []):
            entry.sheets.append({
                "name": sheet_meta.get("name", ""),
                "section_summary": sheet_meta.get("section_types", {}),
                "outline": [],
            })

    return entry


def _outline_from_json(sec: dict) -> dict:
    """Build outline from a raw JSON section dict (from document.json)."""
    result: dict = {"type": sec.get("type", "unknown")}
    sec_type = sec.get("type", "")

    if sec_type == "heading":
        result["level"] = sec.get("level", 0)
        result["title"] = sec.get("title", "")
    elif sec_type == "table":
        content = sec.get("content") or {}
        if isinstance(content, dict):
            result["rows"] = len(content.get("rows", []))
            result["headers"] = content.get("headers", [])
    elif sec_type == "key_value":
        content = sec.get("content") or {}
        if isinstance(content, dict):
            result["keys"] = list(content.keys())
    elif sec_type == "list":
        content = sec.get("content") or {}
        if isinstance(content, dict):
            result["count"] = len(content.get("items", []))
            result["ordered"] = content.get("ordered", False)
        elif isinstance(content, list):
            result["count"] = len(content)
            result["ordered"] = False
    elif sec_type == "text":
        content = sec.get("content") or ""
        result["length"] = len(str(content))
        result["preview"] = str(content)[:80]
    elif sec_type == "image":
        content = sec.get("content") or {}
        if isinstance(content, dict):
            result["path"] = content.get("path", "")
            result["alt"] = content.get("alt", "")

    if sec.get("source_range_a1"):
        result["range"] = sec["source_range_a1"]
    if sec.get("children"):
        result["children"] = [_outline_from_json(c) for c in sec["children"]]
    return result


# ─── Manifest writer ──────────────────────────────────────────────────

def write_manifest(entries: list[FileEntry], output_dir: Path) -> Path:
    """Write manifest.json — AI-optimized catalog of all converted files."""
    # Aggregate summary
    total_sheets = sum(len(e.sheets) for e in entries)
    total_sections = sum(e.total_sections for e in entries)
    type_counts: dict[str, int] = {}
    for e in entries:
        for sheet in e.sheets:
            for t, count in sheet.get("section_summary", {}).items():
                type_counts[t] = type_counts.get(t, 0) + count

    manifest = {
        "xlmelt_version": __version__,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "file_count": len(entries),
        "files": [
            {
                "name": e.name,
                "source": e.source,
                "json_path": f"{e.name}/document.json",
                "html_path": f"{e.name}/document.html",
                "sheets": e.sheets,
                "total_sections": e.total_sections,
                "total_images": e.total_images,
            }
            for e in entries
        ],
        "summary": {
            "total_sheets": total_sheets,
            "total_sections": total_sections,
            "type_counts": type_counts,
        },
    }

    manifest_path = output_dir / "manifest.json"
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)
    return manifest_path


# ─── Index HTML writer ─────────────────────────────────────────────────

def write_index_html(entries: list[FileEntry], output_dir: Path) -> Path:
    """Write index.html — human-navigable page linking to all documents."""
    parts: list[str] = []
    parts.append("<!DOCTYPE html>")
    parts.append('<html lang="ja">')
    parts.append("<head>")
    parts.append('<meta charset="utf-8">')
    parts.append("<title>xlmelt Document Index</title>")
    parts.append(_index_style())
    parts.append("</head>")
    parts.append("<body>")
    parts.append("<h1>xlmelt Document Index</h1>")
    parts.append(f'<p class="meta">Generated: {datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")}'
                 f' | {len(entries)} file(s) | xlmelt v{__version__}</p>')

    # Summary table
    parts.append("<table>")
    parts.append("<thead><tr>")
    parts.append("<th>File</th><th>Source</th><th>Sheets</th>"
                 "<th>Sections</th><th>Types</th><th>Images</th>")
    parts.append("</tr></thead>")
    parts.append("<tbody>")

    total_sections = 0
    total_images = 0

    for e in entries:
        html_link = f"{e.name}/document.html"
        json_link = f"{e.name}/document.json"

        # Collect type badges
        type_counts: dict[str, int] = {}
        for sheet in e.sheets:
            for t, count in sheet.get("section_summary", {}).items():
                type_counts[t] = type_counts.get(t, 0) + count
        # Sanitize type names to prevent XSS in class attributes
        _allowed_types = {"heading", "table", "key_value", "text", "list", "image", "chart"}
        type_badges = " ".join(
            f'<span class="badge badge-{t}">{escape(t)}:{c}</span>'
            for t, c in sorted(type_counts.items())
            if t in _allowed_types
        )

        parts.append("<tr>")
        parts.append(f'<td><a href="{escape(html_link)}">{escape(e.name)}</a>'
                     f' <a href="{escape(json_link)}" class="json-link">[JSON]</a></td>')
        parts.append(f"<td>{escape(e.source)}</td>")
        parts.append(f"<td>{len(e.sheets)}</td>")
        parts.append(f"<td>{e.total_sections}</td>")
        parts.append(f"<td>{type_badges}</td>")
        parts.append(f"<td>{e.total_images}</td>")
        parts.append("</tr>")

        total_sections += e.total_sections
        total_images += e.total_images

    parts.append("</tbody>")
    parts.append("<tfoot><tr>")
    parts.append(f'<td colspan="3"><strong>Total</strong></td>')
    parts.append(f"<td><strong>{total_sections}</strong></td>")
    parts.append(f'<td colspan="2"><strong>{total_images} images</strong></td>')
    parts.append("</tr></tfoot>")
    parts.append("</table>")

    # Per-file outlines
    parts.append("<h2>Document Outlines</h2>")
    for e in entries:
        parts.append(f'<details>')
        parts.append(f'<summary><strong>{escape(e.name)}</strong>'
                     f' ({len(e.sheets)} sheet, {e.total_sections} sections)</summary>')
        for sheet in e.sheets:
            sheet_name = sheet.get("name", "")
            parts.append(f'<h3>{escape(sheet_name)}</h3>')
            outline = sheet.get("outline", [])
            if outline:
                parts.append("<ul>")
                for sec in outline:
                    parts.append(f"<li>{_outline_to_html(sec)}</li>")
                parts.append("</ul>")
        parts.append("</details>")

    parts.append(f'<p class="footer">Generated by <a href="https://github.com/marimo-marine23/xlmelt">xlmelt</a> v{__version__}</p>')
    parts.append("</body>")
    parts.append("</html>")

    index_path = output_dir / "index.html"
    index_path.parent.mkdir(parents=True, exist_ok=True)
    with open(index_path, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))
    return index_path


def _outline_to_html(sec: dict) -> str:
    """Render a single outline entry as inline HTML."""
    t = sec.get("type", "?")
    if t == "heading":
        level = sec.get("level", 0)
        title = escape(sec.get("title", ""))
        text = f'<strong>[H{level}]</strong> {title}'
    elif t == "table":
        rows = sec.get("rows", 0)
        headers = sec.get("headers", [])
        h_str = ", ".join(str(h) for h in headers[:6])
        if len(headers) > 6:
            h_str += f" (+{len(headers)-6})"
        text = f'<strong>[TABLE]</strong> {rows} rows [{escape(h_str)}]'
    elif t == "key_value":
        keys = sec.get("keys", [])
        k_str = ", ".join(str(k) for k in keys[:5])
        if len(keys) > 5:
            k_str += f" (+{len(keys)-5})"
        text = f'<strong>[KV]</strong> {escape(k_str)}'
    elif t == "list":
        count = sec.get("count", 0)
        ordered = "OL" if sec.get("ordered") else "UL"
        text = f'<strong>[{ordered}]</strong> {count} items'
    elif t == "text":
        preview = escape(sec.get("preview", "")[:60])
        text = f'<strong>[TEXT]</strong> {preview}...'
    elif t == "image":
        alt = escape(sec.get("alt", ""))
        text = f'<strong>[IMG]</strong> {alt}'
    else:
        text = f'<strong>[{escape(t.upper())}]</strong>'

    # Children
    children = sec.get("children", [])
    if children:
        child_html = "<ul>" + "".join(f"<li>{_outline_to_html(c)}</li>" for c in children) + "</ul>"
        text += child_html

    return text


def _index_style() -> str:
    return """<style>
body { font-family: 'Hiragino Kaku Gothic ProN', 'Meiryo', sans-serif; max-width: 1100px; margin: 2em auto; padding: 0 1em; color: #333; }
h1 { border-bottom: 2px solid #333; padding-bottom: 0.3em; }
h2 { color: #555; border-bottom: 1px solid #ccc; padding-bottom: 0.2em; margin-top: 2em; }
.meta { color: #888; font-size: 0.9em; }
table { border-collapse: collapse; width: 100%; margin: 1em 0; }
th, td { border: 1px solid #ccc; padding: 0.5em; text-align: left; }
th { background-color: #f5f5f5; font-weight: bold; }
tfoot td { background-color: #fafafa; }
a { color: #2563eb; text-decoration: none; }
a:hover { text-decoration: underline; }
.json-link { font-size: 0.8em; color: #888; }
.badge { display: inline-block; padding: 0.15em 0.5em; border-radius: 3px; font-size: 0.8em; margin: 0.1em; }
.badge-heading { background: #dbeafe; color: #1e40af; }
.badge-table { background: #dcfce7; color: #166534; }
.badge-key_value { background: #fef3c7; color: #92400e; }
.badge-text { background: #f3e8ff; color: #6b21a8; }
.badge-list { background: #ffedd5; color: #9a3412; }
.badge-image { background: #fce7f3; color: #9d174d; }
details { margin: 0.5em 0; padding: 0.5em; border: 1px solid #eee; border-radius: 4px; }
details summary { cursor: pointer; }
details h3 { margin: 0.5em 0 0.2em; font-size: 0.95em; color: #555; }
details ul { margin: 0.2em 0; padding-left: 1.5em; font-size: 0.9em; }
.footer { margin-top: 3em; padding-top: 1em; border-top: 1px solid #eee; color: #aaa; font-size: 0.85em; }
</style>"""


# ─── Top-level entry point ─────────────────────────────────────────────

def write_index(entries: list[FileEntry], output_dir: Path) -> tuple[Path, Path]:
    """Write both index.html and manifest.json. Returns (index_path, manifest_path)."""
    manifest_path = write_manifest(entries, output_dir)
    index_path = write_index_html(entries, output_dir)
    return index_path, manifest_path
