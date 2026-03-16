"""HTML output writer."""

from __future__ import annotations

from html import escape
from pathlib import Path

from ..core.model import DocumentModel, Section, SectionType


class HtmlWriter:
    """Write document model as semantic HTML."""

    def __init__(self, include_style: bool = True):
        self.include_style = include_style

    def write(self, doc: DocumentModel, output_path: Path) -> Path:
        """Write document model to an HTML file."""
        output_path.parent.mkdir(parents=True, exist_ok=True)
        html = self.to_string(doc)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html)
        return output_path

    def to_string(self, doc: DocumentModel) -> str:
        """Convert document model to HTML string."""
        parts: list[str] = []
        parts.append("<!DOCTYPE html>")
        parts.append('<html lang="ja">')
        parts.append("<head>")
        parts.append('<meta charset="utf-8">')
        parts.append(f"<title>{escape(doc.title)}</title>")
        if self.include_style:
            parts.append(self._default_style())
        parts.append("</head>")
        parts.append("<body>")
        parts.append(f"<h1>{escape(doc.title)}</h1>")
        parts.append(f'<p class="source">Source: {escape(doc.source_file)}</p>')

        for sheet in doc.sheets:
            if len(doc.sheets) > 1:
                parts.append(f'<section class="sheet">')
                parts.append(f"<h2>{escape(sheet.name)}</h2>")
            sections = sheet.sections
            idx = 0
            while idx < len(sections):
                section = sections[idx]
                # Check for side-by-side table group
                lg = (section.content or {}).get("layout_group") if section.type == SectionType.TABLE else None
                if lg:
                    # Collect all consecutive tables with the same layout_group
                    group: list[Section] = [section]
                    while idx + 1 < len(sections) and sections[idx + 1].type == SectionType.TABLE \
                            and (sections[idx + 1].content or {}).get("layout_group") == lg:
                        idx += 1
                        group.append(sections[idx])
                    parts.append('<div class="table-group">')
                    for tbl in group:
                        parts.append(self._render_table(tbl))
                    parts.append("</div>")
                else:
                    parts.append(self._render_section(section))
                idx += 1
            if len(doc.sheets) > 1:
                parts.append("</section>")

        parts.append("</body>")
        parts.append("</html>")
        return "\n".join(parts)

    def _render_section(self, section: Section) -> str:
        """Render a section to HTML."""
        if section.type == SectionType.HEADING:
            return self._render_heading(section)
        elif section.type == SectionType.TABLE:
            return self._render_table(section)
        elif section.type == SectionType.KEY_VALUE:
            return self._render_key_value(section)
        elif section.type == SectionType.TEXT:
            return self._render_text(section)
        elif section.type == SectionType.LIST:
            return self._render_list(section)
        elif section.type == SectionType.IMAGE:
            return self._render_image(section)
        else:
            return self._render_text(section)

    def _render_heading(self, section: Section) -> str:
        level = min(max(section.level, 1), 6)
        text = escape(section.title or "")
        return f"<h{level}>{text}</h{level}>"

    def _render_table(self, section: Section) -> str:
        if not section.content:
            return ""
        parts: list[str] = []
        if section.title:
            parts.append(f'<p class="table-title">{escape(section.title)}</p>')
        parts.append("<table>")

        header_rows = section.content.get("header_rows")
        headers = section.content.get("headers")
        rows = section.content.get("rows", [])

        # Render multi-level headers if available, otherwise flat headers
        if header_rows:
            parts.append("<thead>")
            for hr in header_rows:
                parts.append("<tr>")
                for cell in hr:
                    attrs = self._merge_attrs(cell)
                    val = self._cell_value(cell)
                    parts.append(f"<th{attrs}>{self._escape_cell(val)}</th>")
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
                    # Covered by a previous cell's rowspan/colspan — skip
                    continue
                attrs = self._merge_attrs(cell)
                val = self._cell_value(cell)
                parts.append(f"<td{attrs}>{self._escape_cell(val)}</td>")
            parts.append("</tr>")
        parts.append("</tbody>")
        parts.append("</table>")
        return "\n".join(parts)

    def _merge_attrs(self, cell) -> str:
        """Build colspan/rowspan/style HTML attributes from a cell dict."""
        if not isinstance(cell, dict):
            return ""
        attrs = ""
        if cell.get("colspan", 1) > 1:
            attrs += f' colspan="{cell["colspan"]}"'
        if cell.get("rowspan", 1) > 1:
            attrs += f' rowspan="{cell["rowspan"]}"'
        style = self._cell_style_css(cell)
        if style:
            attrs += f' style="{escape(style)}"'
        return attrs

    @staticmethod
    def _cell_style_css(cell) -> str:
        """Build inline CSS from cell style dict."""
        if not isinstance(cell, dict):
            return ""
        style = cell.get("style")
        if not style or not isinstance(style, dict):
            return ""
        parts: list[str] = []
        bg = style.get("bg")
        if bg:
            # Normalize color: ensure # prefix
            color = bg if bg.startswith("#") else f"#{bg}"
            parts.append(f"background-color: {color}")
        color = style.get("color")
        if color:
            color = color if color.startswith("#") else f"#{color}"
            parts.append(f"color: {color}")
        if style.get("bold") == "true":
            parts.append("font-weight: bold")
        return "; ".join(parts)

    def _cell_value(self, cell) -> str:
        """Extract display value from a cell (string or dict)."""
        if isinstance(cell, dict):
            return str(cell.get("value", ""))
        return str(cell)

    @staticmethod
    def _escape_cell(val: str) -> str:
        """Escape cell value for HTML, converting newlines to <br>."""
        escaped = escape(val)
        return escaped.replace("\n", "<br>")

    def _render_key_value(self, section: Section) -> str:
        if not section.content or not isinstance(section.content, dict):
            return ""
        parts: list[str] = ["<dl>"]
        for key, value in section.content.items():
            parts.append(f"<dt>{escape(str(key))}</dt>")
            parts.append(f"<dd>{self._escape_cell(str(value))}</dd>")
        parts.append("</dl>")
        return "\n".join(parts)

    def _render_text(self, section: Section) -> str:
        if not section.content:
            return ""
        text = str(section.content)
        paragraphs = text.split("\n")
        parts = [f"<p>{escape(p)}</p>" for p in paragraphs if p.strip()]
        return "\n".join(parts)

    def _render_list(self, section: Section) -> str:
        if not section.content:
            return ""
        # New format: {"ordered": bool, "items": [str]}
        if isinstance(section.content, dict):
            is_ordered = section.content.get("ordered", False)
            items = section.content.get("items", [])
        elif isinstance(section.content, list):
            # Legacy format: plain list of strings
            is_ordered = False
            items = section.content
        else:
            items = [section.content]
            is_ordered = False

        tag = "ol" if is_ordered else "ul"
        parts = [f"<{tag}>"]
        for item in items:
            parts.append(f"<li>{escape(str(item))}</li>")
        parts.append(f"</{tag}>")
        return "\n".join(parts)

    def _render_image(self, section: Section) -> str:
        if not section.content:
            return ""
        src = section.content.get("path", "")
        alt = section.content.get("alt", "")
        if not src:
            # Chart placeholder (no image file)
            return f'<figure class="chart-placeholder"><p>[{escape(alt)}]</p></figure>'
        return f'<figure><img src="{escape(src)}" alt="{escape(alt)}"></figure>'

    def _default_style(self) -> str:
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
