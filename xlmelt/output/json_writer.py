"""JSON output writer."""

from __future__ import annotations

import json
from pathlib import Path

from ..core.model import DocumentModel


class JsonWriter:
    """Write document model as JSON."""

    def __init__(self, indent: int = 2, ensure_ascii: bool = False):
        self.indent = indent
        self.ensure_ascii = ensure_ascii

    def write(self, doc: DocumentModel, output_path: Path) -> Path:
        """Write document model to a JSON file.

        Returns the path to the written file.
        """
        output_path.parent.mkdir(parents=True, exist_ok=True)
        data = doc.to_dict()
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=self.indent, ensure_ascii=self.ensure_ascii)
        return output_path

    def to_string(self, doc: DocumentModel) -> str:
        """Convert document model to JSON string."""
        data = doc.to_dict()
        return json.dumps(data, indent=self.indent, ensure_ascii=self.ensure_ascii)
