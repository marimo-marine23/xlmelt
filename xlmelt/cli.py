"""CLI entry point for xlmelt."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import click

from . import __version__
from .core.analyzer import StructureAnalyzer
from .output.html_writer import HtmlWriter
from .output.json_writer import JsonWriter


@click.group()
@click.version_option(version=__version__)
def main() -> None:
    """xlmelt - Convert complex Excel files into AI-readable JSON/HTML."""
    pass


@main.command()
@click.argument("input_path", type=click.Path(exists=True))
@click.option("-o", "--output", "output_dir", type=click.Path(), default="./output",
              help="Output directory.")
@click.option("--format", "output_format", type=click.Choice(["json", "html", "both"]),
              default="both", help="Output format.")
@click.option("--images", "image_mode", type=click.Choice(["extract", "skip"]),
              default="extract", help="Image handling mode.")
@click.option("--no-style", is_flag=True, help="Omit CSS from HTML output.")
@click.option("--stdout", is_flag=True, help="Write output to stdout instead of files (single file only).")
def convert(
    input_path: str,
    output_dir: str,
    output_format: str,
    image_mode: str,
    no_style: bool,
    stdout: bool,
) -> None:
    """Convert Excel file(s) to JSON/HTML."""
    input_path_obj = Path(input_path)
    output_dir_obj = Path(output_dir)

    if stdout:
        if not input_path_obj.is_file():
            click.echo("Error: --stdout requires a single file input", err=True)
            raise SystemExit(1)
        if output_format == "both":
            click.echo("Error: --stdout requires --format json or --format html (not both)", err=True)
            raise SystemExit(1)
        _convert_stdout(input_path_obj, output_format, no_style)
        return

    if input_path_obj.is_file():
        _convert_file(input_path_obj, output_dir_obj, output_format, image_mode, no_style)
    elif input_path_obj.is_dir():
        xlsx_files = [
            f for f in (
                list(input_path_obj.glob("*.xlsx"))
                + list(input_path_obj.glob("*.xlsm"))
                + list(input_path_obj.glob("*.xls"))
            )
            if not f.name.startswith("~$")
        ]
        if not xlsx_files:
            click.echo(f"No Excel files found in {input_path}")
            return
        click.echo(f"Found {len(xlsx_files)} Excel file(s)")
        # Detect stem collisions (e.g. report.xlsx + report.xlsm)
        stem_counts: dict[str, int] = {}
        for f in xlsx_files:
            stem_counts[f.stem] = stem_counts.get(f.stem, 0) + 1
        colliding_stems = {s for s, c in stem_counts.items() if c > 1}
        entries = []
        for f in sorted(xlsx_files):
            # Use full filename as dir name when stems collide to prevent overwrite
            stem_override = f.name if f.stem in colliding_stems else None
            entry = _convert_file(f, output_dir_obj, output_format, image_mode, no_style, stem_override=stem_override)
            if entry is not None:
                entries.append(entry)
        # Generate index for batch conversions
        if len(entries) >= 2:
            from .output.index_writer import write_index
            index_path, manifest_path = write_index(entries, output_dir_obj)
            click.echo(f"  -> {index_path}")
            click.echo(f"  -> {manifest_path}")
    else:
        click.echo(f"Error: {input_path} is not a file or directory", err=True)
        raise SystemExit(1)


def _convert_stdout(
    file_path: Path,
    output_format: str,
    no_style: bool,
) -> None:
    """Convert a single Excel file and write to stdout."""
    analyzer = StructureAnalyzer()
    try:
        doc = analyzer.analyze(file_path)
    except Exception as e:
        click.echo(f"Error: {e}", err=True)
        raise SystemExit(1)

    if output_format == "json":
        writer = JsonWriter()
        click.echo(writer.to_string(doc))
    elif output_format == "html":
        writer_html = HtmlWriter(include_style=not no_style)
        click.echo(writer_html.to_string(doc))


def _convert_file(
    file_path: Path,
    output_dir: Path,
    output_format: str,
    image_mode: str,
    no_style: bool,
    stem_override: str | None = None,
) -> Any:
    """Convert a single Excel file.

    Returns a FileEntry for index generation, or None on error.
    """
    from .output.index_writer import FileEntry, build_entry_from_doc

    click.echo(f"Converting: {file_path.name}")

    # Create output folder named after the file
    dir_name = stem_override if stem_override else file_path.stem
    file_output_dir = output_dir / dir_name

    analyzer = StructureAnalyzer()
    extract_images = image_mode == "extract"
    try:
        doc = analyzer.analyze(file_path, file_output_dir if extract_images else None)
    except Exception as e:
        click.echo(f"  Error: {e}", err=True)
        return None

    if output_format in ("json", "both"):
        writer = JsonWriter()
        json_path = file_output_dir / "document.json"
        writer.write(doc, json_path)
        click.echo(f"  -> {json_path}")

    if output_format in ("html", "both"):
        writer_html = HtmlWriter(include_style=not no_style)
        html_path = file_output_dir / "document.html"
        writer_html.write(doc, html_path)
        click.echo(f"  -> {html_path}")

    # Write metadata
    from datetime import datetime, timezone

    sheet_details = []
    for sheet in doc.sheets:
        type_counts: dict[str, int] = {}
        for s in sheet.sections:
            t = s.type.value
            type_counts[t] = type_counts.get(t, 0) + 1
        sheet_details.append({
            "name": sheet.name,
            "sections": len(sheet.sections),
            "section_types": type_counts,
        })

    metadata = {
        "xlmelt_version": __version__,
        "schema_version": 1,
        "converted_at": datetime.now(timezone.utc).isoformat(),
        "source": file_path.name,
        "format": output_format,
        "sheets": sheet_details,
        "total_sections": sum(len(s.sections) for s in doc.sheets),
        "total_images": len(doc.images),
    }
    meta_path = file_output_dir / "metadata.json"
    meta_path.parent.mkdir(parents=True, exist_ok=True)
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=2, ensure_ascii=False)

    click.echo(f"  Done ({len(doc.sheets)} sheet(s), {len(doc.images)} image(s))")

    # Build entry for index generation
    return build_entry_from_doc(dir_name, file_path.name, doc)


@main.command()
@click.argument("input_path", type=click.Path(exists=True))
@click.option("--json", "as_json", is_flag=True, help="Output as machine-readable JSON.")
def inspect(input_path: str, as_json: bool) -> None:
    """Show document structure as a tree."""
    file_path = Path(input_path)
    if not file_path.is_file():
        click.echo("Error: Please specify an Excel file", err=True)
        raise SystemExit(1)

    analyzer = StructureAnalyzer()
    try:
        doc = analyzer.analyze(file_path)
    except Exception as e:
        click.echo(f"Error: {e}", err=True)
        raise SystemExit(1)

    if as_json:
        outline = {
            "title": doc.title,
            "source": doc.source_file,
            "sheets": [],
        }
        for sheet in doc.sheets:
            sheet_outline: dict = {
                "name": sheet.name,
                "sections": [_section_outline(s) for s in sheet.sections],
            }
            outline["sheets"].append(sheet_outline)
        click.echo(json.dumps(outline, indent=2, ensure_ascii=False))
        return

    click.echo(f"Document: {doc.title}")
    click.echo(f"Source: {doc.source_file}")
    click.echo()

    for sheet in doc.sheets:
        click.echo(f"Sheet: {sheet.name}")
        for section in sheet.sections:
            _print_section(section, indent=1)
        click.echo()


def _section_outline(section) -> dict:
    """Create a compact outline dict for a section (for --json output)."""
    from .output.index_writer import section_outline
    return section_outline(section)


def _print_section(section, indent: int = 0) -> None:
    """Print a section as a tree node."""
    prefix = "  " * indent
    type_label = section.type.value.upper()

    if section.type.value == "heading":
        click.echo(f"{prefix}[H{section.level}] {section.title}")
    elif section.type.value == "table":
        rows = section.content.get("rows", []) if section.content else []
        header_rows = section.content.get("header_rows", []) if section.content else []
        headers = section.content.get("headers", []) if section.content else []
        if header_rows and len(header_rows) > 1:
            header_info = f" header_levels={len(header_rows)} headers={headers}"
        elif headers:
            header_info = f" headers={headers}"
        else:
            header_info = ""
        click.echo(f"{prefix}[TABLE] {len(rows)} rows{header_info}")
    elif section.type.value == "key_value":
        pairs = section.content if isinstance(section.content, dict) else {}
        click.echo(f"{prefix}[KV] {len(pairs)} pairs")
        for k, v in pairs.items():
            click.echo(f"{prefix}  {k}: {v}")
    elif section.type.value == "list":
        if isinstance(section.content, dict):
            items = section.content.get("items", [])
            list_type = "OL" if section.content.get("ordered") else "UL"
        elif isinstance(section.content, list):
            items = section.content
            list_type = "UL"
        else:
            items = []
            list_type = "UL"
        click.echo(f"{prefix}[{list_type}] {len(items)} items")
        marker = "  " if list_type == "OL" else "  - "
        for i, item in enumerate(items[:5], 1):
            if list_type == "OL":
                click.echo(f"{prefix}  {i}. {item}")
            else:
                click.echo(f"{prefix}  - {item}")
        if len(items) > 5:
            click.echo(f"{prefix}  ... ({len(items) - 5} more)")
    elif section.type.value == "text":
        text = str(section.content or "")
        preview = text[:60].replace("\n", " ")
        if len(text) > 60:
            preview += "..."
        click.echo(f"{prefix}[TEXT] {preview}")
    else:
        click.echo(f"{prefix}[{type_label}]")

    for child in section.children:
        _print_section(child, indent + 1)


@main.command()
@click.argument("target", type=click.Path(exists=True))
@click.option("--report", "report_path", type=click.Path(),
              help="Write detailed report to file (.md or .txt).")
@click.option("--xlsx", "xlsx_dir", type=click.Path(exists=True),
              help="Directory containing original .xlsx files for coverage check.")
def verify(target: str, report_path: str | None, xlsx_dir: str | None) -> None:
    """Verify JSON↔HTML consistency and xlsx cell coverage.

    TARGET can be:
      - An output directory containing document.json and document.html
      - A parent directory containing output subdirectories
      - An Excel file (will convert first, then verify)
      - A directory of Excel files (verifies each)

    Use --xlsx to specify where original .xlsx files are located for coverage check
    when verifying output directories.
    """
    from .verify import generate_report, verify_file

    target_path = Path(target)
    xlsx_dir_path = Path(xlsx_dir) if xlsx_dir else None
    all_results: list = []

    if target_path.is_file() and target_path.suffix in (".xlsx", ".xlsm", ".xls"):
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_output = Path(tmpdir)
            _convert_file(target_path, tmp_output, "both", "extract", False)
            file_output = tmp_output / target_path.stem
            click.echo()
            click.echo(f"Verifying: {target_path.name}")
            result = verify_file(file_output, xlsx_path=target_path)
            result.name = target_path.stem
            click.echo(result.summary())
            all_results.append(result)
    elif target_path.is_dir():
        if (target_path / "document.json").exists():
            xlsx_path = _find_xlsx_for_output(target_path, xlsx_dir_path)
            click.echo(f"Verifying: {target_path.name}")
            result = verify_file(target_path, xlsx_path=xlsx_path)
            click.echo(result.summary())
            all_results.append(result)
        else:
            # Check for subdirectories containing document.json
            output_dirs = sorted([
                d for d in target_path.iterdir()
                if d.is_dir() and (d / "document.json").exists()
            ])
            if output_dirs:
                for out_dir in output_dirs:
                    xlsx_path = _find_xlsx_for_output(out_dir, xlsx_dir_path)
                    click.echo(f"Verifying: {out_dir.name}")
                    result = verify_file(out_dir, xlsx_path=xlsx_path)
                    click.echo(result.summary())
                    all_results.append(result)
                    click.echo()
            else:
                # Try as directory of Excel files
                xlsx_files = [
                    f for f in (
                        list(target_path.glob("*.xlsx"))
                        + list(target_path.glob("*.xlsm"))
                        + list(target_path.glob("*.xls"))
                    )
                    if not f.name.startswith("~$")
                ]
                if not xlsx_files:
                    click.echo(f"No output directories or Excel files found in {target}")
                    return

                import tempfile
                for xlsx in sorted(xlsx_files):
                    with tempfile.TemporaryDirectory() as tmpdir:
                        tmp_output = Path(tmpdir)
                        _convert_file(xlsx, tmp_output, "both", "extract", False)
                        file_output = tmp_output / xlsx.stem
                        click.echo()
                        click.echo(f"Verifying: {xlsx.name}")
                        result = verify_file(file_output, xlsx_path=xlsx)
                        result.name = xlsx.stem
                        click.echo(result.summary())
                        all_results.append(result)
                        click.echo()
    else:
        click.echo(f"Error: {target} is not a valid target", err=True)
        raise SystemExit(1)

    # Write report if requested
    if report_path and all_results:
        rp = Path(report_path)
        generate_report(all_results, rp)
        click.echo(f"Report written to: {rp}")

    # Exit with error if any failures
    if any(not r.ok for r in all_results):
        raise SystemExit(1)


def _find_xlsx_for_output(output_dir: Path, xlsx_dir: Path | None) -> Path | None:
    """Try to find the original xlsx file for an output directory."""
    # Read metadata.json to get source filename
    meta_path = output_dir / "metadata.json"
    source_name = None
    if meta_path.exists():
        try:
            with open(meta_path, encoding="utf-8") as f:
                meta = json.load(f)
            source_name = meta.get("source")
        except Exception:
            pass

    if not source_name:
        # Guess from directory name
        source_name = output_dir.name + ".xlsx"

    # Sanitize: only use the basename to prevent path traversal
    source_name = Path(source_name).name

    # Search in xlsx_dir if provided
    if xlsx_dir:
        candidate = xlsx_dir / source_name
        if candidate.exists():
            return candidate

    # Search relative to output directory's parent (common layout: output/ and samples/ side by side)
    for search_dir in [output_dir.parent, output_dir.parent.parent]:
        for pattern in [source_name, f"**/{source_name}"]:
            matches = list(search_dir.glob(pattern))
            if matches:
                return matches[0]

    return None


@main.command()
@click.argument("output_dir", type=click.Path(exists=True))
def index(output_dir: str) -> None:
    """Generate index.html and manifest.json for an output directory.

    Scans OUTPUT_DIR for subdirectories containing metadata.json
    and generates an index page linking to all documents plus
    a manifest.json catalog optimized for AI tool consumption.
    """
    from .output.index_writer import build_entry_from_output, write_index

    output_path = Path(output_dir)
    if not output_path.is_dir():
        click.echo(f"Error: {output_dir} is not a directory", err=True)
        raise SystemExit(1)

    # Find all output subdirectories
    output_dirs = sorted([
        d for d in output_path.iterdir()
        if d.is_dir() and (d / "metadata.json").exists()
    ])

    if not output_dirs:
        click.echo(f"No converted output directories found in {output_dir}")
        return

    click.echo(f"Found {len(output_dirs)} document(s)")
    entries = []
    for d in output_dirs:
        entry = build_entry_from_output(d)
        if entry is not None:
            entries.append(entry)
            click.echo(f"  {entry.name} ({len(entry.sheets)} sheets, {entry.total_sections} sections)")

    if entries:
        index_path, manifest_path = write_index(entries, output_path)
        click.echo(f"  -> {index_path}")
        click.echo(f"  -> {manifest_path}")
    else:
        click.echo("No valid entries found")


@main.command()
@click.argument("input_path", type=click.Path(exists=True))
@click.option("--report", "report_path", type=click.Path(),
              help="Write score report to file (.md or .txt).")
@click.option("--json", "as_json", is_flag=True, help="Output as JSON.")
def score(input_path: str, report_path: str | None, as_json: bool) -> None:
    """Score AI-readability of Excel file conversion.

    Measures how much easier the converted JSON/HTML is for AI consumption
    compared to the raw Excel data. Scores range from 0-100.
    """
    from .score import DirectoryScore, FileScore, generate_score_report, score_directory, score_file

    input_path_obj = Path(input_path)
    all_scores: list[FileScore] = []

    if input_path_obj.is_file() and input_path_obj.suffix in (".xlsx", ".xlsm", ".xls"):
        click.echo(f"Scoring: {input_path_obj.name}")
        s = score_file(input_path_obj)
        all_scores.append(s)
    elif input_path_obj.is_dir():
        xlsx_files = [
            f for f in (
                list(input_path_obj.glob("*.xlsx"))
                + list(input_path_obj.glob("*.xlsm"))
                + list(input_path_obj.glob("*.xls"))
            )
            if not f.name.startswith("~$")
        ]
        if not xlsx_files:
            click.echo(f"No Excel files found in {input_path}")
            return
        click.echo(f"Scoring {len(xlsx_files)} file(s)...")
        for f in sorted(xlsx_files):
            click.echo(f"  {f.name}")
            s = score_file(f)
            all_scores.append(s)
    else:
        click.echo(f"Error: {input_path} is not a valid Excel file or directory", err=True)
        raise SystemExit(1)

    # Directory efficiency score (2+ files)
    dir_score: DirectoryScore | None = None
    if len(all_scores) >= 2:
        # Try to find existing manifest.json in the output directory
        manifest_path = None
        if input_path_obj.is_dir():
            # Check common output locations
            for candidate in [
                Path("./output/manifest.json"),
                input_path_obj.parent / "output" / "manifest.json",
            ]:
                if candidate.exists():
                    manifest_path = candidate
                    break
        dir_score = score_directory(all_scores, manifest_path)

    if as_json:
        output: dict[str, Any] = {
            "scores": [s.to_dict() for s in all_scores],
        }
        if all_scores:
            n = len(all_scores)
            output["averages"] = {
                "readability_raw": round(sum(s.readability_raw for s in all_scores) / n, 1),
                "readability_json": round(sum(s.readability_json for s in all_scores) / n, 1),
                "readability_improvement": round(sum(s.readability_improvement for s in all_scores) / n, 1),
                "token_saved_pct": round(sum(s.token_saved_pct for s in all_scores) / n, 1),
                "cell_coverage_pct": round(sum(s.cell_coverage_pct for s in all_scores) / n, 1),
                "overall": round(sum(s.overall for s in all_scores) / n, 1),
            }
        if dir_score:
            output["directory_efficiency"] = dir_score.to_dict()
        click.echo(json.dumps(output, indent=2, ensure_ascii=False))
    else:
        click.echo()
        for s in all_scores:
            click.echo(s.summary())
            click.echo()
        if dir_score:
            click.echo(dir_score.summary())
            click.echo()

    if report_path:
        rp = Path(report_path)
        generate_score_report(all_scores, rp)
        click.echo(f"Report written to: {rp}")


if __name__ == "__main__":
    main()
