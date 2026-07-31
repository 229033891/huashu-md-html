#!/usr/bin/env python3
"""
any_to_md.py — Convert any file to Markdown using Microsoft markitdown.

Supports: PDF, DOCX, PPTX, XLSX, XLS, HTML, CSV, JSON, XML, EPub, ZIP,
images (EXIF + optional LLM description), audio (with transcription),
YouTube URLs (with auto subtitles), Outlook .msg, and more.

Local Excel (.xlsx/.xls/.xlsm): post-processes markdown tables (NaN/<NA>/Unnamed:n → empty,
drops columns that are entirely empty, trims trailing all-empty columns).

Part of huashu-md-html skill — md is source, html is product.
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path

from _encoding_io import reconfigure_stdio_utf8

HELP_INSTALL = """
markitdown is not installed. Install it with:

    python -m pip install 'markitdown[all]'
    # or: python3 -m pip install 'markitdown[all]'

Or, for a slimmer install (only the formats you need):

    python -m pip install 'markitdown[pdf,docx,pptx,xlsx]'
"""


def ensure_markitdown():
    try:
        from markitdown import MarkItDown  # noqa: F401
    except ImportError:
        sys.stderr.write(HELP_INSTALL)
        sys.exit(2)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Convert any file (or URL) to Markdown via Microsoft markitdown.",
    )
    p.add_argument(
        "source",
        help="File path or URL (http(s)://, file://, data:, YouTube URL).",
    )
    p.add_argument(
        "-o", "--output",
        help="Output .md path. If omitted, writes to <source-stem>.md in CWD; use '-' for stdout.",
    )
    p.add_argument(
        "--llm-describe",
        action="store_true",
        help="Enable LLM-based image description. Requires OPENAI_API_KEY.",
    )
    p.add_argument(
        "--llm-model",
        default=os.environ.get("MARKITDOWN_LLM_MODEL", "gpt-4o"),
        help="LLM model for image description (default: gpt-4o, or env MARKITDOWN_LLM_MODEL).",
    )
    p.add_argument(
        "--azure-doc-intel",
        default=os.environ.get("AZURE_DOC_INTEL_ENDPOINT"),
        help="Azure Document Intelligence endpoint for high-fidelity PDF OCR.",
    )
    p.add_argument(
        "--enable-plugins",
        action="store_true",
        help="Enable third-party markitdown plugins.",
    )
    p.add_argument(
        "--quiet",
        action="store_true",
        help="Suppress non-error stderr output.",
    )
    p.add_argument(
        "--no-clean-excel-tables",
        action="store_true",
        help="For .xlsx/.xls/.xlsm files, skip post-clean (NaN→empty, drop all-empty columns).",
    )
    return p.parse_args()


def build_converter(args: argparse.Namespace):
    from markitdown import MarkItDown

    kwargs = {"enable_plugins": args.enable_plugins}

    if args.llm_describe:
        try:
            from openai import OpenAI
        except ImportError:
            sys.stderr.write(
                "--llm-describe requires the 'openai' package: python -m pip install openai\n",
            )
            sys.exit(2)
        if not os.environ.get("OPENAI_API_KEY"):
            sys.stderr.write(
                "--llm-describe requires OPENAI_API_KEY environment variable.\n",
            )
            sys.exit(2)
        kwargs["llm_client"] = OpenAI()
        kwargs["llm_model"] = args.llm_model

    if args.azure_doc_intel:
        kwargs["docintel_endpoint"] = args.azure_doc_intel

    return MarkItDown(**kwargs)


def resolve_output_path(source: str, output: str | None) -> Path | None:
    if output == "-":
        return None
    if output:
        return Path(output)
    # Default: <source-stem>.md in CWD
    if source.startswith(("http://", "https://", "data:", "file://")):
        # URL → use a generic name
        return Path("converted.md")
    return Path(Path(source).stem + ".md")


_EXCEL_SUFFIXES = frozenset({".xlsx", ".xls", ".xlsm"})


def _is_local_excel_path(source: str) -> bool:
    if "://" in source:
        return False
    try:
        return Path(source).suffix.lower() in _EXCEL_SUFFIXES
    except OSError:
        return False


def _clean_table_cell(text: str) -> str:
    s = text.strip()
    if s in {"NaN", "<NA>", "nan"}:
        return ""
    if re.fullmatch(r"Unnamed:\s*\d+", s):
        return ""
    return text.strip()


def _is_markdown_table_row(line: str) -> bool:
    s = line.strip()
    return len(s) >= 2 and s.startswith("|") and s.endswith("|")


def _parse_table_row(line: str) -> list[str]:
    inner = line.strip()[1:-1]
    return [_clean_table_cell(c) for c in inner.split("|")]


def _normalize_row_width(rows: list[list[str]], width: int) -> None:
    for r in rows:
        while len(r) < width:
            r.append("")
        del r[width:]


def _is_separator_row(row: list[str]) -> bool:
    """True for GFM separator rows like ``| --- | :---: |`` (must not count as column content)."""
    if not row:
        return False
    return all((not c.strip()) or re.fullmatch(r":?-{3,}:?", c.strip()) for c in row)


def _col_has_content(rows: list[list[str]], col: int) -> bool:
    """Whether any non-separator data/header cell in this column is non-empty."""
    for r in rows:
        if _is_separator_row(r):
            continue
        if (r[col] if col < len(r) else "").strip():
            return True
    return False


def _clean_markdown_tables(md: str) -> str:
    """Normalize MarkItDown/pandas-style Excel tables: drop NaN/Unnamed noise, remove empty columns."""
    lines = md.splitlines()
    out: list[str] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if _is_markdown_table_row(line):
            block: list[str] = []
            while i < len(lines) and _is_markdown_table_row(lines[i]):
                block.append(lines[i])
                i += 1
            rows = [_parse_table_row(row) for row in block]
            if not rows:
                out.extend(block)
                continue
            width = max(len(r) for r in rows)
            _normalize_row_width(rows, width)
            keep_cols = [j for j in range(width) if _col_has_content(rows, j)]
            if not keep_cols:
                out.extend(block)
                continue
            slim_rows = [[r[j] for j in keep_cols] for r in rows]
            # Drop trailing empty columns again (ignore separator rows)
            while slim_rows and slim_rows[0] and not any(
                (r[-1].strip() if r and not _is_separator_row(r) else False) for r in slim_rows
            ):
                for row in slim_rows:
                    if row:
                        row.pop()
            for slim in slim_rows:
                out.append("| " + " | ".join(slim) + " |")
            continue
        out.append(line)
        i += 1
    return "\n".join(out)


def warn_known_pitfalls(source: str, content: str, quiet: bool) -> None:
    if quiet:
        return
    suffix = Path(source).suffix.lower() if "://" not in source else ""
    if suffix == ".pdf" and len(content.strip()) < 200:
        sys.stderr.write(
            "[hint] Output is very short — this PDF may be a scanned document.\n"
            "       Try --llm-describe or --azure-doc-intel for OCR.\n",
        )
    if suffix in {".pptx", ".ppt"}:
        sys.stderr.write(
            "[hint] PPTX text + speaker notes preserved; animations and layout are dropped.\n",
        )


def main() -> int:
    reconfigure_stdio_utf8()
    ensure_markitdown()
    args = parse_args()

    converter = build_converter(args)

    try:
        result = converter.convert(args.source)
    except Exception as exc:  # noqa: BLE001 — markitdown wraps various errors
        sys.stderr.write(f"[error] markitdown failed: {exc}\n")
        return 1

    content = result.text_content or ""
    if (
        not args.no_clean_excel_tables
        and _is_local_excel_path(args.source)
        and content.strip()
    ):
        content = _clean_markdown_tables(content)

    out_path = resolve_output_path(args.source, args.output)
    if out_path is None:
        sys.stdout.write(content)
    else:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(content, encoding="utf-8")
        if not args.quiet:
            sys.stderr.write(f"[ok] {args.source} → {out_path}\n")

    warn_known_pitfalls(args.source, content, args.quiet)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
