"""Shared text read helpers for Windows mixed UTF-8 / ANSI (GBK) sources."""

from __future__ import annotations

import io
import locale
from pathlib import Path


def read_text_path(path: Path, encoding: str | None = None) -> str:
    """Read text file: explicit encoding, or UTF-8 / UTF-8-BOM first, then GBK family.

    ``encoding`` — force codec name (e.g. ``\"utf-8\"``, ``\"gbk\"``).
    ``None`` — auto-detect for CN/EN mixed Windows repos.
    """
    raw = path.read_bytes()
    if encoding:
        return raw.decode(encoding)

    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig")

    for enc in ("utf-8",):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass

    for enc in ("gbk", "gb2312", "cp936"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass

    pref = locale.getpreferredencoding(False)
    if pref and pref.lower() not in {"utf-8", "utf8"}:
        try:
            return raw.decode(pref)
        except (UnicodeDecodeError, LookupError):
            pass

    return raw.decode("utf-8", errors="replace")


def decode_process_stderr(data: bytes) -> str:
    """Decode Pandoc/tool stderr on Windows (may be UTF-8 or system OEM/ANSI)."""
    if not data:
        return ""
    for enc in ("utf-8", "utf-8-sig"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            pass
    pref = locale.getpreferredencoding(False)
    if pref:
        try:
            return data.decode(pref)
        except (UnicodeDecodeError, LookupError):
            pass
    for enc in ("gbk", "cp936"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            pass
    return data.decode("utf-8", errors="replace")


def reconfigure_stdio_utf8() -> None:
    """Avoid UnicodeEncodeError when printing --help / logs on GBK Windows consoles."""
    import sys

    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError, ValueError, io.UnsupportedOperation):
            pass
