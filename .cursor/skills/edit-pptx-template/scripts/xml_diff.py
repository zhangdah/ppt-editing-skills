"""
Diff two .pptx files at the OOXML level. Use this AFTER xml_patch.py to verify
that only the slides / runs / attributes you intended were touched.

Two output modes:

    (default)   one-line summary per changed zip member, plus a short XML
                preview of what changed inside text-bearing parts.
    --full      print the unified text diff for every changed XML/text part.

Usage:
    python3 xml_diff.py before.pptx after.pptx
    python3 xml_diff.py before.pptx after.pptx --full
    python3 xml_diff.py before.pptx after.pptx --only ppt/slides/

Pure standard library.
"""

from __future__ import annotations

import argparse
import difflib
import io
import os
import sys
import xml.dom.minidom
import zipfile
from typing import Iterable


# Members we treat as text/XML for diff purposes. (Most things in a .pptx are
# XML; this list is a positive filter to avoid trying to text-diff PNGs etc.)
TEXTUAL_EXTS = (".xml", ".rels", ".vml", ".txt", ".json")


def is_textual(name: str) -> bool:
    return name.endswith(TEXTUAL_EXTS)


def list_members(path: str) -> dict[str, bytes]:
    out: dict[str, bytes] = {}
    with zipfile.ZipFile(path) as zf:
        for info in zf.infolist():
            out[info.filename] = zf.read(info.filename)
    return out


def pretty_xml(raw: bytes) -> str:
    """Pretty-print XML for human-readable diffs. Falls back to raw text for
    non-XML payloads."""
    try:
        dom = xml.dom.minidom.parseString(raw)
        return dom.toprettyxml(indent="  ")
    except Exception:
        try:
            return raw.decode("utf-8")
        except UnicodeDecodeError:
            return raw.decode("latin-1", errors="replace")


def member_diff(name: str, a: bytes, b: bytes,
                full: bool, max_preview_lines: int = 12) -> str:
    """Return a printable diff string for one changed zip member."""
    if not is_textual(name):
        return f"  [binary] {name}: {len(a)} -> {len(b)} bytes"

    a_text = pretty_xml(a)
    b_text = pretty_xml(b)
    diff_lines = list(difflib.unified_diff(
        a_text.splitlines(keepends=False),
        b_text.splitlines(keepends=False),
        fromfile=f"a/{name}", tofile=f"b/{name}",
        lineterm="",
    ))
    if not diff_lines:
        return f"  [text]   {name}: bytes differ but pretty-XML matches " \
               f"(likely whitespace/attr-order only)"

    if full:
        return "\n".join(diff_lines)

    # Show only +/- lines, capped, for the summary mode.
    changed = [ln for ln in diff_lines if ln.startswith(("+", "-"))
               and not ln.startswith(("+++", "---"))]
    preview = changed[:max_preview_lines]
    more = len(changed) - len(preview)
    head = f"  [xml]    {name}: {len(changed)} changed line(s)"
    body = "\n".join(f"      {line}" for line in preview)
    tail = f"\n      ... (+{more} more lines)" if more > 0 else ""
    return f"{head}\n{body}{tail}"


def diff_decks(path_a: str, path_b: str, only_prefix: str | None,
               full: bool) -> int:
    a = list_members(path_a)
    b = list_members(path_b)

    a_keys = set(a)
    b_keys = set(b)

    added = sorted(b_keys - a_keys)
    removed = sorted(a_keys - b_keys)
    common = sorted(a_keys & b_keys)
    changed = [k for k in common if a[k] != b[k]]

    if only_prefix:
        added = [k for k in added if k.startswith(only_prefix)]
        removed = [k for k in removed if k.startswith(only_prefix)]
        changed = [k for k in changed if k.startswith(only_prefix)]

    print(f"Deck A: {os.path.abspath(path_a)} ({len(a_keys)} parts)")
    print(f"Deck B: {os.path.abspath(path_b)} ({len(b_keys)} parts)")
    print(f"  added:   {len(added)}")
    print(f"  removed: {len(removed)}")
    print(f"  changed: {len(changed)}")
    if only_prefix:
        print(f"  filter:  {only_prefix}")

    if not (added or removed or changed):
        print("\n[ OK ] decks are byte-identical (within filter).")
        return 0

    if added:
        print("\n--- added ---")
        for name in added:
            print(f"  +{name} ({len(b[name])} bytes)")
    if removed:
        print("\n--- removed ---")
        for name in removed:
            print(f"  -{name} ({len(a[name])} bytes)")
    if changed:
        print("\n--- changed ---")
        for name in changed:
            print(member_diff(name, a[name], b[name], full=full))
            if full:
                print()  # blank line between full diffs

    return 1


def main(argv: Iterable[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="OOXML-level diff between two .pptx files.")
    ap.add_argument("before", help="original .pptx")
    ap.add_argument("after", help="edited .pptx")
    ap.add_argument("--only",
                    help="restrict diff to members starting with this prefix, "
                         "e.g. 'ppt/slides/'")
    ap.add_argument("--full", action="store_true",
                    help="print full unified diff for every changed XML part")
    args = ap.parse_args(list(argv) if argv is not None else None)

    for p in (args.before, args.after):
        if not os.path.exists(p):
            print(f"[FAIL] not found: {p}", file=sys.stderr)
            return 1
        if not zipfile.is_zipfile(p):
            print(f"[FAIL] not a valid .pptx (zip): {p}", file=sys.stderr)
            return 1

    return diff_decks(args.before, args.after,
                      only_prefix=args.only, full=args.full)


if __name__ == "__main__":
    sys.exit(main())
