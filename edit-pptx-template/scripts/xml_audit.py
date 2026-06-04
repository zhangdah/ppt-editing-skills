"""
Audit a .pptx by reading its slide XML directly.

Outputs JSON describing every text-bearing element with a STABLE ID, the
current text content, and every font attribute that's currently set on that
element. This is the single source of truth shared between humans and agents
when planning text/font edits — both sides see exactly the same fields.

Usage:
    python3 xml_audit.py deck.pptx                  # write deck.audit.json
    python3 xml_audit.py deck.pptx --out a.json     # custom path
    python3 xml_audit.py deck.pptx --slides 1,3-4   # subset of slides
    python3 xml_audit.py deck.pptx --pretty         # pretty-printed JSON

Stable IDs:
    s{slide}                              — slide N (1-based)
    s{slide}.sp{cNvPr_id}                 — shape with non-visual id `cNvPr_id`
    s{slide}.sp{cNvPr_id}.p{para}         — paragraph #para (0-based) inside it
    s{slide}.sp{cNvPr_id}.p{para}.r{run}  — run #run (0-based) inside that paragraph

`cNvPr_id` is taken from the .pptx itself (`<p:cNvPr id="...">`), not from
positional ordering. Adding or removing other shapes will not renumber it.

Pure standard library — no python-pptx, no lxml.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile
from typing import Iterable
from xml.etree import ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pptx_fingerprint import fingerprint, is_stale


# OOXML namespaces we care about.
NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def _q(prefix: str, tag: str) -> str:
    """Return a Clark-notation qualified name, e.g. ('a','t') -> '{...}t'."""
    return f"{{{NS[prefix]}}}{tag}"


# -----------------------------------------------------------------------------
# Slide XML reading
# -----------------------------------------------------------------------------
SLIDE_RE = re.compile(r"^ppt/slides/slide(\d+)\.xml$")


def list_slide_parts(zf: zipfile.ZipFile) -> list[tuple[int, str]]:
    """Return [(slide_index_1based, zip_member_name), ...] in slide order."""
    found: list[tuple[int, str]] = []
    for name in zf.namelist():
        m = SLIDE_RE.match(name)
        if m:
            found.append((int(m.group(1)), name))
    found.sort(key=lambda x: x[0])
    return found


def slide_size_emu(zf: zipfile.ZipFile) -> tuple[int, int]:
    """Return (slide_width_EMU, slide_height_EMU) from presentation.xml."""
    try:
        raw = zf.read("ppt/presentation.xml")
    except KeyError:
        return (0, 0)
    root = ET.fromstring(raw)
    sz = root.find(_q("p", "sldSz"))
    if sz is None:
        return (0, 0)
    cx = int(sz.attrib.get("cx", "0"))
    cy = int(sz.attrib.get("cy", "0"))
    return (cx, cy)


# -----------------------------------------------------------------------------
# Run-property parsing — translate <a:rPr>/<a:pPr>/<a:defRPr> into a flat dict
# -----------------------------------------------------------------------------
def _parse_color(elem: ET.Element | None) -> str | None:
    """Return the hex color of a <...Fill> wrapper, or 'theme:accent1' for
    theme references. None if no fill is set on this element."""
    if elem is None:
        return None
    solid = elem.find(_q("a", "solidFill"))
    if solid is None:
        # Could be noFill / gradFill / pattFill — return the tag name so the
        # caller knows it's not a plain sRGB.
        for child in elem:
            local = child.tag.split("}", 1)[-1]
            if local.endswith("Fill"):
                return f"complex:{local}"
        return None
    srgb = solid.find(_q("a", "srgbClr"))
    if srgb is not None:
        return srgb.attrib.get("val")
    scheme = solid.find(_q("a", "schemeClr"))
    if scheme is not None:
        return f"theme:{scheme.attrib.get('val')}"
    return "complex:solidFill"


def _bool_attr(elem: ET.Element, name: str) -> bool | None:
    """OOXML uses '1'/'0' for booleans. Return True/False or None if absent."""
    v = elem.attrib.get(name)
    if v is None:
        return None
    return v == "1"


def parse_run_properties(rpr: ET.Element | None) -> dict:
    """Return a dict of font properties present on this <a:rPr> (or pPr/defRPr).
    Missing attributes are reported as None — meaning 'inherits from parent'."""
    if rpr is None:
        return {
            "size_pt": None, "bold": None, "italic": None, "underline": None,
            "strike": None, "color": None, "font_name": None,
            "font_name_ea": None, "font_name_cs": None,
        }

    size_raw = rpr.attrib.get("sz")
    size_pt = (int(size_raw) / 100) if size_raw else None

    underline = rpr.attrib.get("u")  # 'sng' / 'dbl' / 'none' / None
    strike = rpr.attrib.get("strike")  # 'noStrike' / 'sngStrike' / 'dblStrike'

    latin = rpr.find(_q("a", "latin"))
    ea = rpr.find(_q("a", "ea"))
    cs = rpr.find(_q("a", "cs"))

    return {
        "size_pt": size_pt,
        "bold": _bool_attr(rpr, "b"),
        "italic": _bool_attr(rpr, "i"),
        "underline": underline,
        "strike": strike,
        "color": _parse_color(rpr),
        "font_name": latin.attrib.get("typeface") if latin is not None else None,
        "font_name_ea": ea.attrib.get("typeface") if ea is not None else None,
        "font_name_cs": cs.attrib.get("typeface") if cs is not None else None,
    }


# -----------------------------------------------------------------------------
# Shape walking
# -----------------------------------------------------------------------------
def _xfrm_bbox_in(spPr: ET.Element | None) -> list[float] | None:
    """Return [left_in, top_in, width_in, height_in] from <a:xfrm>, or None."""
    if spPr is None:
        return None
    xfrm = spPr.find(_q("a", "xfrm"))
    if xfrm is None:
        return None
    off = xfrm.find(_q("a", "off"))
    ext = xfrm.find(_q("a", "ext"))
    if off is None or ext is None:
        return None
    try:
        x = int(off.attrib.get("x", "0"))
        y = int(off.attrib.get("y", "0"))
        cx = int(ext.attrib.get("cx", "0"))
        cy = int(ext.attrib.get("cy", "0"))
    except ValueError:
        return None
    return [round(x / 914400, 3), round(y / 914400, 3),
            round(cx / 914400, 3), round(cy / 914400, 3)]


def _is_text_box(sp: ET.Element) -> bool:
    """Return True if this <p:sp> is a TEXT_BOX (cNvSpPr txBox=1)."""
    cNvSpPr = sp.find(f".//{_q('p','cNvSpPr')}")
    if cNvSpPr is None:
        return False
    return cNvSpPr.attrib.get("txBox") == "1"


def _shape_kind(sp: ET.Element) -> str:
    """Return a human-readable shape kind: 'text_box', 'rect', 'rounded_rect',
    'oval', 'right_arrow', etc., based on <a:prstGeom prst="...">."""
    if _is_text_box(sp):
        return "text_box"
    spPr = sp.find(_q("p", "spPr"))
    if spPr is not None:
        prst = spPr.find(_q("a", "prstGeom"))
        if prst is not None:
            return prst.attrib.get("prst") or "auto"
    return "auto"


def walk_shapes(slide_root: ET.Element) -> list[ET.Element]:
    """Yield every <p:sp> in document order from the slide root.
    Group shapes are flattened (their children are yielded too)."""
    out: list[ET.Element] = []

    def visit(elem: ET.Element) -> None:
        for child in elem:
            local = child.tag.split("}", 1)[-1]
            if local == "sp":
                out.append(child)
            elif local == "grpSp":
                # Recurse into groups.
                visit(child)
            # Pictures (pic), connectors (cxnSp) etc. are intentionally
            # skipped: this skill only audits/edits text-bearing shapes.

    spTree = slide_root.find(f"{_q('p','cSld')}/{_q('p','spTree')}")
    if spTree is not None:
        visit(spTree)
    return out


def shape_to_record(sp: ET.Element, slide_idx: int) -> dict | None:
    """Convert a <p:sp> element into an audit record. Returns None for
    shapes with no text body (we only care about textual content)."""
    nvSpPr = sp.find(_q("p", "nvSpPr"))
    cNvPr = nvSpPr.find(_q("p", "cNvPr")) if nvSpPr is not None else None
    if cNvPr is None:
        return None
    sp_id = cNvPr.attrib.get("id", "0")
    sp_name = cNvPr.attrib.get("name", "")

    spPr = sp.find(_q("p", "spPr"))
    bbox = _xfrm_bbox_in(spPr)
    kind = _shape_kind(sp)

    txBody = sp.find(_q("p", "txBody"))
    paragraphs: list[dict] = []
    if txBody is not None:
        for p_idx, p in enumerate(txBody.findall(_q("a", "p"))):
            pPr = p.find(_q("a", "pPr"))
            align = pPr.attrib.get("algn") if pPr is not None else None
            level = int(pPr.attrib.get("lvl", "0")) if pPr is not None else 0
            defRPr = pPr.find(_q("a", "defRPr")) if pPr is not None else None
            default_font = parse_run_properties(defRPr)

            runs = []
            run_idx = 0
            for child in p:
                local = child.tag.split("}", 1)[-1]
                if local == "r":
                    rPr = child.find(_q("a", "rPr"))
                    t = child.find(_q("a", "t"))
                    runs.append({
                        "id": f"s{slide_idx}.sp{sp_id}.p{p_idx}.r{run_idx}",
                        "text": (t.text or "") if t is not None else "",
                        "font": parse_run_properties(rPr),
                    })
                    run_idx += 1
                elif local == "br":
                    runs.append({
                        "id": f"s{slide_idx}.sp{sp_id}.p{p_idx}.r{run_idx}",
                        "text": "\n",
                        "font": parse_run_properties(child.find(_q("a", "rPr"))),
                        "kind": "line_break",
                    })
                    run_idx += 1

            paragraphs.append({
                "id": f"s{slide_idx}.sp{sp_id}.p{p_idx}",
                "align": align,
                "level": level,
                "default_font": default_font,
                "runs": runs,
            })

    if not paragraphs:
        # Shape has no text — skip it. (This skill doesn't edit non-text shapes.)
        return None

    return {
        "id": f"s{slide_idx}.sp{sp_id}",
        "name": sp_name,
        "kind": kind,
        "bbox_in": bbox,
        "paragraphs": paragraphs,
    }


# -----------------------------------------------------------------------------
# Top-level
# -----------------------------------------------------------------------------
def audit(deck_path: str, slide_filter: set[int] | None = None) -> dict:
    with zipfile.ZipFile(deck_path) as zf:
        sw_emu, sh_emu = slide_size_emu(zf)
        result_slides: list[dict] = []
        for idx, member in list_slide_parts(zf):
            if slide_filter is not None and idx not in slide_filter:
                continue
            raw = zf.read(member)
            root = ET.fromstring(raw)
            shapes_records: list[dict] = []
            for sp in walk_shapes(root):
                rec = shape_to_record(sp, idx)
                if rec is not None:
                    shapes_records.append(rec)
            result_slides.append({
                "id": f"s{idx}",
                "index": idx,
                "part": member,
                "shape_count": len(shapes_records),
                "shapes": shapes_records,
            })

    return {
        "deck": os.path.basename(deck_path),
        "deck_path": os.path.abspath(deck_path),
        "fingerprint": fingerprint(deck_path),
        "slide_size_in": [round(sw_emu / 914400, 3),
                          round(sh_emu / 914400, 3)] if sw_emu else None,
        "slide_count": len(result_slides),
        "slides": result_slides,
    }


def parse_slide_arg(arg: str | None) -> set[int] | None:
    if not arg:
        return None
    out: set[int] = set()
    for piece in arg.split(","):
        piece = piece.strip()
        if not piece:
            continue
        if "-" in piece:
            a, b = piece.split("-", 1)
            out.update(range(int(a), int(b) + 1))
        else:
            out.add(int(piece))
    return out


def cmd_check_stale(audit_path: str) -> int:
    """Compare a previous audit's fingerprint against the live deck on disk.

    Exit codes:
      0 — deck unchanged since audit (safe to act on)
      1 — deck has changed since audit (re-audit before continuing)
      2 — error (audit not found, malformed, missing fingerprint, etc.)
    """
    if not os.path.exists(audit_path):
        print(f"[FAIL] audit not found: {audit_path}", file=sys.stderr)
        return 2
    try:
        with open(audit_path, "r", encoding="utf-8") as f:
            audit_data = json.load(f)
    except json.JSONDecodeError as e:
        print(f"[FAIL] audit not valid JSON: {e}", file=sys.stderr)
        return 2

    deck_path = audit_data.get("deck_path")
    prior = audit_data.get("fingerprint")
    if not deck_path or not prior:
        print("[FAIL] audit is missing 'deck_path' or 'fingerprint'; "
              "regenerate it with the current xml_audit.py", file=sys.stderr)
        return 2

    stale, reason = is_stale(deck_path, prior)
    if stale:
        print(f"[STALE] {deck_path}\n        {reason}\n"
              f"        re-run xml_audit.py against the current deck "
              f"before drafting any patch.", file=sys.stderr)
        return 1

    print(f"[FRESH] {deck_path}\n        {reason}")
    return 0


def main(argv: Iterable[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Audit a .pptx by reading slide XML; emit run-level JSON.")
    ap.add_argument("deck", nargs="?",
                    help="path to .pptx (omit when using --check-stale)")
    ap.add_argument("--out",
                    help="output JSON path (default: <deck>.audit.json)")
    ap.add_argument("--slides",
                    help="comma-separated slide numbers / ranges, e.g. '1,3-5'")
    ap.add_argument("--pretty", action="store_true",
                    help="pretty-print JSON with 2-space indent")
    ap.add_argument("--check-stale", metavar="AUDIT_JSON",
                    help=("compare the deck on disk to a prior audit; print "
                          "FRESH (exit 0) or STALE (exit 1) and quit. Does "
                          "not write any files."))
    args = ap.parse_args(list(argv) if argv is not None else None)

    if args.check_stale:
        return cmd_check_stale(args.check_stale)

    if not args.deck:
        ap.error("deck path is required unless --check-stale is given")
    if not os.path.exists(args.deck):
        print(f"[FAIL] not found: {args.deck}", file=sys.stderr)
        return 1
    if not zipfile.is_zipfile(args.deck):
        print(f"[FAIL] not a valid .pptx (zip): {args.deck}", file=sys.stderr)
        return 1

    data = audit(args.deck, parse_slide_arg(args.slides))

    out_path = args.out or os.path.splitext(args.deck)[0] + ".audit.json"
    with open(out_path, "w", encoding="utf-8") as f:
        if args.pretty:
            json.dump(data, f, ensure_ascii=False, indent=2)
        else:
            json.dump(data, f, ensure_ascii=False)
    n_runs = sum(
        len(p["runs"])
        for s in data["slides"]
        for sh in s["shapes"]
        for p in sh["paragraphs"]
    )
    print(f"[ OK ] wrote {out_path} "
          f"({data['slide_count']} slide(s), {n_runs} run(s))")
    return 0


if __name__ == "__main__":
    sys.exit(main())
