"""
Apply a patch.json to a .pptx, surgically editing slide XML in place.

Only THREE operations are supported, by design:

    op = "set_text"         # replace the text inside a single run (<a:t>)
    op = "set_font"         # set/clear font attributes on a single run (<a:rPr>)
    op = "set_default_font" # set/clear default font on a paragraph (<a:pPr>/<a:defRPr>)

Everything else (move/resize shapes, add/remove shapes, change fills, change
slide layout, theme, master, etc.) is REJECTED — this script's contract is
"text and font edits only, byte-clean elsewhere".

Usage:
    python3 xml_patch.py deck.pptx patch.json --out edited.pptx
    python3 xml_patch.py deck.pptx patch.json --in-place         # overwrite

Patch format (JSON array; comments shown only as //):
    [
      {
        "op": "set_text",
        "target": "s1.sp4.p0.r0",
        "value": "Updated headline"
      },
      {
        "op": "set_font",
        "target": "s1.sp5.p0.r0",
        "value": {
          "size_pt": 36,
          "bold": true,
          "italic": false,
          "color": "FFFFFF",
          "font_name": "Inter",
          "underline": "sng",
          "strike": "noStrike"
        }
      },
      {
        "op": "set_default_font",
        "target": "s1.sp7.p1",
        "value": {"size_pt": 14, "color": "8C8C8C", "font_name": "Inter"}
      }
    ]

Setting any font field to JSON `null` REMOVES that attribute (so the run/para
inherits from its parent). Omitted fields are left untouched.

Pure standard library — no python-pptx, no lxml.
"""

from __future__ import annotations

import argparse
import io
import json
import os
import re
import shutil
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pptx_lock import assert_writable
from _pptx_fingerprint import is_stale
from typing import Iterable
from xml.etree import ElementTree as ET


NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def _q(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


SLIDE_RE = re.compile(r"^ppt/slides/slide(\d+)\.xml$")
ID_RE = re.compile(
    r"^s(?P<slide>\d+)"
    r"(?:\.sp(?P<sp>\d+))?"
    r"(?:\.p(?P<p>\d+))?"
    r"(?:\.r(?P<r>\d+))?$"
)


# -----------------------------------------------------------------------------
# Errors
# -----------------------------------------------------------------------------
class PatchError(Exception):
    """Raised for invalid patches; message is shown to the user."""


# -----------------------------------------------------------------------------
# ID parsing & navigation
# -----------------------------------------------------------------------------
def parse_id(target: str) -> dict:
    m = ID_RE.match(target)
    if not m:
        raise PatchError(f"invalid target id: {target!r}")
    return {
        "slide": int(m.group("slide")),
        "sp": int(m.group("sp")) if m.group("sp") is not None else None,
        "p": int(m.group("p")) if m.group("p") is not None else None,
        "r": int(m.group("r")) if m.group("r") is not None else None,
    }


def find_shape(slide_root: ET.Element, sp_id: int) -> ET.Element:
    """Return the <p:sp> whose <p:cNvPr id="sp_id">. Walks groups too."""
    target = str(sp_id)

    def visit(elem: ET.Element) -> ET.Element | None:
        for child in elem:
            local = child.tag.split("}", 1)[-1]
            if local == "sp":
                cNvPr = child.find(f".//{_q('p','cNvPr')}")
                if cNvPr is not None and cNvPr.attrib.get("id") == target:
                    return child
            elif local == "grpSp":
                hit = visit(child)
                if hit is not None:
                    return hit
        return None

    spTree = slide_root.find(f"{_q('p','cSld')}/{_q('p','spTree')}")
    if spTree is None:
        raise PatchError("slide has no spTree")
    sp = visit(spTree)
    if sp is None:
        raise PatchError(f"shape id sp{sp_id} not found")
    return sp


def find_paragraph(sp: ET.Element, p_idx: int) -> ET.Element:
    txBody = sp.find(_q("p", "txBody"))
    if txBody is None:
        raise PatchError("shape has no <p:txBody>")
    paragraphs = txBody.findall(_q("a", "p"))
    if not (0 <= p_idx < len(paragraphs)):
        raise PatchError(
            f"paragraph index {p_idx} out of range "
            f"(shape has {len(paragraphs)} paragraph(s))")
    return paragraphs[p_idx]


def find_run(p: ET.Element, r_idx: int) -> tuple[ET.Element, str]:
    """Return (element, kind) where kind is 'r' or 'br'. Counts both <a:r>
    and <a:br> the same way audit does, so IDs line up."""
    counter = -1
    for child in p:
        local = child.tag.split("}", 1)[-1]
        if local in ("r", "br"):
            counter += 1
            if counter == r_idx:
                return child, local
    raise PatchError(f"run index {r_idx} out of range in paragraph")


# -----------------------------------------------------------------------------
# Element editors
# -----------------------------------------------------------------------------
HEX_RE = re.compile(r"^[0-9A-Fa-f]{6}$")


def _set_or_remove_attr(elem: ET.Element, name: str, value) -> None:
    """If value is None, drop the attribute. Otherwise set it to str(value)."""
    if value is None:
        if name in elem.attrib:
            del elem.attrib[name]
    else:
        elem.attrib[name] = str(value)


# OOXML expects this exact child order inside <a:rPr>/<a:defRPr> after attrs.
# We don't enforce it strictly when modifying; PowerPoint is lenient and our
# audit re-reads from XML, but we keep additions in this order to match what
# python-pptx and stock decks emit.
_RPR_CHILD_ORDER = [
    "ln", "fillOverlay", "highlight", "uLnTx", "uLn", "uFillTx", "uFill",
    "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl",
    "extLst", "solidFill", "noFill", "gradFill", "blipFill", "pattFill",
    "grpFill",
]


def _ensure_or_drop_typeface(rpr: ET.Element, child_local: str,
                             typeface: str | None) -> None:
    """child_local in {'latin','ea','cs'}. None drops the child."""
    qname = _q("a", child_local)
    existing = rpr.find(qname)
    if typeface is None:
        if existing is not None:
            rpr.remove(existing)
        return
    if existing is None:
        existing = ET.SubElement(rpr, qname)
    existing.attrib["typeface"] = typeface


def _ensure_or_drop_solid_fill(rpr: ET.Element, color: str | None) -> None:
    """Manage the <a:solidFill><a:srgbClr val="..."/></a:solidFill> child.
    color=None drops solidFill entirely; missing fill -> inherit."""
    fill = rpr.find(_q("a", "solidFill"))
    if color is None:
        if fill is not None:
            rpr.remove(fill)
        return
    if not HEX_RE.match(color):
        raise PatchError(
            f"color must be 6-hex like 'FFFFFF', got {color!r}")
    if fill is None:
        fill = ET.SubElement(rpr, _q("a", "solidFill"))
    fill.clear()
    srgb = ET.SubElement(fill, _q("a", "srgbClr"))
    srgb.attrib["val"] = color.upper()


def apply_font(rpr: ET.Element, font: dict) -> None:
    """Mutate <a:rPr> (or <a:defRPr>) in place to reflect the given font dict.
    Only keys present in the dict are touched; value=None means 'remove'."""
    for key, value in font.items():
        if key == "size_pt":
            if value is None:
                rpr.attrib.pop("sz", None)
            else:
                if not isinstance(value, (int, float)) or value <= 0:
                    raise PatchError(
                        f"size_pt must be a positive number, got {value!r}")
                rpr.attrib["sz"] = str(int(round(float(value) * 100)))
        elif key == "bold":
            _set_or_remove_attr(
                rpr, "b", None if value is None else ("1" if value else "0"))
        elif key == "italic":
            _set_or_remove_attr(
                rpr, "i", None if value is None else ("1" if value else "0"))
        elif key == "underline":
            # Pass through enum strings: 'sng' / 'dbl' / 'none' / etc.
            if value is None or value is False:
                rpr.attrib.pop("u", None)
            elif value is True:
                rpr.attrib["u"] = "sng"
            else:
                rpr.attrib["u"] = str(value)
        elif key == "strike":
            if value is None or value is False:
                rpr.attrib.pop("strike", None)
            elif value is True:
                rpr.attrib["strike"] = "sngStrike"
            else:
                rpr.attrib["strike"] = str(value)
        elif key == "color":
            _ensure_or_drop_solid_fill(rpr, value)
        elif key == "font_name":
            _ensure_or_drop_typeface(rpr, "latin", value)
        elif key == "font_name_ea":
            _ensure_or_drop_typeface(rpr, "ea", value)
        elif key == "font_name_cs":
            _ensure_or_drop_typeface(rpr, "cs", value)
        else:
            raise PatchError(f"unknown font field: {key!r}")


# -----------------------------------------------------------------------------
# Op implementations
# -----------------------------------------------------------------------------
def op_set_text(slide_root: ET.Element, parts: dict, value) -> None:
    if parts["p"] is None or parts["r"] is None:
        raise PatchError("set_text target must point to a run, e.g. s1.sp4.p0.r0")
    if not isinstance(value, str):
        raise PatchError(f"set_text value must be a string, got {type(value).__name__}")

    sp = find_shape(slide_root, parts["sp"])
    p = find_paragraph(sp, parts["p"])
    run, kind = find_run(p, parts["r"])
    if kind == "br":
        raise PatchError("cannot set_text on a <a:br> (line break) run")

    t = run.find(_q("a", "t"))
    if t is None:
        t = ET.SubElement(run, _q("a", "t"))
    t.text = value


def op_set_font(slide_root: ET.Element, parts: dict, value) -> None:
    if parts["p"] is None or parts["r"] is None:
        raise PatchError("set_font target must point to a run, e.g. s1.sp4.p0.r0")
    if not isinstance(value, dict):
        raise PatchError("set_font value must be an object of font fields")

    sp = find_shape(slide_root, parts["sp"])
    p = find_paragraph(sp, parts["p"])
    run, kind = find_run(p, parts["r"])

    rpr = run.find(_q("a", "rPr"))
    if rpr is None:
        rpr = ET.Element(_q("a", "rPr"))
        run.insert(0, rpr)
    apply_font(rpr, value)


def op_set_default_font(slide_root: ET.Element, parts: dict, value) -> None:
    if parts["p"] is None or parts["r"] is not None:
        raise PatchError(
            "set_default_font target must point to a paragraph, "
            "e.g. s1.sp4.p0 (no .rN suffix)")
    if not isinstance(value, dict):
        raise PatchError("set_default_font value must be an object of font fields")

    sp = find_shape(slide_root, parts["sp"])
    p = find_paragraph(sp, parts["p"])

    pPr = p.find(_q("a", "pPr"))
    if pPr is None:
        pPr = ET.Element(_q("a", "pPr"))
        p.insert(0, pPr)

    defRPr = pPr.find(_q("a", "defRPr"))
    if defRPr is None:
        defRPr = ET.SubElement(pPr, _q("a", "defRPr"))
    apply_font(defRPr, value)


OPS = {
    "set_text": op_set_text,
    "set_font": op_set_font,
    "set_default_font": op_set_default_font,
}


# -----------------------------------------------------------------------------
# Patch driver
# -----------------------------------------------------------------------------
def parse_patch_doc(doc) -> tuple[list, str | None]:
    """Accept both bare-array and {"based_on": ..., "ops": [...]} forms.

    Returns (ops, based_on_audit_path_or_None).
    """
    if isinstance(doc, list):
        return doc, None
    if isinstance(doc, dict):
        if "ops" not in doc:
            raise PatchError(
                "patch object must have 'ops' (array) at minimum")
        if not isinstance(doc["ops"], list):
            raise PatchError("patch 'ops' must be a JSON array")
        based_on = doc.get("based_on")
        if based_on is not None and not isinstance(based_on, str):
            raise PatchError("patch 'based_on' must be a string path or absent")
        return doc["ops"], based_on
    raise PatchError(
        "patch must be a JSON array of ops, or an object with 'ops' (and "
        "optionally 'based_on' pointing to the audit it was drafted against)")


def assert_audit_fresh(audit_path: str, *, force: bool = False) -> None:
    """If the patch references an audit, verify the deck still matches it."""
    if force:
        return
    if not os.path.isabs(audit_path):
        # Relative paths in patch.json are resolved against the patch's own
        # directory at load time; the caller has already done that.
        pass
    if not os.path.exists(audit_path):
        raise PatchError(
            f"patch's 'based_on' audit not found: {audit_path}\n"
            f"       Re-run xml_audit.py and update 'based_on', or pass "
            f"--force.")
    try:
        with open(audit_path, "r", encoding="utf-8") as f:
            audit_data = json.load(f)
    except json.JSONDecodeError as e:
        raise PatchError(f"audit '{audit_path}' is not valid JSON: {e}")

    deck_path = audit_data.get("deck_path")
    prior = audit_data.get("fingerprint")
    if not deck_path or not prior:
        raise PatchError(
            f"audit '{audit_path}' has no fingerprint; regenerate it with "
            f"the current xml_audit.py, or pass --force.")

    stale, reason = is_stale(deck_path, prior)
    if stale:
        raise PatchError(
            f"deck has changed since the audit was taken:\n"
            f"           audit: {audit_path}\n"
            f"           deck:  {deck_path}\n"
            f"           why:   {reason}\n"
            f"       Re-run xml_audit.py against the current deck, review "
            f"the new state with the user, then redraft the patch. Pass "
            f"--force to bypass.")


def validate_patch(patch: list) -> None:
    if not isinstance(patch, list):
        raise PatchError("patch must be a JSON array of operation objects")
    for i, item in enumerate(patch):
        if not isinstance(item, dict):
            raise PatchError(f"patch[{i}] must be an object")
        for required in ("op", "target", "value"):
            if required not in item:
                raise PatchError(f"patch[{i}] missing required field: {required!r}")
        if item["op"] not in OPS:
            raise PatchError(
                f"patch[{i}].op = {item['op']!r} is not allowed. "
                f"Allowed ops: {sorted(OPS)}. "
                f"This tool only edits text and font properties; layout, "
                f"shape, and theme changes are out of scope.")


def group_by_slide(patch: list) -> dict[int, list[dict]]:
    out: dict[int, list[dict]] = {}
    for item in patch:
        parts = parse_id(item["target"])
        out.setdefault(parts["slide"], []).append({**item, "_parts": parts})
    return out


_ROOT_OPEN_RE = re.compile(rb"<\s*p:sld\b[^>]*>", re.DOTALL)
_XMLNS_RE = re.compile(rb'xmlns:([A-Za-z_][\w.\-]*)\s*=\s*"([^"]+)"')


def _extract_original_root_tag(raw: bytes) -> bytes | None:
    """Return the literal opening <p:sld ...> tag from the source XML, so we
    can restore its full namespace declaration block after re-serialization.
    Returns None if the root is not <p:sld> (e.g. a different slide-like part)."""
    m = _ROOT_OPEN_RE.search(raw)
    return m.group(0) if m else None


def _scan_xmlns_decls(raw: bytes) -> list[tuple[str, str]]:
    """Return [(prefix, uri), ...] for every xmlns:* declaration anywhere in
    the document. OOXML files can declare extension namespaces inline (e.g.
    p14, p15, mc, p:ext blocks), and ElementTree will assign auto-generated
    'ns0', 'ns1', ... prefixes for any URI it doesn't already know about,
    losing the original prefix and breaking files."""
    out: list[tuple[str, str]] = []
    for m in _XMLNS_RE.finditer(raw):
        out.append((m.group(1).decode("ascii"), m.group(2).decode("ascii")))
    return out


def apply_patch_to_xml(raw_xml: bytes, ops: list[dict]) -> bytes:
    """Apply all ops belonging to a single slide. Returns the new XML bytes."""
    original_root_tag = _extract_original_root_tag(raw_xml)

    # ElementTree resolves namespace prefixes via a process-global registry
    # (`register_namespace`). To preserve every original prefix in the output,
    # walk the source for xmlns:* declarations and register each before
    # serialization. We restore the previous registry afterwards so this
    # function has no global side effects.
    extra_decls = _scan_xmlns_decls(raw_xml)
    saved_registry = dict(getattr(ET, "_namespace_map", {}))
    try:
        for prefix, uri in extra_decls:
            # Skip the empty/default-namespace declaration to avoid
            # accidentally altering the default ns mapping.
            if not prefix:
                continue
            ET.register_namespace(prefix, uri)

        root = ET.fromstring(raw_xml)
        for item in ops:
            op_fn = OPS[item["op"]]
            if item["_parts"]["sp"] is None:
                raise PatchError(
                    f"target {item['target']!r} must include a shape (.spN)")
            op_fn(root, item["_parts"], item["value"])

        buf = io.BytesIO()
        ET.ElementTree(root).write(buf, xml_declaration=True,
                                   encoding="UTF-8",
                                   short_empty_elements=True)
        out = buf.getvalue()
    finally:
        if hasattr(ET, "_namespace_map"):
            ET._namespace_map.clear()
            ET._namespace_map.update(saved_registry)

    # ET writes every xmlns it actually uses on the root element. The set of
    # declarations may differ from the original (extension namespaces declared
    # inline in the source get hoisted to the root). Restore the original
    # opening <p:sld ...> tag verbatim — and union any ET-emitted xmlns:*
    # declarations into it so that prefixes registered above are still bound.
    if original_root_tag is not None:
        m = _ROOT_OPEN_RE.search(out)
        if m is not None:
            new_root_tag = m.group(0)
            original_prefixes = {
                p for p, _ in _scan_xmlns_decls(original_root_tag)
            }
            extras = [
                f' xmlns:{p}="{u}"'.encode("ascii")
                for p, u in _scan_xmlns_decls(new_root_tag)
                if p not in original_prefixes
            ]
            if extras:
                # Insert extras just before the closing '>' of the original tag.
                merged = original_root_tag[:-1] + b"".join(extras) \
                    + original_root_tag[-1:]
            else:
                merged = original_root_tag
            out = out[:m.start()] + merged + out[m.end():]
    return out


def write_pptx(src_path: str, out_path: str,
               replacements: dict[str, bytes]) -> None:
    """Copy src to out, replacing the listed members with new bytes.
    All other members are copied byte-for-byte."""
    if os.path.abspath(src_path) == os.path.abspath(out_path):
        raise PatchError(
            "src and out are the same path; use --in-place or pick a different --out")
    with zipfile.ZipFile(src_path) as zin, \
         zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = replacements.get(info.filename, zin.read(info.filename))
            new_info = zipfile.ZipInfo(filename=info.filename,
                                       date_time=info.date_time)
            new_info.compress_type = info.compress_type
            new_info.external_attr = info.external_attr
            new_info.create_system = info.create_system
            zout.writestr(new_info, data)


def apply(deck_path: str, patch_path: str, out_path: str,
          force_stale: bool = False) -> dict:
    with open(patch_path, "r", encoding="utf-8") as f:
        doc = json.load(f)
    ops, based_on = parse_patch_doc(doc)

    if based_on is not None:
        if not os.path.isabs(based_on):
            based_on = os.path.normpath(
                os.path.join(os.path.dirname(os.path.abspath(patch_path)),
                             based_on))
        assert_audit_fresh(based_on, force=force_stale)

    validate_patch(ops)
    ops_total = len(ops)
    by_slide = group_by_slide(ops)

    replacements: dict[str, bytes] = {}
    summary_per_slide: list[dict] = []
    with zipfile.ZipFile(deck_path) as zf:
        slide_members = {}
        for name in zf.namelist():
            m = SLIDE_RE.match(name)
            if m:
                slide_members[int(m.group(1))] = name

        for slide_idx, slide_ops in by_slide.items():
            if slide_idx not in slide_members:
                raise PatchError(
                    f"patch refers to slide {slide_idx}, but deck only has "
                    f"slides {sorted(slide_members)}")
            member = slide_members[slide_idx]
            new_xml = apply_patch_to_xml(zf.read(member), slide_ops)
            replacements[member] = new_xml
            summary_per_slide.append({
                "slide": slide_idx,
                "part": member,
                "ops_applied": len(slide_ops),
            })

    write_pptx(deck_path, out_path, replacements)
    return {
        "deck_in": os.path.abspath(deck_path),
        "deck_out": os.path.abspath(out_path),
        "ops_total": ops_total,
        "slides_modified": summary_per_slide,
        "based_on": based_on,
    }


def main(argv: Iterable[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Apply a JSON patch to a .pptx (text + font edits only).")
    ap.add_argument("deck", help="path to source .pptx")
    ap.add_argument("patch", help="path to patch.json")
    group = ap.add_mutually_exclusive_group(required=True)
    group.add_argument("--out", help="path for the edited .pptx")
    group.add_argument("--in-place", action="store_true",
                       help="overwrite the source deck (writes via temp file)")
    ap.add_argument("--force", action="store_true",
                    help=("bypass safety checks: the lock-holder probe AND "
                          "the audit-staleness probe. Use only when you are "
                          "certain no other app holds the deck open and the "
                          "deck has not been edited since the audit."))
    args = ap.parse_args(list(argv) if argv is not None else None)

    if not os.path.exists(args.deck):
        print(f"[FAIL] deck not found: {args.deck}", file=sys.stderr)
        return 1
    if not zipfile.is_zipfile(args.deck):
        print(f"[FAIL] deck is not a valid .pptx (zip): {args.deck}", file=sys.stderr)
        return 1
    if not os.path.exists(args.patch):
        print(f"[FAIL] patch not found: {args.patch}", file=sys.stderr)
        return 1

    if args.in_place:
        out_path = args.deck + ".tmp"
        assert_writable(args.deck, force=args.force, label="source deck")
    else:
        out_path = args.out
        assert_writable(out_path, force=args.force, label="output deck")

    try:
        summary = apply(args.deck, args.patch, out_path,
                        force_stale=args.force)
    except PatchError as e:
        print(f"[FAIL] patch error: {e}", file=sys.stderr)
        if args.in_place and os.path.exists(out_path):
            os.remove(out_path)
        return 2
    except Exception as e:
        print(f"[FAIL] unexpected error: {type(e).__name__}: {e}", file=sys.stderr)
        if args.in_place and os.path.exists(out_path):
            os.remove(out_path)
        return 3

    if args.in_place:
        shutil.move(out_path, args.deck)
        summary["deck_out"] = os.path.abspath(args.deck)

    print(f"[ OK ] applied {summary['ops_total']} op(s) across "
          f"{len(summary['slides_modified'])} slide(s) -> {summary['deck_out']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
