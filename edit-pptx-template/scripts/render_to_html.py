"""
Render a .pptx as a static HTML preview for visual review without PowerPoint.

This uses LibreOffice's own rendering engine (the same code that draws Impress
slides on screen), which gives results visually equivalent to opening the deck
in PowerPoint — fonts, colors, line wrapping, shape geometry, all faithful.

Pipeline:
    .pptx --[LibreOffice headless]--> .pdf --[PyMuPDF]--> N PNGs --> HTML index

The output is a single self-contained HTML file with the slide images embedded
as base64. Open it in a browser, or hand it to someone who doesn't have
PowerPoint, or let an agent read it back to confirm the visual layout.

Usage:
    python render_to_html.py deck.pptx                   # writes deck_preview.html
    python render_to_html.py deck.pptx --out out.html
    python render_to_html.py deck.pptx --dpi 144         # crisper render (default 144)
    python render_to_html.py deck.pptx --keep-pdf        # also keep the intermediate PDF

Requirements:
    * LibreOffice (`soffice`) installed and on PATH, or in a standard
      location like /Applications/LibreOffice.app on macOS.
    * `pymupdf` (`pip install pymupdf`) for PDF -> PNG.
"""

from __future__ import annotations

import argparse
import base64
import os
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass


SOFFICE_CANDIDATES = [
    "soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
    "/snap/bin/libreoffice",
]


def find_soffice() -> str | None:
    """Locate the LibreOffice binary, falling back to common install paths."""
    for cand in SOFFICE_CANDIDATES:
        if os.path.isabs(cand):
            if os.path.exists(cand):
                return cand
        else:
            found = shutil.which(cand)
            if found:
                return found
    return None


@dataclass
class SlideImage:
    index: int
    width: int
    height: int
    png_bytes: bytes


def convert_pptx_to_pdf(pptx_path: str, soffice: str, work_dir: str) -> str:
    """Run LibreOffice headless to produce a PDF next to the input."""
    cmd = [
        soffice, "--headless", "--norestore", "--nologo", "--nofirststartwizard",
        "--convert-to", "pdf", "--outdir", work_dir, pptx_path,
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    if proc.returncode != 0:
        raise RuntimeError(
            f"LibreOffice failed (exit {proc.returncode}):\n"
            f"stdout: {proc.stdout}\nstderr: {proc.stderr}"
        )
    base = os.path.splitext(os.path.basename(pptx_path))[0]
    pdf_path = os.path.join(work_dir, base + ".pdf")
    if not os.path.exists(pdf_path):
        raise RuntimeError(
            f"LibreOffice did not produce a PDF at {pdf_path}.\n"
            f"stdout: {proc.stdout}\nstderr: {proc.stderr}"
        )
    return pdf_path


def render_pdf_to_pngs(pdf_path: str, dpi: int) -> list[SlideImage]:
    """Render every page of a PDF to a PNG buffer using PyMuPDF."""
    try:
        import fitz  # PyMuPDF
    except ImportError as e:
        raise RuntimeError(
            "PyMuPDF is required: pip install pymupdf"
        ) from e

    # PyMuPDF prints non-fatal warnings (e.g. "No common ancestor in structure
    # tree") for PDFs LibreOffice produces. Silence those so the script's
    # output stays clean — they don't affect rendering.
    fitz.TOOLS.mupdf_display_errors(False)

    # PyMuPDF default resolution is 72 DPI; matrix scales it.
    scale = dpi / 72
    matrix = fitz.Matrix(scale, scale)

    images: list[SlideImage] = []
    with fitz.open(pdf_path) as doc:
        for i, page in enumerate(doc, start=1):
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            images.append(SlideImage(
                index=i,
                width=pix.width,
                height=pix.height,
                png_bytes=pix.tobytes("png"),
            ))
    return images


HTML_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>{title} — preview</title>
<style>
  * {{ box-sizing: border-box; }}
  body {{
    font-family: -apple-system, "Helvetica Neue", Arial, sans-serif;
    background: #f5f5f7;
    margin: 0;
    padding: 24px;
    color: #1B2A4A;
  }}
  h1 {{ margin: 0 0 4px; font-size: 18px; font-weight: 600; }}
  .meta {{ color: #6b7280; font-size: 12px; margin-bottom: 24px; }}
  .toc {{
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
    margin-bottom: 20px;
    font-size: 12px;
  }}
  .toc a {{
    background: white;
    border: 1px solid #d0d5dd;
    border-radius: 4px;
    padding: 3px 8px;
    color: #1B2A4A;
    text-decoration: none;
  }}
  .toc a:hover {{ background: #e8f1fa; }}
  .slide {{
    background: white;
    border: 1px solid #d0d5dd;
    border-radius: 8px;
    margin: 0 0 24px;
    overflow: hidden;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    scroll-margin-top: 16px;
  }}
  .slide-header {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 8px 14px;
    font-size: 12px;
    color: #4a4a4a;
    border-bottom: 1px solid #eef0f3;
    background: #fafbfc;
  }}
  .slide-header .label {{ font-weight: 600; }}
  .slide-img {{
    display: block;
    width: 100%;
    height: auto;
    background: white;
  }}
</style>
</head>
<body>
<h1>{title}</h1>
<div class="meta">
  {n_slides} slide(s) &middot; rendered at {dpi} DPI by LibreOffice &middot;
  open in a browser to inspect visual layout
</div>
<div class="toc">{toc}</div>
{slides}
</body>
</html>
"""


def build_html(deck_name: str, images: list[SlideImage], dpi: int) -> str:
    toc = "".join(
        f'<a href="#slide-{img.index}">{img.index}</a>' for img in images
    )
    slide_blocks = []
    for img in images:
        b64 = base64.b64encode(img.png_bytes).decode("ascii")
        slide_blocks.append(
            f'<section class="slide" id="slide-{img.index}">'
            f'  <div class="slide-header">'
            f'    <span class="label">Slide {img.index}</span>'
            f'    <span>{img.width} × {img.height} px</span>'
            f'  </div>'
            f'  <img class="slide-img" alt="Slide {img.index}" '
            f'       src="data:image/png;base64,{b64}">'
            f'</section>'
        )
    return HTML_TEMPLATE.format(
        title=deck_name,
        n_slides=len(images),
        dpi=dpi,
        toc=toc,
        slides="\n".join(slide_blocks),
    )


def render(pptx_path: str, out_path: str, dpi: int = 144,
           keep_pdf: bool = False) -> tuple[str, list[SlideImage]]:
    """Render `pptx_path` to a self-contained HTML at `out_path`."""
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "Could not find LibreOffice (`soffice`). Install it via:\n"
            "  brew install --cask libreoffice    # macOS\n"
            "  apt-get install libreoffice        # Debian/Ubuntu\n"
            "Or pass --soffice to point at the binary."
        )

    with tempfile.TemporaryDirectory(prefix="render_to_html_") as tmp:
        pdf_path = convert_pptx_to_pdf(pptx_path, soffice, tmp)
        images = render_pdf_to_pngs(pdf_path, dpi=dpi)
        if keep_pdf:
            kept = os.path.splitext(out_path)[0] + ".pdf"
            shutil.copy(pdf_path, kept)

    deck_name = os.path.basename(pptx_path)
    html_text = build_html(deck_name, images, dpi=dpi)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html_text)
    return out_path, images


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("deck", help="path to .pptx")
    ap.add_argument("--out", help="output HTML path "
                                  "(default: <deck>_preview.html next to the input)")
    ap.add_argument("--dpi", type=int, default=144,
                    help="render resolution (default 144; use 192 for retina)")
    ap.add_argument("--keep-pdf", action="store_true",
                    help="also keep the intermediate PDF next to the HTML")
    ap.add_argument("--soffice", help="explicit path to the LibreOffice binary")
    args = ap.parse_args(argv)

    if not os.path.exists(args.deck):
        print(f"[FAIL] not found: {args.deck}", file=sys.stderr)
        return 1
    if args.soffice:
        SOFFICE_CANDIDATES.insert(0, args.soffice)

    out_path = args.out or os.path.splitext(args.deck)[0] + "_preview.html"
    try:
        out, images = render(args.deck, out_path,
                             dpi=args.dpi, keep_pdf=args.keep_pdf)
    except RuntimeError as e:
        print(f"[FAIL] {e}", file=sys.stderr)
        return 1

    size_kb = os.path.getsize(out) / 1024
    print(f"[ OK ] wrote {out} ({len(images)} slide(s), {size_kb:.0f} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
