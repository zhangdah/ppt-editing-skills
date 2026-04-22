"""
Verify layout bounds of a generated .pptx file.

Usage:
    python verify_layout.py path/to/deck.pptx

Checks, per slide:
  1. No shape extends outside the slide canvas (left/top >= 0 and
     left+width <= slide_width, top+height <= slide_height).
  2. Warn on large text boxes whose configured height is clearly too small
     for their content (heuristic: avg chars per line * #paragraphs).
  3. Warn on shapes that visually overlap significantly (bounding-box IoU
     above a threshold) — often an unintended stacking error.

Prints a report. Exits non-zero only on hard bound violations; overlap and
overflow heuristics are warnings (they sometimes occur intentionally).
"""

import sys
import os

from pptx import Presentation
from pptx.util import Emu


EMU_PER_INCH = 914400
OVERLAP_IOU_THRESHOLD = 0.6


def iou(a, b):
    ax1, ay1, ax2, ay2 = a
    bx1, by1, bx2, by2 = b
    ix1, iy1 = max(ax1, bx1), max(ay1, by1)
    ix2, iy2 = min(ax2, bx2), min(ay2, by2)
    if ix2 <= ix1 or iy2 <= iy1:
        return 0.0
    inter = (ix2 - ix1) * (iy2 - iy1)
    area_a = (ax2 - ax1) * (ay2 - ay1)
    area_b = (bx2 - bx1) * (by2 - by1)
    union = area_a + area_b - inter
    return inter / union if union > 0 else 0.0


def emu_to_in(v):
    return v / EMU_PER_INCH


def shape_text(shape):
    if not shape.has_text_frame:
        return ""
    return "\n".join(p.text for p in shape.text_frame.paragraphs)


def main():
    if len(sys.argv) < 2:
        print("usage: verify_layout.py <deck.pptx>")
        sys.exit(2)
    path = sys.argv[1]
    if not os.path.exists(path):
        print(f"[FAIL] not found: {path}")
        sys.exit(1)

    prs = Presentation(path)
    sw, sh = prs.slide_width, prs.slide_height
    print(f"slide canvas: {emu_to_in(sw):.2f} x {emu_to_in(sh):.2f} inches\n")

    hard_fail = False
    for idx, slide in enumerate(prs.slides, start=1):
        print(f"--- slide {idx} ({len(slide.shapes)} shapes) ---")
        boxes = []
        for s in slide.shapes:
            if s.left is None or s.top is None:
                continue
            l, t = s.left, s.top
            w = s.width or 0
            h = s.height or 0
            right, bottom = l + w, t + h
            name = s.name
            if l < 0 or t < 0 or right > sw or bottom > sh:
                print(
                    f"  [OUT OF BOUNDS] {name}: "
                    f"({emu_to_in(l):.2f},{emu_to_in(t):.2f}) "
                    f"-> ({emu_to_in(right):.2f},{emu_to_in(bottom):.2f})"
                )
                hard_fail = True
            boxes.append(((l, t, right, bottom), name, s))

        for (box, name, s) in boxes:
            if not s.has_text_frame:
                continue
            txt = shape_text(s)
            n_paras = len([p for p in s.text_frame.paragraphs if p.text])
            chars = len(txt)
            w_in = emu_to_in(s.width or 0)
            h_in = emu_to_in(s.height or 0)
            if n_paras >= 1 and chars > 0 and w_in > 0 and h_in > 0:
                approx_lines = max(n_paras, int(chars / max(1, w_in * 10)))
                needed_in = approx_lines * 0.22
                if needed_in > h_in * 1.25:
                    print(
                        f"  [WARN overflow?] {name}: "
                        f"needs ~{needed_in:.2f}in, has {h_in:.2f}in "
                        f"({chars} chars, {n_paras} paras)"
                    )

        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                score = iou(boxes[i][0], boxes[j][0])
                if score > OVERLAP_IOU_THRESHOLD:
                    a = boxes[i][2]
                    b = boxes[j][2]
                    if a.has_text_frame and b.has_text_frame and (
                        shape_text(a).strip() or shape_text(b).strip()
                    ):
                        print(
                            f"  [WARN overlap iou={score:.2f}] "
                            f"{boxes[i][1]} vs {boxes[j][1]}"
                        )

        print()

    if hard_fail:
        print("[FAIL] one or more shapes are outside the slide canvas")
        sys.exit(1)
    print("[PASS] no hard bound violations (review warnings above)")


if __name__ == "__main__":
    main()
