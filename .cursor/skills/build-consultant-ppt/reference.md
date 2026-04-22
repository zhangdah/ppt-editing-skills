# Reference: Reusable Helpers & Verification Scripts

Companion to `SKILL.md`. This document describes the concrete tools shipped with this skill and how to wire them into a generator script.

Everything below is **general-purpose** — nothing about it assumes a specific topic, company, or deck. Pick what you need.

---

## Directory layout

```
build-consultant-ppt/
├── SKILL.md                  # workflow (read first)
├── reference.md              # this file
└── scripts/
    ├── pptx_helpers.py       # import into your generator
    ├── verify_structure.py   # check file validity + XML
    ├── verify_layout.py      # check shapes stay on the canvas
    └── inspect_slide.py      # dump one slide to stdout for debugging
```

The `scripts/` folder is self-contained. Copy `pptx_helpers.py` next to the generator you write (or add the folder to `sys.path`).

---

## 1. `pptx_helpers.py` — helper library

Import what you need:

```python
from pptx_helpers import (
    Palette, Fonts,
    new_presentation, blank_slide, set_slide_bg,
    add_shape, add_rounded_rect, add_oval,
    add_right_arrow, add_down_arrow,
    add_text_box, add_multiline_text,
    text_in_shape, multi_text_in_shape,
    add_bold_prefix_line,
    add_slide_title, add_bottom_bar,
)
```

### 1.1 Setup

| Function | Purpose |
|---|---|
| `new_presentation(width_in=13.333, height_in=7.5)` | Create a 16:9 widescreen `Presentation`. |
| `blank_slide(prs)` | Add a slide using the blank layout (index 6). Always prefer this — it avoids placeholder interference. |
| `set_slide_bg(slide, color)` | Solid background fill for the whole slide. |

### 1.2 Shape primitives

| Function | Purpose |
|---|---|
| `add_shape(slide, left, top, w, h, fill, border, border_width)` | Plain rectangle. Use for bands/footers/title rules. |
| `add_rounded_rect(slide, left, top, w, h, fill, border, border_width, corner=0.05)` | The main consultant "card". Use for any container that holds a titled block of content. |
| `add_oval(slide, left, top, w, h, fill, border, border_width)` | Circle — good for numbered step markers (`1`, `2`, `3`). |
| `add_right_arrow(slide, left, top, w, h, color)` | Arrow as a shape. Prefer this over `add_connector` — connectors with float EMU coordinates cause file corruption. |
| `add_down_arrow(slide, left, top, w, h, color)` | Vertical arrow, same rationale. |

### 1.3 Text

Two patterns, depending on whether you want text as its own element or inside an existing shape:

**Own element** (floats over the slide, no container):

| Function | Purpose |
|---|---|
| `add_text_box(slide, left, top, w, h, text, font_size, font_color, bold, alignment, font_name)` | Single paragraph. |
| `add_multiline_text(slide, left, top, w, h, lines, ...)` | Multiple paragraphs. Each item in `lines` is either a string (uses defaults) or a dict: `{"text": ..., "size": 11, "color": Palette.NAVY, "bold": True, "align": PP_ALIGN.LEFT, "indent": 0}`. This is the most powerful helper — use it for any structured list. |

**Inside a shape** (text with a background/border):

| Function | Purpose |
|---|---|
| `text_in_shape(shape, text, font_size, font_color, bold, alignment, font_name)` | Write one line into an existing shape (e.g., a card title). |
| `multi_text_in_shape(shape, lines, font_size, font_color, alignment, font_name)` | Write multiple lines into an existing shape, with sensible margins (8/8/6/6 pt). |

**Mixed-weight paragraph** (the only way to get `**Label:** value` in one line):

| Function | Purpose |
|---|---|
| `add_bold_prefix_line(text_frame, label, value, ...)` | Append a paragraph with a bold prefix + normal tail. Use inside a text frame created by `add_text_box` (pass `text_box.text_frame`). `python-pptx` does not support inline markdown — this is the workaround. |

### 1.4 Slide chrome

Apply these on every content slide so the deck looks cohesive:

| Function | Purpose |
|---|---|
| `add_slide_title(slide, title, subtitle, accent_color)` | Top accent bar + 26pt Navy title + optional 13pt gray subtitle + a short accent rule below. |
| `add_bottom_bar(slide, page_num, total, deck_title)` | Navy footer with deck title on the left and `page / total` on the right. |

### 1.5 Minimal working example

```python
from pptx.enum.text import PP_ALIGN
from pptx_helpers import (
    Palette, Fonts,
    new_presentation, blank_slide,
    add_slide_title, add_bottom_bar,
    add_rounded_rect, text_in_shape, multi_text_in_shape,
)

prs = new_presentation()
deck_title = "Quarterly Review"
total = 3

s = blank_slide(prs)
add_slide_title(s, "Executive Summary", "Q3 2026")
card = add_rounded_rect(s, left=Inches(0.8), top=Inches(2.0),
                        width=Inches(5.5), height=Inches(3.5),
                        fill_color=Palette.LIGHT_BLUE_BG,
                        border_color=Palette.BORDER_LIGHT)
multi_text_in_shape(card, [
    {"text": "Key Wins", "size": 14, "bold": True, "color": Palette.NAVY},
    "• Closed two enterprise deals",
    "• Shipped v2 of the analytics module",
    "• Reduced p95 latency by 38%",
])
add_bottom_bar(s, 1, total, deck_title=deck_title)

prs.save("review.pptx")
```

### 1.6 Units cheat-sheet

| What | How |
|---|---|
| Inches | `Inches(1.5)` |
| Points (font sizes, line widths) | `Pt(12)` |
| Colors | `RGBColor(0x1B, 0x2A, 0x4A)` |
| Never | Pass plain `float`/`int` where `Inches`/`Emu` is expected — PowerPoint rejects non-integer EMU. |

---

## 2. Verification scripts

Run all three after generation, before handing off. They're cheap and catch 95% of issues before the user opens PowerPoint.

### 2.1 `verify_structure.py`

```bash
python3 scripts/verify_structure.py deck.pptx           # just validate
python3 scripts/verify_structure.py deck.pptx 7         # also assert 7 slides
```

Checks file exists, is a valid ZIP (OPC), every XML part parses, `python-pptx` can reopen it, and no slide is empty. Exits non-zero on any hard failure — use the exit code in CI or in the agent repair loop.

### 2.2 `verify_layout.py`

```bash
python3 scripts/verify_layout.py deck.pptx
```

For each slide, reports:
- `[OUT OF BOUNDS]` — any shape extending past the canvas (hard fail).
- `[WARN overflow?]` — heuristic: text likely doesn't fit its container. Review manually.
- `[WARN overlap]` — two text-bearing shapes with IoU > 0.6. Often unintentional stacking.

Exits non-zero only on hard bounds violations; treat warnings as a review checklist.

### 2.3 `inspect_slide.py`

```bash
python3 scripts/inspect_slide.py deck.pptx 3
```

Dumps every shape on slide 3 with its type, name, bounding box (inches), and text content. Use this when the user says "slide 3 looks broken" — you can see exactly what's there without opening PowerPoint.

---

## 3. Recommended debug loop

When a deck is reported broken:

1. `python3 scripts/verify_structure.py deck.pptx` — is the file even valid?
2. `python3 scripts/verify_layout.py deck.pptx` — anything off-canvas or overflowing?
3. `python3 scripts/inspect_slide.py deck.pptx <N>` — read back the suspect slide.
4. Fix the generator script (never hand-edit the `.pptx`).
5. `rm -f deck.pptx && python3 create_ppt.py` — delete first; macOS won't always refresh.
6. Re-run 1 & 2 before declaring done.

See `SKILL.md` Phase 5 for the full root-cause table (`dash_style = 4`, float EMU in connectors, `Inches(i * Inches(x))`, etc.).
