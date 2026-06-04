# ppt-editing-skills

Cursor Agent Skills for two complementary jobs: generating new consultant-style PowerPoint decks with `python-pptx`, and surgically editing existing decks by direct OOXML manipulation.

## What's inside

```
.cursor/skills/build-consultant-ppt/        # build new decks from scratch
├── SKILL.md          # end-to-end workflow: requirements → generate → verify → repair loop
├── reference.md      # helper API docs, usage examples, units cheat-sheet
└── scripts/
    ├── pptx_helpers.py      # reusable helper library (palette, shapes, text, chrome)
    ├── verify_structure.py  # validate .pptx (ZIP / XML / python-pptx roundtrip)
    ├── verify_layout.py     # detect out-of-bounds shapes, overflow, overlap
    ├── inspect_slide.py     # dump a slide to stdout for debugging without PowerPoint
    └── render_to_html.py    # render the deck to a self-contained HTML preview

.cursor/skills/edit-pptx-template/          # edit text + fonts of an existing deck
├── SKILL.md          # workflow: audit → patch → diff → render
├── reference.md      # OOXML model, patch grammar, byte-level guarantees
└── scripts/
    ├── xml_audit.py         # deck.pptx → audit.json with stable per-run IDs
    ├── xml_patch.py         # deck.pptx + patch.json → edited.pptx (XML-direct, std lib only)
    ├── xml_diff.py          # OOXML-level diff between two .pptx files
    └── render_to_html.py    # render the deck to a self-contained HTML preview
```

Pick by intent:
- "build me a new deck" → `build-consultant-ppt`
- "rename this title" / "change this font" / "recolor this text" → `edit-pptx-template`

## How Cursor picks it up

Cursor auto-discovers skills under `.cursor/skills/` in the workspace root. Clone this repo into any project (or copy individual skill folders into an existing `.cursor/skills/`) and the agent will use them automatically when the task matches.

## Install the runtime dependencies

For `build-consultant-ppt` (generating new decks):

```bash
pip install python-pptx
```

For visual preview via `render_to_html.py` (used by both skills):

```bash
pip install pymupdf
brew install --cask libreoffice    # macOS
# or: apt install libreoffice      # Debian/Ubuntu
```

`edit-pptx-template`'s patcher itself is pure standard library — `pymupdf` and LibreOffice are only needed if you want the rendered HTML preview.

## Quick start in your own generator script

```python
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN

# Add the skill's scripts/ to your path, or copy pptx_helpers.py next to your script.
from pptx_helpers import (
    Palette, new_presentation, blank_slide,
    add_slide_title, add_bottom_bar,
    add_rounded_rect, multi_text_in_shape,
)

prs = new_presentation()
s = blank_slide(prs)
add_slide_title(s, "Executive Summary", "Q3 2026")
card = add_rounded_rect(s, Inches(0.8), Inches(2.0), Inches(5.5), Inches(3.5),
                        fill_color=Palette.LIGHT_BLUE_BG,
                        border_color=Palette.BORDER_LIGHT)
multi_text_in_shape(card, [
    {"text": "Key Wins", "size": 14, "bold": True, "color": Palette.NAVY},
    "• Closed two enterprise deals",
    "• Shipped v2 of the analytics module",
    "• Reduced p95 latency by 38%",
])
add_bottom_bar(s, 1, 3, deck_title="Quarterly Review")
prs.save("review.pptx")
```

## Verify before you ship

```bash
python3 .cursor/skills/build-consultant-ppt/scripts/verify_structure.py review.pptx
python3 .cursor/skills/build-consultant-ppt/scripts/verify_layout.py    review.pptx
python3 .cursor/skills/build-consultant-ppt/scripts/inspect_slide.py    review.pptx 1
python3 .cursor/skills/build-consultant-ppt/scripts/render_to_html.py   review.pptx
open review_preview.html
```

See [`SKILL.md`](.cursor/skills/build-consultant-ppt/SKILL.md) for the full workflow and [`reference.md`](.cursor/skills/build-consultant-ppt/reference.md) for helper API details.

## Editing an existing deck (text and fonts, byte-clean elsewhere)

```bash
# 1. audit the deck — emits stable IDs and current font properties as JSON
python3 .cursor/skills/edit-pptx-template/scripts/xml_audit.py existing.pptx --pretty

# 2. read existing.audit.json, draft a patch.json with set_text / set_font / set_default_font ops

# 3. apply
python3 .cursor/skills/edit-pptx-template/scripts/xml_patch.py existing.pptx patch.json --out edited.pptx

# 4. prove only the requested fields moved
python3 .cursor/skills/edit-pptx-template/scripts/xml_diff.py  existing.pptx edited.pptx --only ppt/slides/

# 5. visual sanity check
python3 .cursor/skills/edit-pptx-template/scripts/render_to_html.py edited.pptx
```

The patcher uses pure standard library (no `python-pptx`, no `lxml`) and guarantees that only the run-level text and font attributes you target change — every other byte in the `.pptx` is preserved exactly, including layouts, masters, themes, images, charts, tables, and embedded relationships.

See [`edit-pptx-template/SKILL.md`](.cursor/skills/edit-pptx-template/SKILL.md) for the full workflow and [`edit-pptx-template/reference.md`](.cursor/skills/edit-pptx-template/reference.md) for the OOXML model and patch grammar.

## License

MIT
