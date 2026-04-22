# ppt-editing-skills

Cursor Agent Skills for programmatically generating and debugging consultant-style PowerPoint decks with `python-pptx`.

## What's inside

```
.cursor/skills/build-consultant-ppt/
├── SKILL.md          # end-to-end workflow: requirements → generate → verify → repair loop
├── reference.md      # helper API docs, usage examples, units cheat-sheet
└── scripts/
    ├── pptx_helpers.py      # reusable helper library (palette, shapes, text, chrome)
    ├── verify_structure.py  # validate .pptx (ZIP / XML / python-pptx roundtrip)
    ├── verify_layout.py     # detect out-of-bounds shapes, overflow, overlap
    └── inspect_slide.py     # dump a slide to stdout for debugging without PowerPoint
```

## How Cursor picks it up

Cursor auto-discovers skills under `.cursor/skills/` in the workspace root. Clone this repo into any project (or drop the `build-consultant-ppt/` folder into an existing `.cursor/skills/`) and the agent will use it automatically when you ask to build / fix / refine a PowerPoint deck.

## Install the runtime dependency

```bash
pip install python-pptx
```

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
```

See [`SKILL.md`](.cursor/skills/build-consultant-ppt/SKILL.md) for the full workflow and [`reference.md`](.cursor/skills/build-consultant-ppt/reference.md) for helper API details.

## License

MIT
