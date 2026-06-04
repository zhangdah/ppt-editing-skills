---
name: build-consultant-ppt
description: Generate, debug, and iteratively refine consultant-style PowerPoint presentations using python-pptx. Use when the user wants to create a PPT deck, build slides programmatically, fix broken .pptx files, or iterate on presentation layouts. Covers the full workflow from initial generation, through self-verification, to repair-loop debugging when PowerPoint reports file corruption.
---

# Build Consultant-Style PowerPoint Decks

This skill covers the end-to-end workflow for programmatically generating clean, professional PowerPoint decks (consultant style: navy + accent colors, card-based layouts, minimal clutter) and debugging them when PowerPoint reports corruption or layout issues.

## When to Use

- User asks to create, generate, or build a PowerPoint / PPT / deck
- User wants slides with specific content but doesn't want to build them manually
- User reports a broken `.pptx` file (PowerPoint says "content has problems" or offers to "repair")
- User wants to iteratively refine an existing generated PPT (add/remove slides, adjust layout, fix overflow)

## Co-editing Protocol (read this before any run)

The generator overwrites the target `.pptx` every time it runs. There is no watcher and no merge logic — the script is the source of truth and a regeneration replaces the file wholesale. Before running the generator, tell the user:

1. **Close the deck in PowerPoint / Keynote / LibreOffice before regenerating.** PowerPoint locks the file while open; running the generator may fail to write, or — worse — succeed while PowerPoint holds a stale in-memory copy that overwrites your output the next time the user hits Save.
2. **Don't edit the deck in PowerPoint by hand and expect those edits to survive.** The next regeneration replaces the file. If the user wants tweaks to persist, the tweaks belong **in the generator script**, not in the binary `.pptx`.
3. **If the user did make manual edits in PowerPoint, they must tell the agent before the next regeneration.** Otherwise the script will silently discard their work.
4. **On every resumption, confirm whether the .pptx on disk is still the script's output, or whether the user has touched it since.** If touched, decide together: port the manual edits back into the script, or accept that the next regeneration will overwrite them.

When you (the agent) are about to call `prs.save(...)` or run a generator script, say something like: *"About to (re)write `deck.pptx`. If the deck is open in PowerPoint, please close it first."* This avoids a confusing "where did my changes go?" round-trip.

If the user has spent significant manual effort in PowerPoint and only wants a small text or font fix from here on, **don't regenerate** — that would discard their work. Stop, tell them this skill regenerates from the script, and let them decide whether to port the manual edits back into the script first.

## Core Workflow

Follow this ordered pipeline. Do not skip verification steps — they prevent the most common failure modes.

```
Task Progress:
- [ ] 1. Gather requirements and outline
- [ ] 2. Create a generator script (create_ppt.py)
- [ ] 3. Generate the .pptx file
- [ ] 4. Self-verify file validity and layout bounds
- [ ] 5. Deliver to user; on failure, enter repair loop
```

---

## Phase 1: Requirements

Before writing any code, confirm with the user:

- **Slide count and outline** — number of slides, title per slide, key points per slide
- **Style** — consultant/corporate style by default (navy + one accent color, white background); ask if unsure
- **Aspect ratio** — widescreen 16:9 (13.333 × 7.5 inches) is the default
- **Target audience** — affects tone (technical vs executive)
- **Output location** — workspace root by default

If the user provides an existing outline or transcript, use it directly; don't re-ask.

## Phase 2: Generator Script

**Always use `python-pptx`**. Create a single Python script (e.g., `create_ppt.py`) that generates the deck. This script becomes the source of truth — all future edits happen in the script, then regenerate.

### Install if missing

```bash
pip3 install python-pptx
```

### Script skeleton

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Consultant palette
NAVY = RGBColor(0x1B, 0x2A, 0x4A)
ACCENT_BLUE = RGBColor(0x2E, 0x86, 0xDE)
ACCENT_TEAL = RGBColor(0x00, 0xA8, 0x8F)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MID_GRAY = RGBColor(0x8C, 0x8C, 0x8C)
DARK_GRAY = RGBColor(0x4A, 0x4A, 0x4A)
LIGHT_BLUE_BG = RGBColor(0xE8, 0xF1, 0xFA)
BORDER_LIGHT = RGBColor(0xD0, 0xD5, 0xDD)

# ... build slides using helpers below ...

prs.save("output.pptx")
```

### Reusable helpers

For any non-trivial deck, define small helpers at the top of the script:
- `add_shape(slide, left, top, w, h, fill, border)` — rectangle
- `add_rounded_rect(slide, left, top, w, h, fill, border)` — rounded rectangle for cards
- `add_text_box(slide, left, top, w, h, text, font_size, color, bold, align)` — single-line text
- `add_multiline_text(slide, left, top, w, h, lines)` — structured multi-paragraph text where each line is a dict with `text`, `size`, `color`, `bold`
- `add_slide_title(slide, title, subtitle)` — consistent slide header
- `add_bottom_bar(slide, page_num, total)` — page footer

See [reference.md](reference.md) for full helper implementations.

### Consultant-style guidelines

- **Use `prs.slide_layouts[6]` (blank layout)** — avoid placeholders, position everything manually
- **Slide header**: title in Navy 26pt bold + subtitle in gray 13pt + a thin accent-color rule below
- **Card pattern**: rounded rect with light-color background + thin border, title in Navy bold, description in gray
- **Color restraint**: navy + 1–2 accent colors max; avoid using more than 3 colors on one slide
- **Typography**: Helvetica Neue / Arial; body 10–12pt, titles 14–26pt
- **Bottom bar**: thin navy bar with deck name on left, page number on right

## Phase 3: Generate

Run the script from the terminal:

```bash
python3 create_ppt.py
```

**Critical**: If the output file is open in PowerPoint, the write may silently fail on macOS (file stays locked). Always delete first if regenerating:

```bash
rm -f output.pptx && python3 create_ppt.py
```

Tell the user to close PowerPoint before regenerating if they've opened the file.

## Phase 4: Self-Verify (Mandatory)

Before telling the user "done", run automated verification. Skip this and you will ship broken files. There are three classes of failure to check:

### Check A: File validity

```bash
python3 -c "
from pptx import Presentation
prs = Presentation('output.pptx')
print(f'Slides: {len(prs.slides)}')
for i, slide in enumerate(prs.slides):
    n = sum(1 for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip())
    print(f'  Slide {i+1}: {n} text shapes')
"
```

If this raises an exception, the file is corrupt — go to Phase 5.

### Check B: XML validity (optional, for paranoia)

```bash
python3 -c "
import zipfile, xml.etree.ElementTree as ET
with zipfile.ZipFile('output.pptx') as z:
    for name in z.namelist():
        if name.endswith('.xml'):
            try:
                with z.open(name) as f: ET.parse(f)
            except ET.ParseError as e:
                print(f'BAD: {name}: {e}')
print('XML OK')
"
```

### Check C: Layout bounds

Verify every shape is within the slide. This catches the most common visible bug (content falling off the edge) that doesn't trigger a corruption error but looks broken.

```python
from pptx import Presentation
prs = Presentation('output.pptx')
sw = prs.slide_width / 914400
sh = prs.slide_height / 914400
for i, slide in enumerate(prs.slides):
    for shape in slide.shapes:
        l = shape.left / 914400
        t = shape.top / 914400
        r = l + shape.width / 914400
        b = t + shape.height / 914400
        if r > sw + 0.01 or b > sh + 0.01 or l < -0.01 or t < -0.01:
            text = shape.text_frame.text[:30] if shape.has_text_frame else ''
            print(f'Slide {i+1} OUT OF BOUNDS: [{l:.2f},{t:.2f}] → [{r:.2f},{b:.2f}] | {text}')
```

If anything is out of bounds, fix the generator script and regenerate.

### Check D: Visual rendering (recommended)

Coordinate math (Check C) only catches off-canvas bugs. For visual issues — bad alignment, overlapping text, font fallback, line wrapping mistakes — you need to actually **see** the deck. Use the bundled renderer to produce a self-contained HTML preview:

```bash
python3 .cursor/skills/build-consultant-ppt/scripts/render_to_html.py output.pptx
# wrote output_preview.html (N slide(s), ~K KB)
```

This calls LibreOffice headless to convert the deck to PDF, then renders each page as a PNG embedded into a single HTML file. The output is the same as opening the deck in PowerPoint — fonts, colors, line wraps, shape geometry — and works without any GUI app open.

Open the HTML in a browser, or read it back yourself by inspecting each PNG. If you see a problem (text overflowing a card, two shapes overlapping, wrong color), fix the generator script and regenerate.

Requirements: LibreOffice (`brew install --cask libreoffice` on macOS) and `pip install pymupdf`. If LibreOffice isn't installed, fall back to Check C only.

## Phase 5: Repair Loop

When PowerPoint reports "content has problems" or the user reports visible issues (missing nodes, overlapping text, wrong content), enter this loop. **Don't guess — diagnose systematically.**

### Step 1: Reproduce and classify

Ask the user (or read their screenshot/message) to identify:
- **Corruption**: PowerPoint shows "repair" dialog on open
- **Layout overflow**: content falls off slide or overlaps
- **Text cutoff**: container shapes smaller than their text
- **Wrong content**: stale data (usually because the file didn't actually update)

### Step 2: Inspect the generated file

Don't guess what's on the slide. Use one of two views, depending on what you need:

**Visual (best for layout/alignment issues):** render the deck to HTML and look at it.

```bash
python3 .cursor/skills/build-consultant-ppt/scripts/render_to_html.py output.pptx
open output_preview.html
```

**Coordinate dump (best for "which shape is at which EMU position"):**

```python
from pptx import Presentation
prs = Presentation('output.pptx')
slide = prs.slides[N]  # N = slide index with issue
for shape in slide.shapes:
    l, t = shape.left / 914400, shape.top / 914400
    w, h = shape.width / 914400, shape.height / 914400
    text = shape.text_frame.text[:40] if shape.has_text_frame else '[no-text]'
    print(f'[{l:.1f},{t:.1f}] {w:.1f}x{h:.1f} | {text!r}')
```

### Step 3: Common root causes and fixes

| Symptom | Likely cause | Fix |
|---------|--------------|-----|
| PowerPoint offers to "repair" | Invalid attribute value in XML (e.g., raw int passed to enum field like `line.dash_style = 4`) | Use the proper enum: `from pptx.enum.dml import MSO_LINE_DASH_STYLE; shape.line.dash_style = MSO_LINE_DASH_STYLE.DASH` |
| XML has non-integer EMU | Float arithmetic on `Inches()` values passed to `add_connector` or similar | Cast to `int()`, or use `MSO_SHAPE` arrows instead of connectors |
| Nodes disappear / overflow | `Inches(x + i * Inches(y))` pattern — multiplying an already-Inches value | Use plain floats inside `Inches(...)`: `Inches(0.6 + i * 2.55)` where both are floats |
| Text cut off in container | Background rect height < text frame height | Measure required height (font_size * num_lines * 1.2 / 72 + padding) and enlarge the container |
| File seemingly not updated | PowerPoint has the file open and is locking it | `rm -f file.pptx` first, and ask the user to close PowerPoint before regenerating |
| `repair` truncates content | PowerPoint's auto-repair often drops the offending shape entirely | After user runs "repair", the visible damage shows which shape was bad — that shape's generation code is the culprit |

### Step 4: Fix the generator, not the .pptx

Never hand-edit the `.pptx`. Always fix `create_ppt.py`, then:

```bash
rm -f output.pptx && python3 create_ppt.py
```

Re-run Phase 4 verification before handing back to the user.

### Step 5: Tell the user to reopen

PowerPoint **does not auto-refresh** open files. After regenerating, instruct:
> "Please close the file in PowerPoint and reopen it."

If they report the file still looks old, it means PowerPoint had it locked during save; see Phase 3 guidance.

---

## Known Gotchas

Keep these in mind — they cause the majority of failures.

### Gotcha 1: `Inches(i * Inches(x))` is wrong

```python
# WRONG — multiplies Emu by Emu
spacing = Inches(2.5)
x = Inches(0.6 + i * spacing)  # broken

# RIGHT — do math on floats, then wrap
x = Inches(0.6 + i * 2.5)
```

### Gotcha 2: Raw integers for enum fields silently break PowerPoint

`python-pptx` doesn't validate. If you set `line.dash_style = 4` (an int), the file saves fine, but PowerPoint will reject it on open. Always use the enum class.

### Gotcha 3: `add_connector` with float coordinates

Division like `db_x + db_w / 2` produces a float in EMU, which serializes as e.g. `10835640.0` in XML. PowerPoint requires integer EMU. Use `int(...)` or prefer `MSO_SHAPE.RIGHT_ARROW` / `DOWN_ARROW` shapes instead.

### Gotcha 4: Macos file locking

On macOS, a file open in PowerPoint is not locked against writes by Python, but the user's view won't update. On Windows, the write fails. Always `rm -f` before regenerating, and remind the user to close the file.

### Gotcha 5: Emoji in text

Some fonts don't render emoji, and they can confuse text-measurement. If the user wants icons, prefer a small colored shape (circle with a letter, or a unicode geometric like ▸, ●, ◆) over emoji.

### Gotcha 6: `text_frame.text = "a\nb"` collapses newlines

Setting `.text` directly only makes one paragraph. For multiline, add paragraphs explicitly:

```python
tf.text = "first line"
p = tf.add_paragraph()
p.text = "second line"
```

Or use a multiline helper that iterates and creates paragraphs.

---

## Output Conventions

- Save the generator script as `create_ppt.py` (or `build_<name>.py`) in the workspace root or a `scripts/` folder
- Save the deck with a descriptive name: `<Topic>_Presentation.pptx`
- After generation, summarize to the user: slides count, filename, how to view
- If iterating, keep the script — don't regenerate from scratch each turn; edit the relevant section with `StrReplace`

## When to Preview Without PowerPoint

If the user can't or won't open PowerPoint (e.g., to confirm a fix), use the read-back technique from Phase 5 Step 2 to describe what's on each slide. Do not fabricate screenshots or claim visual verification you haven't done.

## Additional Resources

- For the full reusable helper library, see [reference.md](reference.md)
