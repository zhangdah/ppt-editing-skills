# Reference: XML-Direct PPT Editing

Companion to `SKILL.md`. This document explains the OOXML model the scripts work against, the exact patch grammar, and what each script does in detail.

Everything below is **general-purpose** — nothing about it assumes a specific topic, company, or deck.

---

## Directory layout

```
edit-pptx-template/
├── SKILL.md                  # workflow (read first)
├── reference.md              # this file
└── scripts/
    ├── xml_audit.py          # deck.pptx → audit.json (per-run, with stable IDs)
    ├── xml_patch.py          # deck.pptx + patch.json → edited.pptx
    ├── xml_diff.py           # diff two .pptx at OOXML level
    └── render_to_html.py     # LibreOffice → PDF → PNG → HTML preview
```

The scripts have **zero internal dependencies on each other** at runtime; each is a standalone CLI. They communicate exclusively through file artifacts (audit JSON, patch JSON, edited pptx, HTML).

Standard library only — except `render_to_html.py`, which uses `pymupdf` and the system `libreoffice`.

---

## 1. The OOXML model in two minutes

A `.pptx` is a zip. The interesting parts for this skill:

```
ppt/
  presentation.xml          # slide size, slide list
  slides/
    slide1.xml              # one file per slide
    slide2.xml
    ...
```

Inside `slideN.xml`, every text-bearing shape looks roughly like this:

```xml
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="4" name="TextBox 3"/>     <!-- shape ID is here -->
    <p:cNvSpPr txBox="1"/>                 <!-- txBox=1 → it's a TextBox -->
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>                                 <!-- DO NOT TOUCH (geometry/fill) -->
    <a:xfrm>
      <a:off x="1097280" y="1645920"/>     <!-- position (EMU) -->
      <a:ext cx="10058400" cy="457200"/>   <!-- size (EMU) -->
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square"/>
    <a:lstStyle/>
    <a:p>                                  <!-- paragraph 0 -->
      <a:pPr algn="l">                     <!-- ← set_default_font edits inside <a:defRPr> -->
        <a:defRPr sz="1400" b="1">
          <a:solidFill><a:srgbClr val="2E86DE"/></a:solidFill>
          <a:latin typeface="Helvetica Neue"/>
        </a:defRPr>
      </a:pPr>
      <a:r>                                <!-- run 0 -->
        <a:rPr ... />                      <!-- ← set_font edits these attrs -->
        <a:t>SYSTEM DESIGN PROPOSAL</a:t>  <!-- ← set_text edits this text -->
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>
```

### The three layers of font properties

A run's effective font is computed by walking up:

```
run's <a:rPr>           ← set on the run itself (highest priority)
  ↓ falls back to
paragraph's <a:defRPr>  ← inside <a:pPr>; set_default_font edits this
  ↓ falls back to
shape's <a:lstStyle>    ← rarely set; this skill leaves it alone
  ↓ falls back to
slide layout / master / theme
```

The audit reports both layers so you can decide which to edit:

- Change is paragraph-wide (or applies to all runs that don't override) → **`set_default_font` on `pN`**
- Change should apply to one specific run only → **`set_font` on `pN.rN`**
- Change is text content → **`set_text` on `pN.rN`**

---

## 2. `xml_audit.py`

```bash
python3 xml_audit.py deck.pptx [--out path] [--slides 1,3-5] [--pretty]
python3 xml_audit.py --check-stale deck.audit.json
```

### `--check-stale`

Cheap, side-effect-free probe to ask "is the deck on disk still the one this audit was taken from?"

- **Exit 0 (`[FRESH]`)** — the live file's `(size, mtime_ns, sha256)` all match the audit's recorded fingerprint. Safe to draft / apply patches.
- **Exit 1 (`[STALE]`)** — at least one of those has changed. Re-audit before going further.
- **Exit 2** — error (audit not found, malformed, missing fingerprint).

The agent should run this at the top of any turn where the user might have touched the deck in PowerPoint (after a pause, after a hand-off, on resume, on doubt).

### Output schema (one example slide)

```jsonc
{
  "deck": "deck.pptx",
  "deck_path": "/abs/path/deck.pptx",
  "fingerprint": {
    "size": 31013,
    "mtime_ns": 1714201234567890,
    "sha256": "5f9b6abe2d91e14b0e27a2bc5955fe983a793d73e0e3c6d26fdb4682bb4868ab"
  },
  "slide_size_in": [13.333, 7.5],
  "slide_count": 7,
  "slides": [
    {
      "id": "s1",
      "index": 1,
      "part": "ppt/slides/slide1.xml",
      "shape_count": 6,
      "shapes": [
        {
          "id": "s1.sp4",
          "name": "TextBox 3",
          "kind": "text_box",
          "bbox_in": [1.2, 1.8, 11.0, 0.5],
          "paragraphs": [
            {
              "id": "s1.sp4.p0",
              "align": "l",
              "level": 0,
              "default_font": {
                "size_pt": 14.0,
                "bold": true,
                "italic": null,
                "underline": null,
                "strike": null,
                "color": "2E86DE",
                "font_name": "Helvetica Neue",
                "font_name_ea": null,
                "font_name_cs": null
              },
              "runs": [
                {
                  "id": "s1.sp4.p0.r0",
                  "text": "SYSTEM DESIGN PROPOSAL",
                  "font": { "size_pt": null, "bold": null, ... }
                }
              ]
            }
          ]
        }
      ]
    }
  ]
}
```

### Reading the output

- `font.size_pt: null` on a run = **inherited from the paragraph's `default_font`**, not "missing"
- `kind` is one of `text_box`, `rect`, `rounded_rect`, `oval`, `right_arrow`, etc., taken from `<a:prstGeom prst="...">`. Only text-bearing shapes are emitted; pictures, connectors, and shapes with no `<p:txBody>` are skipped.
- `bbox_in` is `[left, top, width, height]` in inches (EMU/914400) for context only — this skill never edits these
- `<a:br>` line breaks are reported as runs with `"kind": "line_break"` and `"text": "\n"` so the run-index matches what `xml_patch.py` uses
- `fingerprint` captures the deck's identity at audit time. Two-stage compare: the cheap `(size, mtime_ns)` rules out "definitely changed"; the `sha256` confirms equality even when an app preserves mtime across a save. Don't hand-edit it — the staleness gate relies on it.

---

## 3. `xml_patch.py`

```bash
python3 xml_patch.py deck.pptx patch.json --out edited.pptx
python3 xml_patch.py deck.pptx patch.json --in-place
python3 xml_patch.py deck.pptx patch.json --out edited.pptx --force
```

### Safety gates (run before any byte is written)

1. **Lock-holder probe** — checks for a sibling `~$deck.pptx` Office lockfile, runs `lsof` on the target path (POSIX), or attempts an exclusive open on Windows. If anything else holds the file, refuses to write and exits 2 with a message identifying the holder. Implemented in `_pptx_lock.py`.

2. **Audit-staleness gate** — only triggered when the patch is in object form with `based_on` set. Loads the referenced audit, compares its `fingerprint` against the deck on disk, and refuses to apply if they differ (exits 2 with the specific reason: size / mtime / sha mismatch). Implemented in `_pptx_fingerprint.py`.

`--force` skips both gates. Use it only in non-interactive contexts where you've validated the situation by other means.

### Patch grammar

A patch may be either a JSON array of ops (legacy) or a JSON object `{ "based_on"?, "ops": [...] }` (recommended).

```jsonc
{
  "based_on": "deck.audit.json",  // optional; relative to this patch file
  "ops": [
    {
      "op": "set_text",                      // required
      "target": "s{slide}.sp{id}.p{n}.r{n}", // required, must be a run for set_text
      "value": "new literal string"          // required, must be a string
    },
    {
      "op": "set_font",                      // run-level font override
      "target": "s1.sp4.p0.r0",
      "value": {                             // any subset of these keys
        "size_pt": 36,                       // number > 0
        "bold": true,                        // true / false / null
        "italic": false,
        "underline": "sng",                  // bool or "sng"|"dbl"|"none"|...
        "strike": "sngStrike",               // bool or "noStrike"|"sngStrike"|"dblStrike"
        "color": "FFFFFF",                   // 6-hex; uppercased automatically
        "font_name": "Inter",
        "font_name_ea": "PingFang SC",
        "font_name_cs": null                 // null = remove the attribute
      }
    },
    {
      "op": "set_default_font",              // paragraph-level default font
      "target": "s1.sp7.p1",                 // points at a paragraph (no .rN)
      "value": { ... same keys as above ... }
    }
  ]
}
```

### Validation rules

- Unknown `op` → reject (the patcher only supports `set_text`, `set_font`, `set_default_font`)
- Missing `op` / `target` / `value` → reject
- `set_text` target without `.rN` → reject
- `set_text` value not a string → reject
- `set_text` on a `<a:br>` run → reject (line breaks have no text)
- `set_default_font` target with `.rN` → reject (must be paragraph-level)
- `color` not 6-hex → reject
- `size_pt` ≤ 0 or non-numeric → reject
- `target` references a slide / shape / paragraph / run that doesn't exist → reject with a precise "out of range" message

### Byte-level guarantees

The patcher writes the new zip such that:

- **Every member except modified `slideN.xml` parts is byte-identical to the source** (read raw, written raw)
- Member ordering, compression, dates, and external attributes are preserved
- The opening `<p:sld>` tag of edited slides is restored from the source, then unioned with any `xmlns:*` prefixes the serializer needed to put on the root because they were originally declared inline (e.g. `p14:`, `a16:`, `mc:`). This is a semantically-equivalent rewrite — every prefix that bound to a URI in the source still binds to the same URI in the output.
- All other namespace declarations and root attributes never drift

These guarantees are what `xml_diff.py` is designed to check.

### Inheritance pitfalls

- Setting a property on a run **adds an `<a:rPr>` element if missing**. The added element appears before any existing `<a:t>` (correct OOXML order).
- Setting `color: null` on a run that has its own `<a:solidFill>` removes only that element — the run then inherits the paragraph's color.
- Setting fields on `set_default_font` for a paragraph that has no `<a:pPr>` will **create** the `<a:pPr>` and `<a:defRPr>` elements as needed.
- The audit always shows current state, so re-run `xml_audit.py` on the edited deck if you suspect a chain is mis-resolving.

---

## 4. `xml_diff.py`

```bash
python3 xml_diff.py before.pptx after.pptx [--only ppt/slides/] [--full]
```

### Default mode

Prints a one-line summary per changed zip member, plus a capped (12 lines) list of `+`/`-` line changes for each changed XML part. Returns **exit code 0** if the decks are byte-identical (within `--only`), **exit code 1** otherwise.

### `--full`

Prints the complete unified diff for every changed XML/text part. Use when the summary's preview gets truncated (`... +N more lines`) or when you need to verify a complex multi-op patch.

### What "clean" looks like

A clean patch produces:

- Changes only in `ppt/slides/slideN.xml` parts
- Changes match the patch ops 1:1 (one text replaced per `set_text`, one or more attributes per `set_font`/`set_default_font`)
- No `+` / `-` lines in `ppt/theme/`, `ppt/slideLayouts/`, `ppt/slideMasters/`, `_rels/`, `[Content_Types].xml`

Any deviation from the above means the patcher (or the patch) did something unexpected; investigate before shipping.

---

## 5. `render_to_html.py`

```bash
python3 render_to_html.py deck.pptx [--out preview.html] [--zoom 1.5]
```

Pipeline: **PPTX → LibreOffice headless → PDF → PyMuPDF (PNG per page) → embed all PNGs into one self-contained HTML**.

Requirements:

- `libreoffice` on PATH (`brew install --cask libreoffice` on macOS, `apt install libreoffice` on Linux)
- `pip install pymupdf`

The output is a single HTML file the user can open directly. It's the high-fidelity visual oracle for "does this actually look right after the edit?"

---

## Recommended debug loop

```
1. xml_audit.py deck.pptx --pretty
2. (read JSON, draft patch.json)
3. xml_patch.py deck.pptx patch.json --out edited.pptx
4. xml_diff.py deck.pptx edited.pptx --only ppt/slides/
5. xml_audit.py edited.pptx --pretty   # confirm new state matches intent
6. render_to_html.py edited.pptx       # visual check
```

If step 4 shows changes outside `ppt/slides/`, **stop**. The patcher should never change those parts; investigate before delivering.

If step 5 shows the new state but step 6 looks wrong, the issue is almost always inheritance — re-read the audit and check whether a parent paragraph or layout is overriding what you think you set.

---

## Hard boundaries (must not violate)

The patcher contract guarantees that, for every `--out` it writes:

1. Only `ppt/slides/slideN.xml` parts may differ from the source
2. Within those parts, the only mutable elements are `<a:t>` text content, `<a:rPr>` attributes/children, and `<a:pPr>/<a:defRPr>` attributes/children
3. The zip member list, member order, member compression mode, and `[Content_Types].xml` are byte-identical to the source
4. No element is added or removed except: `<a:rPr>` may be created if missing; `<a:pPr>`/`<a:defRPr>` may be created if missing; `<a:solidFill>`/`<a:srgbClr>`/`<a:latin>`/`<a:ea>`/`<a:cs>` children may be created or removed inside `<a:rPr>` or `<a:defRPr>`

If any future change to the scripts would relax one of these, that change does not belong here — it's a different tool with a different contract.
