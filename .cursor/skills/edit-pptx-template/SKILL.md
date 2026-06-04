---
name: edit-pptx-template
description: Surgically edit text content and font properties inside an existing .pptx file by directly modifying its OOXML. Use when the user wants to change wording, rename a title, swap a font, recolor text, adjust font size, or toggle bold/italic/underline on a deck that already exists, without altering layout, shape geometry, fills, images, charts, slide masters, or themes. Provides byte-level guarantees that only the requested fields change. Pure standard library — no python-pptx, no lxml.
---

# Edit an Existing PowerPoint Deck (XML-direct)

This skill is for one job: **changing text and font properties inside a deck that already exists, with surgical precision and zero side effects on the rest of the file.**

It works by directly editing the XML inside the `.pptx` zip. Pure standard library; no `python-pptx`, no `lxml`, no plugins.

## When to Use

- User wants to **change wording** in an existing deck ("rename slide 3's title to ...")
- User wants to **change a font** ("make all body text Inter instead of Helvetica Neue")
- User wants to **recolor text** ("the 'March 2026' should be red")
- User wants to **resize text** ("bump the headline up to 44pt")
- User wants to **toggle bold / italic / underline / strike** on specific runs
- User says "don't touch the layout, just fix the wording"

## Co-editing Protocol (read this before any run)

The user and the agent share one `.pptx` on disk. Two mechanical hazards must be guarded:

1. **Write conflict** — PowerPoint / Keynote / LibreOffice hold the file open while editing. Writing the same path during that window either fails outright or is silently overwritten the next time the user hits Save.
2. **Stale audit** — `audit.json` is a one-shot snapshot of the deck. If the user opens the deck, edits something, and saves, that snapshot no longer matches reality, and any patch drafted against it may hit stale text or vanished IDs.

The scripts now defend against **both** automatically. You do not need to ask the user to "remember to close PowerPoint" or "let me know when you save" — the tooling will surface the conflict and refuse to proceed.

### Built-in safety gates

| Gate | When | What it checks | Output |
| --- | --- | --- | --- |
| **lock-holder probe** | Before `xml_patch.py` writes anything | Is the target path open in PowerPoint, Keynote, LibreOffice, or any other process? (Detects the `~$deck.pptx` Office lockfile and queries `lsof` on POSIX.) | Refuses to write; exits 2; names the holder PID/command and the action needed. |
| **audit fingerprint** | Stamped into every `xml_audit.py` output (`fingerprint: {size, mtime_ns, sha256}`) | n/a — just metadata. | Stored under the `fingerprint` key. |
| **stale-audit gate** | Before `xml_patch.py` applies a patch with `based_on: "deck.audit.json"` | Does the deck on disk still match the audit's fingerprint? | Refuses to apply; exits 2; tells you what changed (size / mtime / sha) and to re-audit. |
| **on-demand probe** | Whenever you want to know if your audit is still valid | Compares fingerprint without any side effects | `xml_audit.py --check-stale path/to/audit.json` → exit 0 = FRESH, exit 1 = STALE. |

`--force` on `xml_patch.py` bypasses both the lock and stale gates. Reserve it for cases where you are certain (e.g. CI pipelines where no human is editing).

### Recommended agent loop

When the user hands off to PowerPoint, then comes back and asks for another edit, do this **before drafting any new patch**:

```bash
python3 scripts/xml_audit.py --check-stale path/to/deck.audit.json
```

- **Exit 0 (FRESH)** — your audit is still valid; proceed.
- **Exit 1 (STALE)** — the deck has changed. Stop, re-audit, show the user the new state, confirm what they want changed, then draft a fresh patch.

The check is essentially free (single sha256 of the deck), so run it any time you're unsure: at the start of a turn, after a long pause, after the user mentions PowerPoint, etc.

### Tying patches to audits

To make the stale-audit gate enforce itself, write your patches in the metadata form rather than the bare-array form:

```json
{
  "based_on": "deck.audit.json",
  "ops": [
    {"op": "set_text", "target": "s1.sp4.p0.r0", "value": "REVISED PROPOSAL"}
  ]
}
```

`based_on` is resolved relative to the patch file's own directory. With it set, `xml_patch.py` will refuse to apply if the deck has drifted from the audit. This is the recommended form for any serious work; the bare-array form (just `[...ops...]`) is still accepted for quick one-off scripting.

## Scope of This Skill

This skill performs exactly three kinds of edit, expressed as patch operations:

| op | what it changes |
| --- | --- |
| `set_text` | the literal characters inside one run (`<a:t>`) |
| `set_font` | font attributes on one run (`<a:rPr>`) |
| `set_default_font` | the default font on one paragraph (`<a:pPr>/<a:defRPr>`) |

If the requested change cannot be expressed as one of these three operations, this skill is the wrong tool — for example moving / resizing / rotating shapes, adding or deleting shapes / slides / paragraphs, changing shape fills, borders, line styles, slide layout, master, or theme, inserting images / charts / tables. Tell the user what's out of scope and stop; do not bend the patcher into a layout editor.

The hard rule: **if the change isn't expressible as `set_text` / `set_font` / `set_default_font`, this skill does not perform it.**

---

## Core Workflow

```
Task Progress:
- [ ] 1. Audit the deck → JSON with stable IDs
- [ ] 2. Show the user the audit (or a relevant subset) and confirm what to change
- [ ] 3. Write a patch.json with one entry per change
- [ ] 4. Apply the patch → produces edited deck
- [ ] 5. Diff before vs. after at OOXML level → confirm only intended bytes moved
- [ ] 6. Render edited deck to HTML → human visual check
- [ ] 7. Deliver the edited deck, the patch, the diff, and the preview to the user
```

Each step has its own script in `scripts/`. They share a single ID convention so output of one feeds into the next without parsing.

---

## Phase 1: Audit

Run `scripts/xml_audit.py` to dump every text-bearing element with stable IDs and current font properties.

```bash
python3 scripts/xml_audit.py path/to/deck.pptx --pretty
# writes path/to/deck.audit.json
```

Useful flags:

- `--slides 1,3-5` — restrict to specific slides
- `--out custom.json` — custom output path

The audit is the **single source of truth** shared between you (the agent), the user (via JSON they can read), and the next steps. Every shape, paragraph, and run has a deterministic ID derived from the XML itself, not from positional ordering:

```
s{slide}                                # slide N (1-based)
s{slide}.sp{cNvPr_id}                   # shape (id is from <p:cNvPr id="...">)
s{slide}.sp{cNvPr_id}.p{para_idx}       # paragraph (0-based)
s{slide}.sp{cNvPr_id}.p{para_idx}.r{run_idx}  # run (0-based)
```

These IDs are stable: adding/removing other shapes does **not** renumber them, because `cNvPr_id` is the shape's persistent OOXML id.

### What the audit reveals

For every run, both the literal `<a:rPr>` attributes AND the inherited `default_font` (from the paragraph's `<a:pPr>/<a:defRPr>`) are reported. If a run shows `font.size_pt: null` but its paragraph's `default_font.size_pt: 14.0`, that means the run inherits 14pt — there's no need to set it on the run.

This matters when planning patches: prefer editing the **default_font** when the change should affect a whole paragraph (or all runs that don't override), and edit the **run's font** when you want a single run-level override.

---

## Phase 2: Confirm with the User

Show the user the relevant section of the audit (use code blocks). Highlight:

1. **The current values** of fields they want to change
2. **The IDs you'll target**
3. **Inheritance reality** — e.g. "this run has no font set, it inherits 14pt bold from its paragraph; do you want the change to apply to the whole paragraph or just this run?"

Only proceed once the user has confirmed scope. This is cheap insurance against painful re-runs.

---

## Phase 3: Write a Patch

A patch is a JSON document with a list of operations. The three allowed operations were summarized in *Scope* above; here is the precise targeting:

| op | target shape | what it changes |
| --- | --- | --- |
| `set_text` | a run (`...rN`) | the literal characters in `<a:t>` |
| `set_font` | a run (`...rN`) | attributes on that run's `<a:rPr>` |
| `set_default_font` | a paragraph (`...pN`) | attributes on the paragraph's `<a:pPr>/<a:defRPr>` |

### Two accepted forms

**Recommended form** — object with `based_on` so the staleness gate kicks in:

```json
{
  "based_on": "deck.audit.json",
  "ops": [
    {
      "op": "set_text",
      "target": "s1.sp4.p0.r0",
      "value": "REVISED PROPOSAL"
    },
    {
      "op": "set_font",
      "target": "s1.sp5.p0.r0",
      "value": {"italic": true, "size_pt": 36, "color": "FFFFFF", "font_name": "Inter"}
    },
    {
      "op": "set_default_font",
      "target": "s1.sp7.p1",
      "value": {"size_pt": 13, "color": "8C8C8C"}
    }
  ]
}
```

**Bare-array form** — accepted for quick one-offs; no staleness gate:

```json
[
  {"op": "set_text", "target": "s1.sp4.p0.r0", "value": "REVISED PROPOSAL"}
]
```

`based_on` is resolved relative to the patch's own directory. Path it to the audit you drafted the patch from, and the patcher will refuse to apply if the deck has drifted.

### Font field semantics

| field | type | notes |
| --- | --- | --- |
| `size_pt` | number > 0 | stored as `sz="<pt*100>"` |
| `bold` / `italic` | bool | stored as `b="1"`/`i="1"` |
| `underline` | bool or string | `true`→`sng`; pass `"dbl"`/`"none"`/etc. for specifics |
| `strike` | bool or string | `true`→`sngStrike`; pass `"noStrike"`/`"dblStrike"` for specifics |
| `color` | 6-hex string (`"FF0000"`) | uppercased; rejects anything else |
| `font_name` | string | sets `<a:latin typeface="...">` |
| `font_name_ea` / `font_name_cs` | string | East Asian / Complex Script typefaces |
| **any field set to `null`** | — | **REMOVES** the attribute (the run inherits from parent) |
| **any field omitted** | — | **untouched** (current value preserved) |

That `null = remove` distinction is the single most useful rule. To "make this run inherit color from its paragraph", set `"color": null`.

---

## Phase 4: Apply the Patch

```bash
python3 scripts/xml_patch.py path/to/deck.pptx path/to/patch.json --out edited.pptx
# or, to overwrite in place:
python3 scripts/xml_patch.py path/to/deck.pptx path/to/patch.json --in-place
```

The patcher:

1. Loads the deck zip and the patch
2. Validates the patch (rejects unknown ops, missing fields, bad colors, illegal targets)
3. Groups ops by slide; for each modified slide:
   - parses `slideN.xml`
   - mutates only the targeted `<a:t>` / `<a:rPr>` / `<a:defRPr>` nodes
   - re-serializes
   - scans the source XML for every `xmlns:*` declaration (including ones declared inline on extension elements like `p14:`, `a16:`) and registers them, so prefixes are preserved
   - restores the original root `<p:sld>` opening tag verbatim, unioning in any prefixes the serializer hoisted to the root
4. Writes a new zip, copying every other member byte-for-byte

If anything goes wrong — invalid id, missing run, unknown op — the patcher exits non-zero with a clear message and writes nothing.

---

## Phase 5: Diff Before vs. After

Always diff. This is the contract: prove that only the requested attributes / texts moved.

```bash
python3 scripts/xml_diff.py path/to/deck.pptx edited.pptx --only ppt/slides/
```

A clean diff for the patch above looks like:

```
--- changed ---
  [xml]    ppt/slides/slide1.xml: 9 changed line(s)
      -              <a:t>SYSTEM DESIGN PROPOSAL</a:t>
      +              <a:t>REVISED PROPOSAL</a:t>
      +              <a:rPr i="1"/>
      ...
```

If you see any change in `ppt/theme/`, `ppt/slideLayouts/`, `ppt/slideMasters/`, or any unrelated `<a:xfrm>` / `<a:prstGeom>` block, **stop and investigate** — something went wrong.

Use `--full` to see the complete unified diff for any changed XML part.

### Expected non-content noise

When the source had inline `xmlns:foo` declarations on a child element (common with `p14:creationId`, `a16:colId`, `a16:rowId`, `mc:AlternateContent`, etc.), the patcher hoists those declarations to the root `<p:sld>` element. You will see corresponding `+ xmlns:foo="..."` on the root and `- xmlns:foo="..."` on the original child in the diff. This is semantically a no-op — both forms describe the same XML — and PowerPoint, LibreOffice and `python-pptx` all accept it.

---

## Phase 6: Render and Verify

```bash
python3 scripts/render_to_html.py edited.pptx --out edited.html
```

Renders via LibreOffice headless: PPTX → PDF → one PNG per slide, embedded into a single self-contained HTML file for visual review.

Requirements: `libreoffice` on PATH (`brew install --cask libreoffice` on macOS, `apt install libreoffice` on Linux) and `pip install pymupdf` in your venv.

If the diff is clean but the visual is wrong, the issue is almost always inheritance — e.g. the run inherits a property you didn't realize. Re-audit the edited deck (`xml_audit.py edited.pptx`) and compare.

---

## Phase 7: Deliver

Tell the user:

1. **Path to edited deck**
2. **One-line summary** of what changed (e.g. "renamed s1.sp4.p0.r0 to 'REVISED PROPOSAL', italicized s1.sp5.p0.r0, recolored s1.sp7.p1 to red")
3. **Path to the HTML preview**
4. **Diff result** — "5 changed lines in ppt/slides/slide1.xml, all expected; nothing else touched"

Keep the patch.json next to the edited deck — it's the receipt of what changed and lets the user roll forward by editing the patch and re-running.

---

## Repair Loop (when something looks wrong)

If the visual or audit doesn't match expectation:

1. **Re-audit the edited deck** — what does the actual output XML say?
2. **Compare to your patch** — did each op land on the right ID?
3. **Check inheritance** — is the run's effective style coming from a parent you didn't touch?
4. **Diff with `--full`** — see the complete change set; look for anything you didn't ask for
5. **Fix the patch and re-apply against the original deck**, not the broken edited one

If the edit the user wants is structurally impossible with the three allowed ops — for example they want to move a shape, change a fill, or insert a new slide — stop. Tell them that's outside this skill's scope and what they need instead. Do not stretch the patcher into a generator; its value comes from being narrow and predictable.

---

## Boundaries (read this before pushing)

This skill **must not**:

- Touch `ppt/theme/*`, `ppt/slideLayouts/*`, `ppt/slideMasters/*`, `ppt/notesSlides/*`
- Modify `<a:xfrm>` (position/size), `<a:prstGeom>` (shape geometry), `<p:spPr>` fills/lines outside of text properties
- Add or remove any `<p:sp>`, `<a:p>`, `<a:r>`, or `<a:br>` elements
- Reorder existing elements
- Change zip member names, ordering, or compression of unmodified parts

If a future change to the scripts would require relaxing one of the above, that change does not belong in this skill — it's a different tool with a different contract.
