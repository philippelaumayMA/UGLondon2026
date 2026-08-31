# UG London 2026 Deck Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a 16-slide (+4 appendix) English `.pptx` for a 30-minute talk, reusing six slides from the existing Moody's deck and authoring ten new ones on the same template.

**Architecture:** Duplicate `resources/Consolidated GenAI common slides.pptx` into `deck/`, reorder its slides into the spine, then author new slides using layouts already in that file. Nothing is copied between the two source `.pptx` files. A `tools/verify_deck.py` script holds the expected slide order as data and is extended one task at a time — that is the red/green cycle for a deliverable with no unit tests.

**Tech Stack:** Python 3.11, python-pptx 0.6.23, LibreOffice (`soffice`) for render checks.

**Spec:** `docs/superpowers/specs/2026-08-31-ug-london-deck-design.md`

## Global Constraints

- **English only.** All new text is English; the DEVOXX source is French and must be translated, never pasted.
- **Template fidelity.** New slides use layouts already present in the working deck, by name. Never construct a slide from a blank layout.
- **No cross-file slide copying.** Nothing is imported from `DEVOXX_2026_les_gardiens_du_prompt.pptx`; only re-authored English text derived from it.
- **Product name is `RiskIntegrity for IFRS 17 Navigator`** — the README's name, not "Infoweb Navigator". One place to change if that turns out wrong: `EXPECTED` in `tools/verify_deck.py` and the two slides in Tasks 6.
- **Never invent product facts.** Use case slides carry bracketed prompts in guillemets (`‹ … ›`) for the author to fill. No claim about what Navigator, UVA or the MCP server does may be written by the implementer — none of it is in this repo.
- **Audience-facing vs. speaker-facing.** Fallback plans, timings and cut marks go in speaker notes, never on the slide face.
- **Final state:** 20 slides — 16 main line (index 0–15) then 4 appendix (index 16–19).

---

### Task 1: Working copy and verification harness

**Files:**
- Create: `tools/deck.py`
- Create: `tools/verify_deck.py`
- Create: `.gitignore`
- Create: `deck/UG_London_2026.pptx` (copied binary)

**Interfaces:**
- Consumes: nothing.
- Produces: `tools/deck.py` exposing `open_deck(path) -> Presentation`, `find_layout(prs, name) -> SlideLayout`, `describe_layout(prs, name) -> None`, `reorder(prs, order: list[int]) -> None`, `add_slide_at(prs, layout_name: str, index: int) -> Slide`, `set_ph(slide, idx: int, text: str) -> bool`, `set_ph_lines(slide, idx: int, lines: list[str]) -> bool`, `drop_empty_placeholders(slide) -> None`, `set_notes(slide, text: str) -> None`, `slide_text(slide) -> str`. `tools/verify_deck.py` exposes module-level `EXPECTED: list[tuple[str, str | None]]` and is runnable as `python3 tools/verify_deck.py`.

- [ ] **Step 1: Write the verification script**

Create `tools/verify_deck.py`:

```python
#!/usr/bin/env python3
"""Assert deck/UG_London_2026.pptx matches the spec's slide map.

EXPECTED is the source of truth: one (signature, layout_name) pair per slide,
in order. `signature` is a distinctive substring that must appear somewhere on
the slide, matched case-insensitively. `layout_name` is checked only when not
None -- reused slides keep whatever layout they arrived with.

Signatures deliberately avoid apostrophes: the deck uses curly quotes (U+2019)
throughout, so "Moody's" would never match.
"""
import re
import sys

from pptx import Presentation

DECK = "deck/UG_London_2026.pptx"

EXPECTED = [
    ("Process Automations", None),        # C-1
    ("Driven by Perspective", None),      # C-2
    ("Content Foundation", None),         # C-3
    ("Every Stage of AI Maturity", None), # C-4
    ("AI Solution Suite", None),          # C-5
    ("in three layers", None),            # C-6
    ("ACTUARIAL AI CURVE", None),         # C-7
    ("OUR APPROACH", None),               # C-8
    ("RESPONSIBLE AI", None),             # C-9
    ("WHY MOODY", None),                  # C-10
]


def slide_text(slide):
    """All text on a slide, including inside groups and tables.

    Reads the raw XML rather than walking shapes: python-pptx does not descend
    into grouped shapes, and several source slides are heavily grouped.
    """
    return " | ".join(re.findall(r"<a:t>(.*?)</a:t>", slide._element.xml, re.S))


def main():
    prs = Presentation(DECK)
    slides = list(prs.slides)
    errors = []

    if len(slides) != len(EXPECTED):
        errors.append(f"slide count is {len(slides)}, expected {len(EXPECTED)}")

    for i, (sig, layout) in enumerate(EXPECTED):
        if i >= len(slides):
            errors.append(f"slide {i}: missing entirely (expected {sig!r})")
            continue
        text = slide_text(slides[i])
        if sig.lower() not in text.lower():
            errors.append(f"slide {i}: signature {sig!r} not found")
        if layout is not None and slides[i].slide_layout.name != layout:
            errors.append(
                f"slide {i}: layout is {slides[i].slide_layout.name!r}, "
                f"expected {layout!r}"
            )

    if errors:
        print(f"FAIL ({len(errors)} problem(s))")
        for e in errors:
            print("  -", e)
        return 1
    print(f"PASS - {len(slides)} slides in expected order")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 2: Run it to verify it fails**

```bash
cd /Users/laumayp/Development/UG2026 && python3 tools/verify_deck.py
```

Expected: FAIL — a `PackageNotFoundError` traceback, because `deck/UG_London_2026.pptx` does not exist yet.

- [ ] **Step 3: Create the working copy**

```bash
cd /Users/laumayp/Development/UG2026 && mkdir -p deck && cp "resources/Consolidated GenAI common slides.pptx" deck/UG_London_2026.pptx
```

- [ ] **Step 4: Run it to verify it passes**

```bash
cd /Users/laumayp/Development/UG2026 && python3 tools/verify_deck.py
```

Expected: `PASS - 10 slides in expected order`

If a signature fails here, the source deck differs from what the spec recorded — stop and re-read the source rather than weakening the signature.

- [ ] **Step 5: Write the helper module**

Create `tools/deck.py`:

```python
"""Helpers for editing the UG London deck with python-pptx 0.6.23.

python-pptx has no public API for reordering slides or inserting one at a
position, so these reach into the `sldIdLst` element directly.
"""
import re

from pptx import Presentation

DECK = "deck/UG_London_2026.pptx"


def open_deck(path=DECK):
    return Presentation(path)


def find_layout(prs, name):
    """Find a layout by name across ALL masters.

    prs.slide_layouts only exposes the first master's layouts; this deck has
    two masters and 155 layouts, so most names are invisible to that property.
    """
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            if layout.name == name:
                return layout
    raise KeyError(f"no layout named {name!r}")


def describe_layout(prs, name):
    """Print a layout's placeholders so a task can target them by idx."""
    layout = find_layout(prs, name)
    print(f"layout {name!r}:")
    for ph in layout.placeholders:
        pf = ph.placeholder_format
        print(f"  idx={pf.idx}  type={pf.type}  name={ph.name!r}")


def reorder(prs, order):
    """Rebuild slide order. `order` is a list of current indices."""
    lst = prs.slides._sldIdLst
    ids = list(lst)
    if sorted(order) != list(range(len(ids))):
        raise ValueError(f"order must be a permutation of 0..{len(ids) - 1}")
    for el in ids:
        lst.remove(el)
    for i in order:
        lst.append(ids[i])


def add_slide_at(prs, layout_name, index):
    """Append a slide on the named layout, then move it to `index`."""
    slide = prs.slides.add_slide(find_layout(prs, layout_name))
    lst = prs.slides._sldIdLst
    el = list(lst)[-1]
    lst.remove(el)
    lst.insert(index, el)
    return slide


def set_ph(slide, idx, text):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            ph.text_frame.text = text
            return True
    return False


def set_ph_lines(slide, idx, lines):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            tf = ph.text_frame
            tf.text = lines[0]
            for line in lines[1:]:
                tf.add_paragraph().text = line
            return True
    return False


def drop_empty_placeholders(slide):
    """Remove placeholders left empty, so they do not render as prompt text."""
    for ph in list(slide.placeholders):
        if not ph.text_frame.text.strip():
            ph._element.getparent().remove(ph._element)


def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def slide_text(slide):
    return " | ".join(re.findall(r"<a:t>(.*?)</a:t>", slide._element.xml, re.S))
```

- [ ] **Step 6: Verify the helpers import and see both masters**

```bash
cd /Users/laumayp/Development/UG2026 && python3 -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck, find_layout
prs = open_deck()
n = sum(len(m.slide_layouts) for m in prs.slide_masters)
print('masters', len(prs.slide_masters), 'layouts', n)
for name in ['Cover 1', 'Agenda 3', 'Divider 1 - Short title',
             '2 Column, Equal - Subhead', '1 Column - Subhead',
             'Executive Summary/Key Takeaways 2']:
    print(' found:', find_layout(prs, name).name)
"
```

Expected: `masters 2 layouts 155`, then each of the six layout names printed. A `KeyError` means the layout name in the spec is wrong — list the real names before continuing.

- [ ] **Step 7: Confirm .gitignore covers the working files**

`.gitignore` already exists and excludes Office lock files (`~$*`), `build/` and `deck/`. No change needed — just confirm:

```bash
cd /Users/laumayp/Development/UG2026 && git check-ignore -v deck/UG_London_2026.pptx build/x.pdf 'resources/~$test.pptx'
```

Expected: all three report a matching `.gitignore` rule.

The source decks in `resources/` **are** tracked — they are irreplaceable input. `deck/` is not: the working deck is a ~10 MB binary rewritten by all eight authoring tasks, so tracking it would add roughly 80 MB of history for a file fully regenerable from `resources/` plus `tools/`. If the finished deck should be versioned, commit it once at the end rather than un-ignoring the directory.

- [ ] **Step 8: Commit**

```bash
cd /Users/laumayp/Development/UG2026 && git add tools/ && git commit -m "build: deck helpers and structural verification harness"
```

---

### Task 2: Reorder into spine plus appendix

**Files:**
- Modify: `deck/UG_London_2026.pptx`
- Modify: `tools/verify_deck.py` (EXPECTED)

**Interfaces:**
- Consumes: `reorder`, `open_deck` from Task 1.
- Produces: a 10-slide deck ordered C-1, C-5, C-4, C-6, C-8, C-9, then appendix C-2, C-3, C-7, C-10.

- [ ] **Step 1: Update EXPECTED to the target order**

Replace the `EXPECTED` list in `tools/verify_deck.py` with:

```python
EXPECTED = [
    # --- spine ---
    ("Process Automations", None),        # 0  C-1  intro: agentic solutions
    ("AI Solution Suite", None),          # 1  C-5  intro: the three tiers
    ("Every Stage of AI Maturity", None), # 2  C-4  section 2
    ("in three layers", None),            # 3  C-6  section 3
    ("OUR APPROACH", None),               # 4  C-8  section 4
    ("RESPONSIBLE AI", None),             # 5  C-9  section 5b
    # --- appendix ---
    ("Driven by Perspective", None),      # 6  C-2
    ("Content Foundation", None),         # 7  C-3
    ("ACTUARIAL AI CURVE", None),         # 8  C-7
    ("WHY MOODY", None),                  # 9  C-10
]
```

- [ ] **Step 2: Run to verify it fails**

```bash
cd /Users/laumayp/Development/UG2026 && python3 tools/verify_deck.py
```

Expected: FAIL — several `signature ... not found` lines, because the deck is still in source order.

- [ ] **Step 3: Reorder the deck**

```bash
cd /Users/laumayp/Development/UG2026 && python3 -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck, reorder, DECK
prs = open_deck()
# source indices: C-1=0 C-2=1 C-3=2 C-4=3 C-5=4 C-6=5 C-7=6 C-8=7 C-9=8 C-10=9
reorder(prs, [0, 4, 3, 5, 7, 8, 1, 2, 6, 9])
prs.save(DECK)
print('reordered')
"
```

- [ ] **Step 4: Run to verify it passes**

```bash
cd /Users/laumayp/Development/UG2026 && python3 tools/verify_deck.py
```

Expected: `PASS - 10 slides in expected order`

- [ ] **Step 5: Commit**

```bash
cd /Users/laumayp/Development/UG2026 && git add tools/verify_deck.py && git commit -m "feat: reorder deck into spine and appendix"
```

---

### Task 3: Title slide

**Files:**
- Modify: `deck/UG_London_2026.pptx`
- Modify: `tools/verify_deck.py`

**Interfaces:**
- Consumes: `add_slide_at`, `describe_layout`, `set_ph`, `drop_empty_placeholders`, `set_notes`.
- Produces: slide 0 on layout `Cover 1`; every later index shifts by one.

- [ ] **Step 1: Add the expectation**

Insert as the first entry of `EXPECTED` in `tools/verify_deck.py`:

```python
    ("Bringing Trustworthy GenAI to Insurance", "Cover 1"),  # 0  title
```

- [ ] **Step 2: Run to verify it fails**

Expected: FAIL — `slide count is 10, expected 11` plus signature mismatches from the shift.

- [ ] **Step 3: Discover the layout's placeholders**

```bash
cd /Users/laumayp/Development/UG2026 && python3 -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck, describe_layout
describe_layout(open_deck(), 'Cover 1')
"
```

Note the printed `idx` values. The next step assumes title is `idx=0` and the subtitle is the next-lowest idx — if the printout disagrees, use the real numbers.

- [ ] **Step 4: Author the slide**

```bash
cd /Users/laumayp/Development/UG2026 && python3 -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck, add_slide_at, set_ph, drop_empty_placeholders, set_notes, DECK
prs = open_deck()
s = add_slide_at(prs, 'Cover 1', 0)
set_ph(s, 0, 'Bringing Trustworthy GenAI to Insurance')
set_ph(s, 1, 'User Group London 2026')
drop_empty_placeholders(s)
set_notes(s, 'Speaker names to be confirmed. 30 minutes including three live demos. Budget 0:30 here.')
prs.save(DECK)
print('title slide added')
"
```

- [ ] **Step 5: Run to verify it passes**

Expected: `PASS - 11 slides in expected order`

- [ ] **Step 6: Commit**

```bash
cd /Users/laumayp/Development/UG2026 && git add tools/verify_deck.py && git commit -m "feat: add title slide"
```

---

### Task 4: Security landscape slide

**Files:**
- Modify: `deck/UG_London_2026.pptx`
- Modify: `tools/verify_deck.py`

**Interfaces:**
- Consumes: `add_slide_at`, `set_ph_lines`, `drop_empty_placeholders`, `set_notes`.
- Produces: new slide at index 6, immediately before `RESPONSIBLE AI`.

Content is translated and condensed from DEVOXX slides 6 and 9. The English below is the deliverable — do not re-translate from the French.

- [ ] **Step 1: Add the expectation**

Insert into `EXPECTED` between `OUR APPROACH` (index 5) and `RESPONSIBLE AI`:

```python
    ("OWASP", "2 Column, Equal - Subhead"),  # 6  security landscape
```

- [ ] **Step 2: Run to verify it fails**

Expected: FAIL — count 11 vs 12, and `RESPONSIBLE AI` found at the wrong index.

- [ ] **Step 3: Discover the layout's placeholders**

```bash
cd /Users/laumayp/Development/UG2026 && python3 -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck, describe_layout
describe_layout(open_deck(), '2 Column, Equal - Subhead')
"
```

Expect a title, a subhead, and two body placeholders. Map the two bodies to left and right by their `idx` order.

- [ ] **Step 4: Author the slide**

Substitute the real `idx` values from Step 3 for `LEFT` and `RIGHT`:

```bash
cd /Users/laumayp/Development/UG2026 && python3 -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck, add_slide_at, set_ph, set_ph_lines, drop_empty_placeholders, set_notes, DECK
LEFT, RIGHT = 1, 2   # <-- replace with idx values printed in Step 3
prs = open_deck()
s = add_slide_at(prs, '2 Column, Equal - Subhead', 6)
set_ph(s, 0, 'Securing GenAI: the landscape')
set_ph_lines(s, LEFT, [
    'Regulatory pressure',
    'The EU AI Act sets the precedent: a deliberately broad definition of AI, covering both generative and autonomous-action tooling.',
    'Obligations scale with the risk posed to the user.',
    'A worldwide trend, not a European one.',
])
set_ph_lines(s, RIGHT, [
    'OWASP GenAI Security Project',
    'The working checklist for anyone shipping LLM features.',
    'LLM01 Prompt Injection - LLM02 Sensitive Information Disclosure',
    'LLM06 Excessive Agency - LLM07 System Prompt Leakage',
    'Every entry carries a vulnerability, a scenario and a mitigation.',
])
drop_empty_placeholders(s)
set_notes(s, (
    'Budget 1:30. FLEX - if running late, compress to one spoken sentence over the next slide.\n'
    'The exposure is reputational, human and financial.\n'
    'Full OWASP LLM Top 10 for questions: LLM01 Prompt Injection, LLM02 Sensitive Information '
    'Disclosure, LLM03 Supply Chain, LLM04 Data and Model Poisoning, LLM05 Improper Output Handling, '
    'LLM06 Excessive Agency, LLM07 System Prompt Leakage, LLM08 Vector and Embedding Weaknesses, '
    'LLM09 Misinformation, LLM10 Unbounded Consumption.'
))
prs.save(DECK)
print('security slide added')
"
```

- [ ] **Step 5: Run to verify it passes**

Expected: `PASS - 12 slides in expected order`

- [ ] **Step 6: Commit**

```bash
cd /Users/laumayp/Development/UG2026 && git add tools/verify_deck.py && git commit -m "feat: add security landscape slide"
```

---

### Task 5: Use case intro slide

**Files:**
- Modify: `deck/UG_London_2026.pptx`
- Modify: `tools/verify_deck.py`

**Interfaces:**
- Consumes: `add_slide_at`, `set_ph_lines`, `drop_empty_placeholders`, `set_notes`.
- Produces: new slide at index 8, after `RESPONSIBLE AI`.

- [ ] **Step 1: Add the expectation**

Append to `EXPECTED` after `RESPONSIBLE AI` (index 7):

```python
    ("one foundation", "Agenda 3"),  # 8  three use cases
```

- [ ] **Step 2: Run to verify it fails**

Expected: FAIL — count 12 vs 13.

- [ ] **Step 3: Discover the layout's placeholders**

```bash
cd /Users/laumayp/Development/UG2026 && python3 -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck, describe_layout
describe_layout(open_deck(), 'Agenda 3')
"
```

If `Agenda 3` turns out to have fewer than three body placeholders, use `3 Column - Subhead` instead and update the expectation's layout name to match.

- [ ] **Step 4: Author the slide**

```bash
cd /Users/laumayp/Development/UG2026 && python3 -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck, add_slide_at, set_ph, set_ph_lines, drop_empty_placeholders, set_notes, DECK
A, B, C = 1, 2, 3   # <-- replace with idx values printed in Step 3
prs = open_deck()
s = add_slide_at(prs, 'Agenda 3', 8)
set_ph(s, 0, 'Three use cases, one foundation')
set_ph_lines(s, A, ['RiskIntegrity for IFRS 17 Navigator',
                    'Applications layer', 'In production today'])
set_ph_lines(s, B, ['Upgrade Validation Assistant',
                    'Applications layer'])
set_ph_lines(s, C, ['Moody\\u2019s MCP Server',
                    'Open platform layer', 'Your agents, our data'])
drop_empty_placeholders(s)
set_notes(s, (
    'Budget 1:00. CORE.\n'
    'Call back to the three layers slide: each use case proves one layer. '
    'The MCP server was already named on the access-channels slide.'
))
prs.save(DECK)
print('use case intro added')
"
```

- [ ] **Step 5: Run to verify it passes**

Expected: `PASS - 13 slides in expected order`

- [ ] **Step 6: Commit**

```bash
cd /Users/laumayp/Development/UG2026 && git add tools/verify_deck.py && git commit -m "feat: add use case intro slide"
```

---

### Task 6: Use case 1 — context slide and demo card

This task establishes the pattern that Task 7 replicates. Review it before proceeding.

**Files:**
- Modify: `deck/UG_London_2026.pptx`
- Modify: `tools/verify_deck.py`

**Interfaces:**
- Consumes: `add_slide_at`, `set_ph`, `set_ph_lines`, `drop_empty_placeholders`, `set_notes`.
- Produces: slides at indices 9 and 10. Demo scaffolding lives in **speaker notes**, not on the slide face — the audience must never see the fallback plan.

- [ ] **Step 1: Add both expectations**

Append to `EXPECTED` after the use case intro:

```python
    ("RiskIntegrity for IFRS 17 Navigator", "1 Column - Subhead"),  # 9
    ("DEMO", "Divider 1 - Short title"),                            # 10
```

- [ ] **Step 2: Run to verify it fails**

Expected: FAIL — count 13 vs 15.

- [ ] **Step 3: Discover both layouts**

```bash
cd /Users/laumayp/Development/UG2026 && python3 -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck, describe_layout
prs = open_deck()
describe_layout(prs, '1 Column - Subhead')
describe_layout(prs, 'Divider 1 - Short title')
"
```

- [ ] **Step 4: Author both slides**

```bash
cd /Users/laumayp/Development/UG2026 && python3 -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck, add_slide_at, set_ph, set_ph_lines, drop_empty_placeholders, set_notes, DECK
BODY = 1   # <-- replace with the body idx printed for '1 Column - Subhead'
prs = open_deck()

s = add_slide_at(prs, '1 Column - Subhead', 9)
set_ph(s, 0, 'RiskIntegrity for IFRS 17 Navigator')
set_ph_lines(s, BODY, [
    'Today: \\u2039 the workflow as it runs now, and where it hurts \\u2014 one line \\u203a',
    'The assistant: \\u2039 what it does \\u2014 one line \\u203a',
    'Grounded and governed: \\u2039 what it is grounded in, who signs off \\u2014 one line \\u203a',
    'Status: \\u2039 production / pilot / discovery \\u203a',
])
drop_empty_placeholders(s)
set_notes(s, 'Budget 1:00. CORE. Tie the grounded-and-governed line back to the Responsible AI slide.')

d = add_slide_at(prs, 'Divider 1 - Short title', 10)
set_ph(d, 0, 'DEMO')
drop_empty_placeholders(d)
set_notes(d, (
    'RiskIntegrity for IFRS 17 Navigator. Budget 4:00. CORE.\\n'
    'What we will show: \\u2039 3-4 bullets \\u203a\\n'
    'Entry state: \\u2039 what is on screen before I start \\u2014 set up before the talk \\u203a\\n'
    'Exit state: \\u2039 what they should have seen \\u203a\\n'
    'Fallback: \\u2039 screenshot or recording to cut to if it fails \\u203a\\n'
    'Security artifact to surface: a citation \\u2014 shows grounding, not guessing.'
))
prs.save(DECK)
print('use case 1 added')
"
```

- [ ] **Step 5: Run to verify it passes**

Expected: `PASS - 15 slides in expected order`

- [ ] **Step 6: Commit**

```bash
cd /Users/laumayp/Development/UG2026 && git add tools/verify_deck.py && git commit -m "feat: add use case 1 context and demo card"
```

---

### Task 7: Use cases 2 and 3

**Files:**
- Modify: `deck/UG_London_2026.pptx`
- Modify: `tools/verify_deck.py`

**Interfaces:**
- Consumes: same helpers as Task 6, same layouts, same field structure.
- Produces: slides at indices 11–14.

- [ ] **Step 1: Add four expectations**

Append to `EXPECTED` after the use case 1 demo card. All three demo cards carry the same on-slide text, so their `"DEMO"` signatures are not unique — two demo cards swapped would pass this check. The unique context slides on either side pin the ordering, which is sufficient here:

```python
    ("Upgrade Validation Assistant", "1 Column - Subhead"),  # 11
    ("DEMO", "Divider 1 - Short title"),                     # 12
    ("MCP Server", "1 Column - Subhead"),                    # 13
    ("DEMO", "Divider 1 - Short title"),                     # 14
```

- [ ] **Step 2: Run to verify it fails**

Expected: FAIL — count 15 vs 19.

- [ ] **Step 3: Author all four slides**

```bash
cd /Users/laumayp/Development/UG2026 && python3 -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck, add_slide_at, set_ph, set_ph_lines, drop_empty_placeholders, set_notes, DECK
BODY = 1   # <-- same idx used in Task 6
FIELDS = [
    'Today: \\u2039 the workflow as it runs now, and where it hurts \\u2014 one line \\u203a',
    'The assistant: \\u2039 what it does \\u2014 one line \\u203a',
    'Grounded and governed: \\u2039 what it is grounded in, who signs off \\u2014 one line \\u203a',
    'Status: \\u2039 production / pilot / discovery \\u203a',
]
DEMO_FIELDS = (
    'What we will show: \\u2039 3-4 bullets \\u203a\\n'
    'Entry state: \\u2039 what is on screen before I start \\u2014 set up before the talk \\u203a\\n'
    'Exit state: \\u2039 what they should have seen \\u203a\\n'
    'Fallback: \\u2039 screenshot or recording to cut to if it fails \\u203a\\n'
)
prs = open_deck()

s = add_slide_at(prs, '1 Column - Subhead', 11)
set_ph(s, 0, 'Upgrade Validation Assistant')
set_ph_lines(s, BODY, FIELDS)
drop_empty_placeholders(s)
set_notes(s, 'Budget 1:00. CORE.')

d = add_slide_at(prs, 'Divider 1 - Short title', 12)
set_ph(d, 0, 'DEMO')
drop_empty_placeholders(d)
set_notes(d, 'Upgrade Validation Assistant. Budget 4:00. CORE.\\n' + DEMO_FIELDS +
             'Security artifact to surface: a guardrail refusing an out-of-scope request.')

s = add_slide_at(prs, '1 Column - Subhead', 13)
set_ph(s, 0, 'Moody\\u2019s MCP Server')
set_ph_lines(s, BODY, FIELDS)
drop_empty_placeholders(s)
set_notes(s, 'Budget 1:00. FLEX - the demo can carry this section alone if time is short.')

d = add_slide_at(prs, 'Divider 1 - Short title', 14)
set_ph(d, 0, 'DEMO')
drop_empty_placeholders(d)
set_notes(d, 'Moody\\u2019s MCP Server. Budget 3:00. CORE.\\n' + DEMO_FIELDS +
             'Security artifact to surface: tool permissions scoped per user and tenant.')

prs.save(DECK)
print('use cases 2 and 3 added')
"
```

- [ ] **Step 4: Run to verify it passes**

Expected: `PASS - 19 slides in expected order`

- [ ] **Step 5: Commit**

```bash
cd /Users/laumayp/Development/UG2026 && git add tools/verify_deck.py && git commit -m "feat: add use cases 2 and 3"
```

---

### Task 8: Conclusion slide

**Files:**
- Modify: `deck/UG_London_2026.pptx`
- Modify: `tools/verify_deck.py`

**Interfaces:**
- Consumes: `add_slide_at`, `set_ph_lines`, `drop_empty_placeholders`, `set_notes`.
- Produces: slide at index 15, the last of the main line. Appendix follows at 16–19.

- [ ] **Step 1: Add the expectation**

Append after the use case 3 demo card:

```python
    ("Trustworthy, in production", "Executive Summary/Key Takeaways 2"),  # 15
```

- [ ] **Step 2: Run to verify it fails**

Expected: FAIL — count 19 vs 20.

- [ ] **Step 3: Discover the layout's placeholders**

```bash
cd /Users/laumayp/Development/UG2026 && python3 -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck, describe_layout
describe_layout(open_deck(), 'Executive Summary/Key Takeaways 2')
"
```

- [ ] **Step 4: Author the slide**

```bash
cd /Users/laumayp/Development/UG2026 && python3 -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck, add_slide_at, set_ph, set_ph_lines, drop_empty_placeholders, set_notes, DECK
BODY = 1   # <-- replace with the body idx printed in Step 3
prs = open_deck()
s = add_slide_at(prs, 'Executive Summary/Key Takeaways 2', 15)
set_ph(s, 0, 'Trustworthy, in production')
set_ph_lines(s, BODY, [
    'Grounded, not guessing \\u2014 answers cited to AXIS documentation and your own dataset.',
    'Governed access \\u2014 security is platform engineering, not prompt engineering.',
    'Human sign-off \\u2014 actuaries stay in control of anything touching a reported number.',
    'Transparent and auditable \\u2014 cited outputs and logged actions your regulator can follow.',
    'Layer three, the open platform, is a conversation we want to start. With this room.',
])
drop_empty_placeholders(s)
set_notes(s, (
    'Budget 1:30. CORE.\\n'
    'Close on the invitation, not the summary. This audience is exactly who layer three is for.\\n'
    'Appendix follows: why Moody\\u2019s credentials, the data foundation, and the actuarial AI curve.'
))
prs.save(DECK)
print('conclusion added')
"
```

- [ ] **Step 5: Run to verify it passes**

Expected: `PASS - 20 slides in expected order`

- [ ] **Step 6: Commit**

```bash
cd /Users/laumayp/Development/UG2026 && git add tools/verify_deck.py && git commit -m "feat: add conclusion slide"
```

---

### Task 9: Speaker notes on reused slides

The six reused slides arrived without timing or cut marks. Every slide must carry its budget so the deck is rehearsable.

**Files:**
- Modify: `deck/UG_London_2026.pptx`
- Create: `tools/verify_notes.py`

**Interfaces:**
- Consumes: `open_deck`, `set_notes`.
- Produces: `tools/verify_notes.py`, runnable as `python3 tools/verify_notes.py`, asserting every main-line slide (0–15) has a `Budget` line in its notes.

- [ ] **Step 1: Write the notes verifier**

Create `tools/verify_notes.py`:

```python
#!/usr/bin/env python3
"""Assert every main-line slide carries a timing budget in its speaker notes."""
import sys

from pptx import Presentation

DECK = "deck/UG_London_2026.pptx"
MAIN_LINE = 16


def main():
    prs = Presentation(DECK)
    errors = []
    for i, slide in enumerate(list(prs.slides)[:MAIN_LINE]):
        notes = (
            slide.notes_slide.notes_text_frame.text
            if slide.has_notes_slide
            else ""
        )
        if "Budget" not in notes:
            errors.append(f"slide {i}: no Budget line in speaker notes")
        elif not any(m in notes for m in ("CORE", "FLEX")):
            errors.append(f"slide {i}: notes have no CORE/FLEX mark")
    if errors:
        print(f"FAIL ({len(errors)} problem(s))")
        for e in errors:
            print("  -", e)
        return 1
    print(f"PASS - all {MAIN_LINE} main-line slides carry budget and cut marks")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 2: Run to verify it fails**

```bash
cd /Users/laumayp/Development/UG2026 && python3 tools/verify_notes.py
```

Expected: FAIL — the six reused slides (1, 2, 3, 4, 5, 7) have no notes, and the title slide's note has no CORE/FLEX mark.

- [ ] **Step 3: Add notes to the reused slides and the title**

```bash
cd /Users/laumayp/Development/UG2026 && python3 -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck, set_notes, DECK
NOTES = {
  0: 'Budget 0:30. CORE. Speaker names to be confirmed. 30 minutes including three live demos.',
  1: 'Budget 1:30. FLEX - can be cut to a 20-second verbal framing.\\n'
     'Open Moody\\u2019s-wide for credibility: four agentic use cases already running.',
  2: 'Budget 1:30. FLEX - overlaps the previous slide and the next; first to cut.\\n'
     'The three tiers: assistants, agentic solutions, AI-ready data.',
  3: 'Budget 1:30. CORE. Note MCP Servers among the channels - this plants use case 3.',
  4: 'Budget 2:00. CORE. This is the pivot from Moody\\u2019s-wide to Life and AXIS.\\n'
     'Layer three is flagged here and cashed in at the conclusion.',
  5: 'Budget 1:30. CORE. Frontier tools, reusable blocks, built with clients.',
  7: 'Budget 1:30. CORE. The four principles. This is the trust spine the use cases hang from.',
}
prs = open_deck()
slides = list(prs.slides)
for i, text in NOTES.items():
    set_notes(slides[i], text)
prs.save(DECK)
print('notes written to', len(NOTES), 'slides')
"
```

- [ ] **Step 4: Run both verifiers to confirm they pass**

```bash
cd /Users/laumayp/Development/UG2026 && python3 tools/verify_deck.py && python3 tools/verify_notes.py
```

Expected: `PASS - 20 slides in expected order` then `PASS - all 16 main-line slides carry budget and cut marks`

- [ ] **Step 5: Commit**

```bash
cd /Users/laumayp/Development/UG2026 && git add tools/verify_notes.py && git commit -m "feat: add timing budgets and cut marks to speaker notes"
```

---

### Task 10: Render check and README correction

Structural verification cannot see overflowing text or a placeholder that renders as prompt text. This task looks at the deck.

**Files:**
- Create: `build/UG_London_2026.pdf` (generated, git-ignored)
- Modify: `README.md`

**Interfaces:**
- Consumes: the finished deck.
- Produces: a PDF render for visual inspection; a README abstract matching the running order.

- [ ] **Step 1: Render to PDF**

```bash
cd /Users/laumayp/Development/UG2026 && mkdir -p build && soffice --headless --convert-to pdf --outdir build deck/UG_London_2026.pptx && ls -la build/
```

Expected: `build/UG_London_2026.pdf` exists. LibreOffice renders the Moody's template imperfectly — judge layout and overflow, not typography.

- [ ] **Step 2: Inspect every new slide**

Read `build/UG_London_2026.pdf` pages 1, 7, 9, 10, 11, 12, 13, 14, 15, 16 (the ten authored slides). For each, confirm: the title is present and unclipped, body text does not overflow its box, and no empty placeholder is rendering as prompt text ("Click to edit…").

Record anything wrong and fix it in `deck/UG_London_2026.pptx` before continuing. Re-render and re-inspect after each fix.

- [ ] **Step 3: Correct the README abstract**

The published abstract says the talk *turns to* security last. The agreed running order puts security at section 5, before the use cases. Update the abstract in `README.md` so the description matches the talk. Change only the ordering claim — leave the rest of the sentence intact:

Replace `then turn to what matters most for regulated teams: securing AI in production, with concrete examples over technical detail.` with `and set out what matters most for regulated teams — securing AI in production, with concrete examples over technical detail — before showing the work.`

- [ ] **Step 4: Run both verifiers one last time**

```bash
cd /Users/laumayp/Development/UG2026 && python3 tools/verify_deck.py && python3 tools/verify_notes.py
```

Expected: both PASS.

- [ ] **Step 5: Commit**

```bash
cd /Users/laumayp/Development/UG2026 && git add README.md && git commit -m "docs: align README abstract with the talk running order"
```

---

## Handoff notes

Left deliberately unfilled, for the author rather than the implementer:

1. **Speaker names** on slide 0 — not recorded anywhere in the repo.
2. **All `‹ … ›` fields** on the six use case slides and their demo notes — none of these products are described in this repo, and inventing their capabilities would be worse than leaving the prompts visible.
3. **The finished deck is not versioned.** `resources/` is tracked, but `deck/` is ignored because the working file is rewritten by every authoring task. Once the deck is final, commit it deliberately with `git add -f deck/UG_London_2026.pptx` — one commit of one binary, rather than eight.
4. **Use case 1's name** — built as "RiskIntegrity for IFRS 17 Navigator" per the README. If "Infoweb Navigator" is the real product, change it in `EXPECTED` and on slides 9 and 10.
