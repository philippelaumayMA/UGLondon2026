# UG London 2026 Deck Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.
>
> **Also required:** the `shared-life-beta-ai-skills:implementation-moodys-pptx` skill. Every authoring step calls its `moodys_layouts` builders.

**Goal:** Build a 22-slide English `.pptx` for a 30-minute talk — 16 main line, 4 appendix, back cover, disclaimer — reusing six slides from the existing Moody's deck and authoring ten new ones with the skill's TOOLKIT builders.

**Architecture:** Copy `resources/Consolidated GenAI common slides.pptx` into `deck/`, renumber its slide IDs, reorder the existing slides into the spine, then author new slides with `moodys_layouts` opened directly on that deck — no `reset_slides`. The Consolidated deck is verified to be the skill's own official template, so builders resolve their layouts by name against it. `tools/verify_deck.py` holds the expected slide order as data and grows one entry per task; that is the red/green cycle.

**Tech Stack:** Python 3.11 via `uv run --with python-pptx` (add `--with cairosvg` for icons), python-pptx 0.6.23, the `implementation-moodys-pptx` skill, LibreOffice (`soffice`) for render checks.

**Spec:** `docs/superpowers/specs/2026-08-31-ug-london-deck-design.md`

## Global Constraints

- **Slide IDs must be renumbered before any slide is added** (Task 1). The source deck's IDs top out at `2147483646`; python-pptx assigns `max + 1`, so it accepts exactly one new slide and then raises `ValueError: value must be in range 256 to 2147483647`. This is not optional and not recoverable later.
- **Never call `ml.md.reset_slides`.** It would delete the six reused slides. This deck is edited in place.
- **Shell preamble for every authoring command:** `source tools/env.sh`, which exports
  `SKILL`, `MOODYS_ASSETS_DIR` and `DYLD_FALLBACK_LIBRARY_PATH`.
  `MOODYS_ASSETS_DIR` is required because `CLAUDE_PLUGIN_DATA` is unset in a plain shell; without it every icon silently becomes a blue circle. `DYLD_FALLBACK_LIBRARY_PATH` is required on macOS: cairocffi dlopens `libcairo` by bare name and does not search Homebrew's prefix, so `import cairosvg` raises `OSError: no library called "cairo-2" was found` and the authoring command dies mid-slide. Needs `brew install cairo` (present on this machine, 1.18.4).
- **`ml.set_header_style("h1h2")` once** before authoring; pass every content builder a `subhead`.
- **Icons on every item-based builder.** Resolve each name with `ml.md.find_icons("keyword")` **before** using it — `shield` and `check` return nothing. Never repeat an icon within a slide.
- **English only.** The DEVOXX source is French; translate, never paste.
- **Product name is `RiskIntegrity for IFRS 17 Navigator`** — the README's name, not "Infoweb Navigator".
- **Never invent product facts.** Use case slides carry bracketed prompts in guillemets (`‹ … ›`). Nothing in this repo describes what Navigator, UVA or the MCP server do.
- **Audience-facing vs speaker-facing.** Fallback plans, timings and cut marks go in speaker notes, never on the slide face.
- **`ml.disclaimer(prs)` is mandatory and always the last slide.**
- **Final state:** 22 slides — main line 0–15, appendix 16–19, back cover 20, disclaimer 21.

---

### Task 1: Working copy, ID renumber, and verification harness

**Files:**
- Create: `tools/deck.py`, `tools/verify_deck.py`, `.gitignore` (already committed — confirm only)
- Create: `deck/UG_London_2026.pptx` (copied binary, git-ignored)

**Interfaces:**
- Produces: `tools/deck.py` exposing `open_deck(path=DECK) -> Presentation`, `renumber_slide_ids(prs) -> list[int]`, `reorder(prs, order: list[int]) -> None`, `move_slide(prs, old: int, new: int) -> None`, `set_notes(slide, text: str) -> None`, `slide_text(slide) -> str`, and the constant `DECK`. `tools/verify_deck.py` exposes `EXPECTED: list[tuple[str, str | None]]`.

- [ ] **Step 1: Write the verification script**

Create `tools/verify_deck.py`:

```python
#!/usr/bin/env python3
"""Assert deck/UG_London_2026.pptx matches the spec's slide map.

EXPECTED is the source of truth: one (signature, layout_name) pair per slide,
in order. `signature` is a distinctive substring that must appear somewhere on
the slide, matched case-insensitively; None skips the text check, which the
disclaimer needs -- that slide has zero shapes of its own and inherits all its
legal copy from its layout. `layout_name` is checked only when not None --
reused slides keep whatever layout they arrived with.

Signatures deliberately avoid apostrophes: the deck uses curly quotes (U+2019)
throughout, so "Moody's" typed with a straight quote would never match.
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

    Reads raw XML rather than walking shapes: python-pptx does not descend into
    grouped shapes, and several source slides are heavily grouped.
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
        if sig is not None and sig.lower() not in text.lower():
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
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python tools/verify_deck.py
```

Expected: FAIL — a `PackageNotFoundError` traceback, because `deck/UG_London_2026.pptx` does not exist yet.

- [ ] **Step 3: Create the working copy**

```bash
cd /Users/laumayp/Development/UG2026 && mkdir -p deck && cp "resources/Consolidated GenAI common slides.pptx" deck/UG_London_2026.pptx
```

- [ ] **Step 4: Run it to verify it passes**

Expected: `PASS - 10 slides in expected order`

A signature failure here means the source deck differs from what the spec recorded — re-read the source rather than weakening the signature.

- [ ] **Step 5: Write the helper module**

Create `tools/deck.py`:

```python
"""Helpers for editing the UG London deck with python-pptx 0.6.23.

python-pptx has no public API for reordering slides or moving one to a
position, so these reach into the `sldIdLst` element directly.
"""
import re

from pptx import Presentation

DECK = "deck/UG_London_2026.pptx"


def open_deck(path=DECK):
    return Presentation(path)


def renumber_slide_ids(prs, start=256):
    """Renumber every p:sldId id into a low contiguous range.

    The source deck ships ids just under the int32 ceiling (max 2147483646).
    python-pptx assigns max(ids)+1 per added slide, so without this the deck
    accepts exactly ONE new slide and then raises ValueError. Slide ids are
    arbitrary internal identifiers, independent of the r:id relationships, so
    renumbering is lossless. Returns the new ids.
    """
    for n, el in enumerate(prs.slides._sldIdLst, start=start):
        el.set("id", str(n))
    return [int(el.get("id")) for el in prs.slides._sldIdLst]


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


def move_slide(prs, old, new):
    """Move the slide at `old` to index `new`.

    Builders append to the end; this puts the result where the spec wants it.
    """
    lst = prs.slides._sldIdLst
    el = list(lst)[old]
    lst.remove(el)
    lst.insert(new, el)


def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def slide_text(slide):
    return " | ".join(re.findall(r"<a:t>(.*?)</a:t>", slide._element.xml, re.S))
```

- [ ] **Step 6: Renumber the slide IDs**

```bash
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck, renumber_slide_ids, DECK
prs = open_deck()
before = [int(e.get('id')) for e in prs.slides._sldIdLst]
after = renumber_slide_ids(prs)
prs.save(DECK)
print('max id', max(before), '->', max(after))
"
```

Expected: `max id 2147483646 -> 265`

- [ ] **Step 7: Prove the deck now accepts more than one new slide**

This is the whole point of Step 6, so verify it rather than assuming:

```bash
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck
prs = open_deck()
lay = prs.slide_masters[0].slide_layouts[0]
for i in range(3):
    prs.slides.add_slide(lay)
print('added 3 scratch slides OK - not saved')
"
```

Expected: `added 3 scratch slides OK - not saved`. The file is deliberately not saved, so the scratch slides are discarded. Before Step 6 this raised `ValueError` on the second slide.

- [ ] **Step 8: Confirm .gitignore covers the working files**

```bash
cd /Users/laumayp/Development/UG2026 && git check-ignore -v deck/UG_London_2026.pptx build/x.pdf 'resources/~$t.pptx'
```

Expected: all three report a matching `.gitignore` rule. `resources/` is tracked (irreplaceable input); `deck/` is not, because the working deck is a ~10 MB binary rewritten by every authoring task.

- [ ] **Step 9: Run the verifier, then commit**

```bash
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python tools/verify_deck.py && git add tools/ && git commit -m "build: deck helpers, slide-id renumber, verification harness"
```

Expected: `PASS - 10 slides in expected order`, then a commit.

---

### Task 2: Reorder into spine plus appendix

**Files:** Modify `deck/UG_London_2026.pptx`, `tools/verify_deck.py`

**Interfaces:** Consumes `reorder`, `open_deck`. Produces a 10-slide deck ordered C-1, C-5, C-4, C-6, C-8, C-9, then appendix C-2, C-3, C-7, C-10.

- [ ] **Step 1: Update EXPECTED to the target order**

Replace `EXPECTED` in `tools/verify_deck.py`:

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

Expected: FAIL — several `signature ... not found`, because the deck is still in source order.

- [ ] **Step 3: Reorder**

```bash
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck, reorder, DECK
prs = open_deck()
# source indices: C-1=0 C-2=1 C-3=2 C-4=3 C-5=4 C-6=5 C-7=6 C-8=7 C-9=8 C-10=9
reorder(prs, [0, 4, 3, 5, 7, 8, 1, 2, 6, 9])
prs.save(DECK)
print('reordered')
"
```

- [ ] **Step 4: Run to verify it passes, then commit**

```bash
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python tools/verify_deck.py && git add tools/verify_deck.py && git commit -m "feat: reorder deck into spine and appendix"
```

Expected: `PASS - 10 slides in expected order`

---

### Task 3: Resolve the icon names

Every later task needs icon names that actually exist. Resolve them once, here, rather than guessing mid-build. `shield` and `check` return nothing, so guessing produces silent blue circles.

**Files:** Create `tools/icons.py`

**Interfaces:** Produces `tools/icons.py` exposing `ICONS: dict[str, str]`, mapping a role key to a verified official icon name.

- [ ] **Step 1: Search for a candidate per role**

```bash
export SKILL=/Users/laumayp/.claude/plugins/cache/ibu-life-marketplace/shared-life-beta-ai-skills/0.8.0/skills/implementation-moodys-pptx
export MOODYS_ASSETS_DIR=/Users/laumayp/.claude/plugins/data/shared-life-beta-ai-skills-ibu-life-marketplace/moodys-assets
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx --with cairosvg python -c "
import sys, os; sys.path.insert(0, os.environ['SKILL'] + '/scripts')
import moodys_deck as md
md.ensure_icons()
for kw in ['regulation','legal','balance','warning','risk','lock','document',
           'target','cloud','data','people','process','settings','search',
           'approved','audit','network','robot']:
    print(kw.ljust(12), (md.find_icons(kw) or [])[:4])
"
```

Read the output and choose one distinct icon per role below. Do not reuse an icon across a single slide.

- [ ] **Step 2: Write the resolved map**

Create `tools/icons.py`, substituting names that appeared in Step 1's output:

```python
"""Verified official Moody's icon names, one per role.

Every value MUST have appeared in `md.find_icons()` output -- an unresolved
name degrades silently to a bright-blue circle placeholder.
"""
ICONS = {
    # slide 6 - security landscape
    "regulation": "‹ from Step 1 ›",
    "owasp": "‹ from Step 1 ›",
    # slide 8 - three use cases
    "navigator": "‹ from Step 1 ›",
    "uva": "‹ from Step 1 ›",
    "mcp": "‹ from Step 1 ›",
    # slides 9/11/13 - use case fields
    "today": "‹ from Step 1 ›",
    "assistant": "‹ from Step 1 ›",
    "governed": "‹ from Step 1 ›",
    "status": "‹ from Step 1 ›",
    # slide 15 - conclusion principles
    "grounded": "‹ from Step 1 ›",
    "access": "‹ from Step 1 ›",
    "human": "‹ from Step 1 ›",
    "audit": "‹ from Step 1 ›",
    "invite": "‹ from Step 1 ›",
}
```

- [ ] **Step 3: Verify every name resolves**

```bash
export SKILL=/Users/laumayp/.claude/plugins/cache/ibu-life-marketplace/shared-life-beta-ai-skills/0.8.0/skills/implementation-moodys-pptx
export MOODYS_ASSETS_DIR=/Users/laumayp/.claude/plugins/data/shared-life-beta-ai-skills-ibu-life-marketplace/moodys-assets
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx --with cairosvg python -c "
import sys, os; sys.path.insert(0, os.environ['SKILL'] + '/scripts'); sys.path.insert(0, 'tools')
import moodys_deck as md
from icons import ICONS
md.ensure_icons()
bad = [k for k, v in ICONS.items() if md.find_icon(v) is None]
print('unresolved:', bad or 'none')
sys.exit(1 if bad else 0)
"
```

Expected: `unresolved: none` and exit 0. Any listed key has a wrong name — fix it before continuing, or that slide renders blue circles.

- [ ] **Step 4: Commit**

```bash
cd /Users/laumayp/Development/UG2026 && git add tools/icons.py && git commit -m "feat: resolve official icon names for the deck"
```

---

### Task 4: Title slide

**Files:** Modify `deck/UG_London_2026.pptx`, `tools/verify_deck.py`

**Interfaces:** Consumes `ml.cover`, `move_slide`, `set_notes`. Produces slide 0 on layout `Cover 1`; every later index shifts by one.

- [ ] **Step 1: Add the expectation**

Insert as the first entry of `EXPECTED`:

```python
    ("Bringing Trustworthy GenAI to Insurance", "Cover 1"),  # 0  title
```

- [ ] **Step 2: Run to verify it fails**

Expected: FAIL — `slide count is 10, expected 11` plus signature shifts.

- [ ] **Step 3: Author the slide**

Builders append, so the cover is moved to index 0 afterwards.

```bash
export SKILL=/Users/laumayp/.claude/plugins/cache/ibu-life-marketplace/shared-life-beta-ai-skills/0.8.0/skills/implementation-moodys-pptx
export MOODYS_ASSETS_DIR=/Users/laumayp/.claude/plugins/data/shared-life-beta-ai-skills-ibu-life-marketplace/moodys-assets
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx --with cairosvg python -c "
import sys, os; sys.path.insert(0, os.environ['SKILL'] + '/scripts'); sys.path.insert(0, 'tools')
import moodys_layouts as ml
from deck import open_deck, move_slide, set_notes, DECK
prs = open_deck()
ml.set_header_style('h1h2')
s = ml.cover(prs, 'Bringing Trustworthy GenAI to Insurance',
             subtitle='User Group London 2026')
set_notes(s, 'Budget 0:30. CORE. Speaker names to be confirmed. 30 minutes including three live demos.')
move_slide(prs, len(prs.slides._sldIdLst) - 1, 0)
prs.save(DECK)
print('title slide added at index 0')
"
```

- [ ] **Step 4: Run to verify it passes, then commit**

```bash
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python tools/verify_deck.py && git add tools/verify_deck.py && git commit -m "feat: add title slide"
```

Expected: `PASS - 11 slides in expected order`

---

### Task 5: Security landscape slide

**Files:** Modify `deck/UG_London_2026.pptx`, `tools/verify_deck.py`

**Interfaces:** Consumes `ml.columns`, `move_slide`, `set_notes`, `ICONS`. Produces a new slide at index 6, immediately before `RESPONSIBLE AI`.

The English below is the deliverable, translated and condensed from DEVOXX slides 6 and 9. Do not re-translate from the French.

- [ ] **Step 1: Add the expectation**

Insert between `OUR APPROACH` (index 5) and `RESPONSIBLE AI`. `ml.columns` composes on a `Title Only` canvas:

```python
    ("OWASP", "Title Only"),  # 6  security landscape
```

- [ ] **Step 2: Run to verify it fails**

Expected: FAIL — count 11 vs 12, and `RESPONSIBLE AI` at the wrong index.

- [ ] **Step 3: Author the slide**

```bash
export SKILL=/Users/laumayp/.claude/plugins/cache/ibu-life-marketplace/shared-life-beta-ai-skills/0.8.0/skills/implementation-moodys-pptx
export MOODYS_ASSETS_DIR=/Users/laumayp/.claude/plugins/data/shared-life-beta-ai-skills-ibu-life-marketplace/moodys-assets
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx --with cairosvg python -c "
import sys, os; sys.path.insert(0, os.environ['SKILL'] + '/scripts'); sys.path.insert(0, 'tools')
import moodys_layouts as ml
from deck import open_deck, move_slide, set_notes, DECK
from icons import ICONS
prs = open_deck()
ml.set_header_style('h1h2')
s = ml.columns(prs, 'Securing GenAI: the landscape', [
    {'kicker': 'REGULATORY PRESSURE',
     'head': 'The EU AI Act sets the precedent',
     'body': 'A deliberately broad definition of AI, covering both generative and autonomous-action tooling. Obligations scale with the risk posed to the user. A worldwide trend, not a European one.',
     'icon': ml.md.find_icon(ICONS['regulation'])},
    {'kicker': 'OWASP GENAI SECURITY PROJECT',
     'head': 'The working checklist',
     'body': 'LLM01 Prompt Injection. LLM02 Sensitive Information Disclosure. LLM06 Excessive Agency. LLM07 System Prompt Leakage. Every entry carries a vulnerability, a scenario and a mitigation.',
     'icon': ml.md.find_icon(ICONS['owasp'])},
], subhead='Two forces shape what production-ready means')
set_notes(s, (
    'Budget 1:30. FLEX - if running late, compress to one spoken sentence over the next slide.\n'
    'The exposure is reputational, human and financial.\n'
    'Full OWASP LLM Top 10 for questions: LLM01 Prompt Injection, LLM02 Sensitive Information '
    'Disclosure, LLM03 Supply Chain, LLM04 Data and Model Poisoning, LLM05 Improper Output '
    'Handling, LLM06 Excessive Agency, LLM07 System Prompt Leakage, LLM08 Vector and Embedding '
    'Weaknesses, LLM09 Misinformation, LLM10 Unbounded Consumption.'
))
move_slide(prs, len(prs.slides._sldIdLst) - 1, 6)
prs.save(DECK)
print('security slide added at index 6')
"
```

- [ ] **Step 4: Run to verify it passes, then commit**

```bash
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python tools/verify_deck.py && git add tools/verify_deck.py && git commit -m "feat: add security landscape slide"
```

Expected: `PASS - 12 slides in expected order`

---

### Task 6: Use case intro slide

**Files:** Modify `deck/UG_London_2026.pptx`, `tools/verify_deck.py`

**Interfaces:** Consumes `ml.columns`, `move_slide`, `set_notes`, `ICONS`. Produces a new slide at index 8, after `RESPONSIBLE AI`.

- [ ] **Step 1: Add the expectation**

```python
    ("one foundation", "Title Only"),  # 8  three use cases
```

- [ ] **Step 2: Run to verify it fails**

Expected: FAIL — count 12 vs 13.

- [ ] **Step 3: Author the slide**

```bash
export SKILL=/Users/laumayp/.claude/plugins/cache/ibu-life-marketplace/shared-life-beta-ai-skills/0.8.0/skills/implementation-moodys-pptx
export MOODYS_ASSETS_DIR=/Users/laumayp/.claude/plugins/data/shared-life-beta-ai-skills-ibu-life-marketplace/moodys-assets
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx --with cairosvg python -c "
import sys, os; sys.path.insert(0, os.environ['SKILL'] + '/scripts'); sys.path.insert(0, 'tools')
import moodys_layouts as ml
from deck import open_deck, move_slide, set_notes, DECK
from icons import ICONS
prs = open_deck()
ml.set_header_style('h1h2')
s = ml.columns(prs, 'Three use cases, one foundation', [
    {'kicker': 'APPLICATIONS LAYER', 'head': 'RiskIntegrity for IFRS 17 Navigator',
     'body': 'In production today.', 'icon': ml.md.find_icon(ICONS['navigator'])},
    {'kicker': 'APPLICATIONS LAYER', 'head': 'Upgrade Validation Assistant',
     'body': 'Built with clients, on the same foundation.', 'icon': ml.md.find_icon(ICONS['uva'])},
    {'kicker': 'OPEN PLATFORM LAYER', 'head': 'Moody’s MCP Server',
     'body': 'Your agents, our data.', 'icon': ml.md.find_icon(ICONS['mcp'])},
], subhead='Each one proves a layer of the strategy')
set_notes(s, (
    'Budget 1:00. CORE.\n'
    'Call back to the three layers slide: each use case proves one layer. '
    'The MCP server was already named on the access-channels slide.'
))
move_slide(prs, len(prs.slides._sldIdLst) - 1, 8)
prs.save(DECK)
print('use case intro added at index 8')
"
```

- [ ] **Step 4: Run to verify it passes, then commit**

```bash
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python tools/verify_deck.py && git add tools/verify_deck.py && git commit -m "feat: add use case intro slide"
```

Expected: `PASS - 13 slides in expected order`

---

### Task 7: Use case 1 — context slide and demo card

This establishes the pattern Task 8 replicates. Review before proceeding.

**Files:** Modify `deck/UG_London_2026.pptx`, `tools/verify_deck.py`

**Interfaces:** Consumes `ml.icon_text`, `ml.divider`, `move_slide`, `set_notes`, `ICONS`. Produces slides at indices 9 and 10. Demo scaffolding lives in **speaker notes** — the audience must never see the fallback plan.

- [ ] **Step 1: Add both expectations**

```python
    ("RiskIntegrity for IFRS 17 Navigator", "Title Only"),   # 9
    ("DEMO", "Divider 3 - Short title"),                     # 10
```

- [ ] **Step 2: Run to verify it fails**

Expected: FAIL — count 13 vs 15.

- [ ] **Step 3: Author both slides**

```bash
export SKILL=/Users/laumayp/.claude/plugins/cache/ibu-life-marketplace/shared-life-beta-ai-skills/0.8.0/skills/implementation-moodys-pptx
export MOODYS_ASSETS_DIR=/Users/laumayp/.claude/plugins/data/shared-life-beta-ai-skills-ibu-life-marketplace/moodys-assets
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx --with cairosvg python -c "
import sys, os; sys.path.insert(0, os.environ['SKILL'] + '/scripts'); sys.path.insert(0, 'tools')
import moodys_layouts as ml
from deck import open_deck, move_slide, set_notes, DECK
from icons import ICONS
prs = open_deck()
ml.set_header_style('h1h2')
FIELDS = [
    {'head': 'Today', 'body': '‹ the workflow as it runs now, and where it hurts — one line ›',
     'icon': ml.md.find_icon(ICONS['today'])},
    {'head': 'The assistant', 'body': '‹ what it does — one line ›',
     'icon': ml.md.find_icon(ICONS['assistant'])},
    {'head': 'Grounded and governed', 'body': '‹ what it is grounded in, who signs off — one line ›',
     'icon': ml.md.find_icon(ICONS['governed'])},
    {'head': 'Status', 'body': '‹ production / pilot / discovery ›',
     'icon': ml.md.find_icon(ICONS['status'])},
]
s = ml.icon_text(prs, 'RiskIntegrity for IFRS 17 Navigator', FIELDS, subhead='Use case 1')
set_notes(s, 'Budget 1:00. CORE. Tie the grounded-and-governed line back to the Responsible AI slide.')
move_slide(prs, len(prs.slides._sldIdLst) - 1, 9)

d = ml.divider(prs, 'DEMO', eyebrow='Use case 1')
set_notes(d, (
    'RiskIntegrity for IFRS 17 Navigator. Budget 4:00. CORE.\n'
    'What we will show: ‹ 3-4 bullets ›\n'
    'Entry state: ‹ what is on screen before I start — set up before the talk ›\n'
    'Exit state: ‹ what they should have seen ›\n'
    'Fallback: ‹ screenshot or recording to cut to if it fails ›\n'
    'Security artifact to surface: a citation — shows grounding, not guessing.'
))
move_slide(prs, len(prs.slides._sldIdLst) - 1, 10)
prs.save(DECK)
print('use case 1 added at indices 9 and 10')
"
```

- [ ] **Step 4: Run to verify it passes, then commit**

```bash
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python tools/verify_deck.py && git add tools/verify_deck.py && git commit -m "feat: add use case 1 context and demo card"
```

Expected: `PASS - 15 slides in expected order`

---

### Task 8: Use cases 2 and 3

**Files:** Modify `deck/UG_London_2026.pptx`, `tools/verify_deck.py`

**Interfaces:** Same builders, layouts and field structure as Task 7. Produces slides at indices 11–14.

- [ ] **Step 1: Add four expectations**

All three demo cards carry the same on-slide text, so their `"DEMO"` signatures are not unique — two swapped demo cards would pass. The unique context slides on either side pin the ordering, which is sufficient here.

```python
    ("Upgrade Validation Assistant", "Title Only"),   # 11
    ("DEMO", "Divider 3 - Short title"),              # 12
    ("MCP Server", "Title Only"),                     # 13
    ("DEMO", "Divider 3 - Short title"),              # 14
```

- [ ] **Step 2: Run to verify it fails**

Expected: FAIL — count 15 vs 19.

- [ ] **Step 3: Author all four slides**

```bash
export SKILL=/Users/laumayp/.claude/plugins/cache/ibu-life-marketplace/shared-life-beta-ai-skills/0.8.0/skills/implementation-moodys-pptx
export MOODYS_ASSETS_DIR=/Users/laumayp/.claude/plugins/data/shared-life-beta-ai-skills-ibu-life-marketplace/moodys-assets
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx --with cairosvg python -c "
import sys, os; sys.path.insert(0, os.environ['SKILL'] + '/scripts'); sys.path.insert(0, 'tools')
import moodys_layouts as ml
from deck import open_deck, move_slide, set_notes, DECK
from icons import ICONS

def fields():
    return [
        {'head': 'Today', 'body': '‹ the workflow as it runs now, and where it hurts — one line ›',
         'icon': ml.md.find_icon(ICONS['today'])},
        {'head': 'The assistant', 'body': '‹ what it does — one line ›',
         'icon': ml.md.find_icon(ICONS['assistant'])},
        {'head': 'Grounded and governed', 'body': '‹ what it is grounded in, who signs off — one line ›',
         'icon': ml.md.find_icon(ICONS['governed'])},
        {'head': 'Status', 'body': '‹ production / pilot / discovery ›',
         'icon': ml.md.find_icon(ICONS['status'])},
    ]

DEMO = ('What we will show: ‹ 3-4 bullets ›\n'
        'Entry state: ‹ what is on screen before I start — set up before the talk ›\n'
        'Exit state: ‹ what they should have seen ›\n'
        'Fallback: ‹ screenshot or recording to cut to if it fails ›\n')

prs = open_deck()
ml.set_header_style('h1h2')
last = lambda: len(prs.slides._sldIdLst) - 1

s = ml.icon_text(prs, 'Upgrade Validation Assistant', fields(), subhead='Use case 2')
set_notes(s, 'Budget 1:00. CORE.')
move_slide(prs, last(), 11)

d = ml.divider(prs, 'DEMO', eyebrow='Use case 2')
set_notes(d, 'Upgrade Validation Assistant. Budget 4:00. CORE.\n' + DEMO +
             'Security artifact to surface: a guardrail refusing an out-of-scope request.')
move_slide(prs, last(), 12)

s = ml.icon_text(prs, 'Moody’s MCP Server', fields(), subhead='Use case 3')
set_notes(s, 'Budget 1:00. FLEX - the demo can carry this section alone if time is short.')
move_slide(prs, last(), 13)

d = ml.divider(prs, 'DEMO', eyebrow='Use case 3')
set_notes(d, 'Moody’s MCP Server. Budget 3:00. CORE.\n' + DEMO +
             'Security artifact to surface: tool permissions scoped per user and tenant.')
move_slide(prs, last(), 14)

prs.save(DECK)
print('use cases 2 and 3 added at indices 11-14')
"
```

- [ ] **Step 4: Run to verify it passes, then commit**

```bash
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python tools/verify_deck.py && git add tools/verify_deck.py && git commit -m "feat: add use cases 2 and 3"
```

Expected: `PASS - 19 slides in expected order`

---

### Task 9: Conclusion, back cover, disclaimer

**Files:** Modify `deck/UG_London_2026.pptx`, `tools/verify_deck.py`

**Interfaces:** Consumes `ml.icon_text`, `ml.back_cover`, `ml.disclaimer`, `move_slide`, `set_notes`, `ICONS`. Produces the conclusion at index 15 and the two closing slides at 20 and 21, after the appendix. Final deck: 22 slides.

- [ ] **Step 1: Add three expectations**

The conclusion is inserted at 15; the closing pair is appended after the appendix, so they need no move.

```python
    ("Trustworthy, in production", "Title Only"),   # 15  (insert after index 14)
    # ... appendix entries 16-19 stay as they are ...
    ("Thank you", "Back Cover 1"),                  # 20
    (None, "Disclaimer"),                           # 21  layout-only: see below
```

The disclaimer's signature is `None` deliberately. That slide has **zero shapes of its own** — the legal copy is a fixed shape on the `Disclaimer` layout, so slide-level text extraction returns an empty string and any signature would fail. Checking the layout name is the real assertion here.

- [ ] **Step 2: Run to verify it fails**

Expected: FAIL — count 19 vs 22.

- [ ] **Step 3: Author all three**

```bash
export SKILL=/Users/laumayp/.claude/plugins/cache/ibu-life-marketplace/shared-life-beta-ai-skills/0.8.0/skills/implementation-moodys-pptx
export MOODYS_ASSETS_DIR=/Users/laumayp/.claude/plugins/data/shared-life-beta-ai-skills-ibu-life-marketplace/moodys-assets
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx --with cairosvg python -c "
import sys, os; sys.path.insert(0, os.environ['SKILL'] + '/scripts'); sys.path.insert(0, 'tools')
import moodys_layouts as ml
from deck import open_deck, move_slide, set_notes, DECK
from icons import ICONS
prs = open_deck()
ml.set_header_style('h1h2')

s = ml.icon_text(prs, 'Trustworthy, in production', [
    {'head': 'Grounded, not guessing',
     'body': 'Answers cited to AXIS documentation and your own dataset.',
     'icon': ml.md.find_icon(ICONS['grounded'])},
    {'head': 'Governed access',
     'body': 'Security is platform engineering, not prompt engineering.',
     'icon': ml.md.find_icon(ICONS['access'])},
    {'head': 'Human sign-off',
     'body': 'Actuaries stay in control of anything touching a reported number.',
     'icon': ml.md.find_icon(ICONS['human'])},
    {'head': 'Transparent and auditable',
     'body': 'Cited outputs and logged actions your regulator can follow.',
     'icon': ml.md.find_icon(ICONS['audit'])},
    {'head': 'Layer three is open',
     'body': 'The open platform is a conversation we want to start. With this room.',
     'icon': ml.md.find_icon(ICONS['invite'])},
], subhead='Four principles, and one invitation')
set_notes(s, (
    'Budget 1:30. CORE.\n'
    'Close on the invitation, not the summary. This audience is exactly who layer three is for.\n'
    'Appendix follows: why Moody’s credentials, the data foundation, the actuarial AI curve.'
))
move_slide(prs, len(prs.slides._sldIdLst) - 1, 15)

ml.back_cover(prs, 'Thank you')
ml.disclaimer(prs)
prs.save(DECK)
print('total slides:', len(prs.slides._sldIdLst))
"
```

Expected: `total slides: 22`

- [ ] **Step 4: Run to verify it passes, then commit**

```bash
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python tools/verify_deck.py && git add tools/verify_deck.py && git commit -m "feat: add conclusion, back cover and disclaimer"
```

Expected: `PASS - 22 slides in expected order`

---

### Task 10: Speaker notes on the reused slides

The six reused slides arrived without timing or cut marks. Every main-line slide must carry its budget so the deck is rehearsable.

**Files:** Modify `deck/UG_London_2026.pptx`, create `tools/verify_notes.py`

**Interfaces:** Produces `tools/verify_notes.py`, asserting every main-line slide (0–15) has a `Budget` line and a CORE/FLEX mark.

- [ ] **Step 1: Write the notes verifier**

Create `tools/verify_notes.py`:

```python
#!/usr/bin/env python3
"""Assert every main-line slide carries a timing budget and a cut mark."""
import sys

from pptx import Presentation

DECK = "deck/UG_London_2026.pptx"
MAIN_LINE = 16   # slides 16-21 are appendix, back cover and disclaimer


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
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python tools/verify_notes.py
```

Expected: FAIL — the six reused slides (1, 2, 3, 4, 5, 7) have no notes.

- [ ] **Step 3: Add notes to the reused slides**

```bash
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python -c "
import sys; sys.path.insert(0, 'tools')
from deck import open_deck, set_notes, DECK
NOTES = {
  1: 'Budget 1:30. FLEX - can be cut to a 20-second verbal framing.\n'
     'Open Moody’s-wide for credibility: four agentic use cases already running.',
  2: 'Budget 1:30. FLEX - overlaps the previous slide and the next; first to cut.\n'
     'The three tiers: assistants, agentic solutions, AI-ready data.',
  3: 'Budget 1:30. CORE. Note MCP Servers among the channels - this plants use case 3.',
  4: 'Budget 2:00. CORE. This is the pivot from Moody’s-wide to Life and AXIS.\n'
     'Layer three is flagged here and cashed in at the conclusion.',
  5: 'Budget 1:30. CORE. Frontier tools, reusable blocks, built with clients.',
  7: 'Budget 1:30. CORE. The four principles. The trust spine the use cases hang from.',
}
prs = open_deck()
slides = list(prs.slides)
for i, text in NOTES.items():
    set_notes(slides[i], text)
prs.save(DECK)
print('notes written to', len(NOTES), 'slides')
"
```

- [ ] **Step 4: Run both verifiers, then commit**

```bash
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python tools/verify_deck.py && uv run --quiet --with python-pptx python tools/verify_notes.py && git add tools/verify_notes.py && git commit -m "feat: add timing budgets and cut marks to speaker notes"
```

Expected: `PASS - 22 slides in expected order` then `PASS - all 16 main-line slides carry budget and cut marks`

---

### Task 11: QA, render check, and README correction

**Files:** Create `build/UG_London_2026.pdf` (git-ignored), modify `README.md`

- [ ] **Step 1: Run the skill's structural QA**

```bash
export SKILL=/Users/laumayp/.claude/plugins/cache/ibu-life-marketplace/shared-life-beta-ai-skills/0.8.0/skills/implementation-moodys-pptx
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python "$SKILL/scripts/qa_check.py" deck/UG_London_2026.pptx
```

Expected: `errors: 0`.

**Warnings need triage, not blanket fixing.** The six reused slides carry roughly 13 pre-existing overflow and off-slide warnings — those are approved corporate slides and are left alone. Any warning on an authored slide (0, 6, 8, 9–15, 20, 21) is yours: fix it, re-run, repeat.

- [ ] **Step 2: Check the text came out right**

```bash
export SKILL=/Users/laumayp/.claude/plugins/cache/ibu-life-marketplace/shared-life-beta-ai-skills/0.8.0/skills/implementation-moodys-pptx
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python "$SKILL/scripts/extract_text.py" deck/UG_London_2026.pptx | head -80
```

Confirm: no French text survives on slide 6; every `‹ … ›` prompt is intact and none was accidentally filled with invented product detail.

- [x] **Step 3: Render to PDF — NOT POSSIBLE ON THIS MACHINE**

```bash
cd /Users/laumayp/Development/UG2026 && mkdir -p build && soffice --headless --convert-to pdf --outdir build deck/UG_London_2026.pptx && ls -la build/
```

**This does not work here.** `soffice` starts, sits at 0% CPU without touching the output directory, and is SIGKILLed by the OS (exit 137). Reproduced three times: sandboxed, sandboxed with an isolated `-env:UserInstallation` profile, and unsandboxed. The skill's `render.py` fares no better — its macOS PowerPoint backend runs the AppleScript without error but writes no files, then falls through to LibreOffice and reports `RENDER_BACKEND=none`.

Do not spend more time on this. Anyone re-running the plan should skip to Step 4.

- [x] **Step 4: Verify the authored slides — done without a renderer**

The three things Step 4 existed to catch are all checkable from the file itself, and two of the checks are stricter than eyeballing a LibreOffice approximation would have been:

| Original visual check | Replaced by |
|---|---|
| Body text does not overflow; titles unclipped | `qa_check.py` — its overflow and off-slide detectors flag **zero** issues on any authored slide (all 13 warnings sit on reused corporate slides 1, 2, 3, 17, 18) |
| Icons render as real glyphs, not blue circles | `tools/verify_slide_icons.py` — counts `<p:pic>` against `prst="ellipse"` per slide. 22 real icons, 0 placeholders |
| Icons resolve at all | `tools/verify_icons.py` — rasterizes each name rather than only resolving it |

`verify_slide_icons.py` is the important one. `_icon_or_dot` substitutes a blue ellipse of identical footprint whenever an icon fails to rasterize; that substitution passes `qa_check.py`, extracts as no text, and is invisible to every other check here. Counting the two shape kinds catches it deterministically.

**Still owed:** nobody has looked at this deck. Typography, spacing rhythm and the icon glyph choices are unverified. Open it in PowerPoint before the talk.

- [ ] **Step 5: Correct the README abstract**

The published abstract says the talk *turns to* security last; the agreed running order puts it at section 5, before the use cases. Change only the ordering claim:

Replace `then turn to what matters most for regulated teams: securing AI in production, with concrete examples over technical detail.` with `and set out what matters most for regulated teams — securing AI in production, with concrete examples over technical detail — before showing the work.`

- [ ] **Step 6: Final verification and commit**

```bash
export SKILL=/Users/laumayp/.claude/plugins/cache/ibu-life-marketplace/shared-life-beta-ai-skills/0.8.0/skills/implementation-moodys-pptx
cd /Users/laumayp/Development/UG2026 && uv run --quiet --with python-pptx python tools/verify_deck.py && uv run --quiet --with python-pptx python tools/verify_notes.py && uv run --quiet --with python-pptx python "$SKILL/scripts/qa_check.py" deck/UG_London_2026.pptx && git add README.md && git commit -m "docs: align README abstract with the talk running order"
```

Expected: both verifiers PASS and `errors: 0`.

---

## Handoff notes

Left deliberately unfilled, for the author rather than the implementer:

1. **Speaker names** on slide 0 — not recorded anywhere in the repo.
2. **All `‹ … ›` fields** on the six use case slides and their demo notes. None of these products is described in this repo, and inventing their capabilities would be worse than leaving the prompts visible.
3. **The finished deck is not versioned.** `resources/` is tracked; `deck/` is ignored because the working file is rewritten by every authoring task. Once final, commit it deliberately with `git add -f deck/UG_London_2026.pptx` — one commit of one binary, rather than nine.
4. **Use case 1's name** — built as "RiskIntegrity for IFRS 17 Navigator" per the README. If "Infoweb Navigator" is right, change it in `EXPECTED` and on slides 9 and 10.
5. **The reused slides' QA warnings** are pre-existing and untouched. If they should be fixed, that is a separate piece of work on corporate-approved slides.
