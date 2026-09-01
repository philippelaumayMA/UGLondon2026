# UG London 2026 deck — design

**Talk:** Bringing Trustworthy GenAI to Insurance
**Event:** User Group London 2026
**Duration:** ~30 minutes including three live demos
**Output:** one `.pptx`, English, on the Moody's corporate template

## Goal

Assemble a 17-slide main line that funnels from Moody's-wide GenAI down to three Life-insurance use cases, with a security beat in the middle. Sections 1–5 are built from existing material in `resources/`. Sections 6–9 are placeholders with enough scaffolding that the deck is rehearsable and timeable before the demos exist.

## Audience and register

A user group — existing Moody's customers, insurance and actuarial practitioners, not engineers. The abstract's constraint governs the writing: **concrete examples over technical detail**. The narrative opens Moody's-wide for credibility, then narrows to Life/AXIS at slide 5 and stays there.

Because the audience is a user group, the deck closes on the strategy's third layer — the open platform, which slide 5 explicitly frames as *"a conversation we want to start."* This room is that conversation's audience, which makes it a call to action rather than a summary.

## Build method

**Start from a duplicate of `resources/Consolidated GenAI common slides.pptx`**, reorder the existing slides into the spine, then author the new slides with the `implementation-moodys-pptx` skill's `moodys_layouts` (`ml`) builders — opened directly on that deck, with **no `reset_slides`**, which the skill supports for editing an existing deck.

This works because the Consolidated deck *is* the skill's official template. Verified, not assumed: the two files have byte-identical theme colour schemes, the same `Disclaimer` layout (`slideLayout64.xml`, 10 650 characters), and every one of the bundled template's 64 master-1 layout names is present in the Consolidated deck. `ml._layout()` resolves layouts by name against `slide_masters[0]`, so each builder lands on the correct official layout.

Building this way rather than hand-filling placeholders gets the TOOLKIT type ramp, the light/dark rhythm and the official icon set for free.

### Blocker: slide IDs must be renumbered first

The Consolidated deck's `p:sldId` values are pathological — `[…, 2147483635, 2147483646, …]`, one below the int32 ceiling of 2147483647. python-pptx assigns `max(ids) + 1` to each added slide, so **the deck accepts exactly one new slide and then raises `ValueError`**. Renumbering all slide IDs to a low contiguous range (256, 257, …) before any insertion fixes it permanently. Slide IDs are arbitrary internal identifiers, independent of the `r:id` relationships, so renumbering is lossless. This must happen before any authoring step.

### Not from DEVOXX

Nothing is copied out of the DEVOXX file — it has 1 master to Consolidated's 2, and different media. Its only contribution is text for slide 7, which is translated from French and condensed from two slides into one, so re-authoring it as an `ml.columns` slide loses nothing. If a specific DEVOXX diagram is wanted later (the agentic architecture build-up, the guardrails belt), export it as an image and place it deliberately.

### Skill obligations

- `ml.set_header_style("h1h2")` once before authoring; every content builder gets a `subhead`.
- `ml.disclaimer(prs)` is **mandatory and always last**, preceded by `ml.back_cover(prs, "Thank you")`.
- Icons on every item-based builder. The shell needs `MOODYS_ASSETS_DIR=/Users/laumayp/.claude/plugins/data/shared-life-beta-ai-skills-ibu-life-marketplace/moodys-assets`, because `CLAUDE_PLUGIN_DATA` is not set outside the skill's own runtime. Resolve every name with `ml.md.find_icons()` first — `shield` and `check` return nothing, and an unresolved name degrades silently to a blue circle.
- Run everything through `uv run --with python-pptx` (add `--with cairosvg` for icons).

## Slide map

Sources: `C-n` = Consolidated slide *n*, `D-n` = DEVOXX slide *n*. Every reused slide keeps its existing content unchanged. Slides marked ★ already carry the exact section title requested, so they need no edit at all. Builders choose their own official layout — do not override it.

| # | Section | Source | Action | Builder |
|---|---|---|---|---|
| 0 | Title | — | New | `ml.cover` |
| 1 | Agenda | — | New | `ml.agenda` |
| 2 | 1. Intro | C-1 | Reuse | — |
| 3 | 1. Intro | C-5 | Reuse | — |
| 4 | 2. Every stage of AI | C-4 | Reuse ★ | — |
| 5 | 3. Strategy, three layers | C-6 | Reuse ★ | — |
| 6 | 4. Our approach | C-8 | Reuse ★ | — |
| 7 | 5. Security landscape | D-6, D-9 | New, translated | `ml.columns` |
| 8 | 5. Security at Moody's | C-9 | Reuse | — |
| 9 | 6. Three use cases | — | New | `ml.columns` |
| 10 | 7. IFRS 17 Navigator | — | New placeholder | `ml.icon_text` |
| 11 | 7. Demo | — | New demo card | `ml.divider` |
| 12 | 8. UVA | — | New placeholder | `ml.icon_text` |
| 13 | 8. Demo | — | New demo card | `ml.divider` |
| 14 | 9. MCP Server | — | New placeholder | `ml.icon_text` |
| 15 | 9. Demo | — | New demo card | `ml.divider` |
| 16 | 10. Conclusion | — | New | `ml.icon_text` |

Consolidated slides 2, 3, 7 and 10 are not used in the main line. Keep them as an appendix at indices 17–20 rather than deleting: C-3 (data foundation, 590M+ entities) and C-10 (Why Moody's, 130+/100+/30+) are the natural answers to a credibility challenge from the floor, and C-7 (the Actuarial AI Curve) answers "where should we start?"

Two closing slides follow the appendix, giving **23 slides in total**:

| # | Slide | Builder |
|---|---|---|
| 21 | Thank you | `ml.back_cover` |
| 22 | Legal disclaimer | `ml.disclaimer` |

The disclaimer is mandatory and must be the final slide.

### Two links that come for free

- Slide 4 already names **MCP Servers** among the access channels, planting use case 3 twenty minutes before it lands.
- Slide 5's three layers give slide 9 its structure: each use case is tagged to the layer it proves.

## New slide content

### Slide 0 — Title
Talk title, event, date, speaker names. Speaker attribution is unresolved (see Open items).

### Slide 1 — Agenda
Five numbered rows on the template's `Agenda 2` layout (dark left panel, vertical title, big-number table). Rows name the two framing sections and then each use case by product, so the room knows three demos are coming:

1. Moody's AI strategy — three layers, and how we build on them
2. Securing GenAI in production — regulatory pressure, OWASP, and how we answer both
3. RiskIntegrity for IFRS 17 Navigator — use case 1, with a live demo
4. Upgrade Validation Assistant — use case 2, with a live demo
5. Moody's MCP Server — use case 3, with a live demo

The conclusion is deliberately not a row: it is an invitation, and announcing it up front spends it early.

This slide is where the running order is stated out loud — security lands *before* the use cases, which is the opposite of what the originally published abstract implied. The speaker note says to call that out.

**Known QA warning, not a defect.** `qa_check.py` reports `off-slide: shape 'Title 2'` here. The `Agenda 2` layout positions its title at x = −1.68 in and rotates it 270° (`rot="16200000"`); the checker compares raw `left/top/width/height` and has no rotation term, so the pre-rotation box looks off-slide. Rotated about its centre the title occupies x 0.01–2.15, y 0.60–6.11 — inside the dark panel, fully on-slide. The geometry is inherited verbatim from the official template's own layout. Leave it.

### Slide 7 — The security landscape
Condensed and translated from D-6 and D-9, as two `ml.columns` items, each with an icon:

- **Regulatory pressure.** The EU AI Act as precursor: a deliberately broad definition of AI covering both generative and autonomous-action tooling, with obligations scaling to the risk posed to the user. A worldwide trend, not a European one.
- **The OWASP GenAI Security Project.** The working checklist. Show the entries that bite in agentic systems — LLM01 Prompt Injection, LLM02 Sensitive Information Disclosure, LLM06 Excessive Agency, LLM07 System Prompt Leakage — and note each carries vulnerability, scenario and mitigation. The full ten go in this slide's speaker notes, where they are available for a question from the floor without costing slide space.

Speaker note carries D-4's framing: the exposure is reputational, human and financial.

### Slide 8 — Responsible AI (C-9, reused)
Already English, already on-brand, already Life-specific: grounded not guessing; governed access; actuaries stay in control; transparent and auditable. Reused as the "specifically by Moody's" half of section 5.

The D-7/D-8 risk-and-control matrix, labelled *"L'exemple Moody's"*, covers similar ground and would need translating. Hold it in the appendix as the answer to a governance question, not in the main line.

### Slide 9 — Three use cases, one foundation
Three `ml.columns` items, each with an icon, each tagged to its strategy layer:

| Use case | Layer | Note |
|---|---|---|
| RiskIntegrity for IFRS 17 Navigator | Applications | In production today |
| Upgrade Validation Assistant (UVA) | Applications | |
| Moody's MCP Server | Open platform | Your agents, our data |

### Slides 10, 12, 14 — Use case placeholders
An `ml.icon_text` slide each, one item and one icon per field. Same four fields every time, to be filled once the products are briefed:

1. The workflow as it runs today, and where it hurts
2. What the assistant does
3. What is grounded and what is governed — the trust hook that ties back to slide 8
4. Status: production, pilot, or in discovery

### Slides 11, 13, 15 — Demo cards
An `ml.divider` (which resolves to `Divider 3 - Short title`) carrying the word DEMO and the use case name as its eyebrow — a deliberate "look up from the slides" cue. All scaffolding below lives in the **speaker notes**, not on the slide face: the audience must never see the fallback plan, and a divider has no room for it in any case. Each card's notes carry:

- **What we'll show** — 3–4 bullets
- **Entry state** — what is on screen when the demo starts, so it can be set up before the talk
- **Exit state** — what the audience should have seen by the end
- **Fallback** — the screenshot or recording to cut to if the demo fails live
- **Timing** — budget from the table below

Since section 5 is only two slides and carries no demo of its own, surface one security artifact inside each demo — a citation in Navigator, a guardrail in UVA, scoped tool permissions on the MCP server. This is the cheapest way to honour the abstract's promise of concrete security examples without spending more slides.

### Slide 16 — Conclusion
An `ml.icon_text` slide restating the four Responsible AI principles as takeaways, then making the ask: layer three, the open platform, is a conversation Moody's wants to start with this room. Ends on an invitation rather than a summary.

The originally specified `Executive Summary/Key Takeaways 2` layout is dropped: no `ml` builder targets it, and dropping to `md` primitives for one slide would cost more than it returns.

## Timing budget

Thirty minutes across 17 slides and three demos leaves roughly one minute per content slide. The budget:

| Slides | Content | Budget |
|---|---|---|
| 0 | Title | 0:30 |
| 1 | Agenda | 0:30 |
| 2–3 | Intro | 3:00 |
| 4 | Every stage | 1:30 |
| 5 | Three layers | 2:00 |
| 6 | Our approach | 1:30 |
| 7–8 | Security | 3:00 |
| 9 | Use case intro | 1:00 |
| 10–11 | Navigator + demo | 5:00 |
| 12–13 | UVA + demo | 5:00 |
| 14–15 | MCP + demo | 4:00 |
| 16 | Conclusion | 1:30 |
| | **Total** | **28:30** |

Ninety seconds of slack for transitions and a question. That is thin, and live demos overrun rather than underrun.

### Cut list

Every slide is marked CORE or FLEX. FLEX slides can be dropped mid-talk without breaking the narrative, in this order:

1. **Slide 3** (AI Solution Suite) — overlaps slides 2 and 4; saves ~1:30
2. **Slide 7** (security landscape) — compress to one spoken sentence over slide 8; saves ~1:30
3. **Slide 14** (MCP context) — the demo carries itself; saves ~1:00
4. **Slide 2** — reduce to a 20-second verbal framing; saves ~1:00
5. **Slide 1** (agenda) — read the five rows out in one breath, or skip the slide entirely if the preceding session overran; saves ~0:30

CORE: 0, 4, 5, 6, 8, 9, 10, 11, 12, 13, 15, 16. Dropping all five FLEX slides recovers five and a half minutes, which covers one badly overrunning demo.

## Open items

1. ~~**The README abstract no longer matches the running order.**~~ Resolved: `README.md` now reads *"and set out what matters most for regulated teams — securing AI in production … — before showing the work,"* which matches security landing at section 5. If the abstract was already published elsewhere, that copy still reads the other way round; the agenda slide's speaker note says to call the change out loud.
2. **Use case 1's name.** The flow said "Infoweb Navigator"; the deck uses the README's "RiskIntegrity for IFRS 17 Navigator". Correct at build time if the former is the real product name.
3. **Speaker attribution** for the title slide is not recorded anywhere in the repo. The DEVOXX deck credits Loïc Gudet and Philippe Laumay; do not assume that carries over.
4. ~~**`resources/` is untracked.**~~ Resolved: the two source decks are committed, and `.gitignore` excludes the built deck (`deck/`), renders (`build/`) and Office lock files (`~$*`). The deck is regenerable from `resources/` plus `tools/`, so it does not belong in git.
5. **The three use case slides are still placeholders.** Their four fields carry `‹ … ›` guillemet prompts, deliberately, because no product facts about Navigator, UVA or the MCP server exist in this repo. They must be briefed before the deck is presentable.
