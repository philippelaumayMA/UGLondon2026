# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repository is

Preparation material for a conference talk — **"Bringing Trustworthy GenAI to Insurance"** (User Group London 2026; git remote is `philippelaumayMA/UGLondon2026`). It is **not a software project**: there is no build system, no dependencies, no tests, and no source code. Do not invent or scaffold any of these unless asked.

The deliverable is a presentation. Work here means writing, editing, and restructuring slide decks and speaker material.

### Session abstract (from `README.md`)

> Generative AI is reshaping how insurers research, validate, and navigate regulatory work — but in a risk-driven industry, trust is non-negotiable. During this talk we cover Moody's AI strategy, show GenAI at work in Life insurance through the Upgrade Validation Assistant, RiskIntegrity for IFRS 17 Navigator, and a Moody's MCP server, then turn to what matters most for regulated teams: securing AI in production, with concrete examples over technical detail.

This abstract is the contract for the talk. The four beats it promises — Moody's AI strategy, three Life-insurance GenAI demos, then production security — are what the deck has to deliver. `README.md` is currently the only committed file, so treat it as the source of truth for scope.

Note the abstract's closing constraint: **concrete examples over technical detail**. The audience is regulated-industry practitioners, not engineers.

## Working with the decks

`resources/` holds the source material. Both files are **untracked** and large (~10 MB and ~25 MB); there is no `.gitignore`. Confirm with the user before committing them — binaries this size in git history are hard to undo.

| File | Slides | Language | Role |
|---|---|---|---|
| `Consolidated GenAI common slides.pptx` | 10 | English | Moody's corporate boilerplate — AI strategy, Agentic Solutions, AI Solution Suite. Source for the strategy section. |
| `DEVOXX_2026_les_gardiens_du_prompt.pptx` | 37 | French | A prior talk on GenAI security. Source for the security section. |

**Use the `shared-life-beta-ai-skills:implementation-moodys-pptx` skill** for any read, edit, extract, or create operation on these `.pptx` files — it carries the Moody's template conventions. For a quick peek without the skill, a `.pptx` is a zip: `unzip -p FILE docProps/app.xml` lists slide titles, `unzip -p FILE ppt/slides/slide3.xml` dumps one slide's XML.

### The `DEVOXX_2026` deck

The bulk of reusable security content. Its arc: AI application risks (reputational, regulatory pressure) → a risk/control matrix → OWASP GenAI Security Project → an insecure-application walkthrough → a six-slide build-up of an Agentic AI architecture → live prompt demos → then the mitigations: baseline security principles, observability (platform survey, then Langfuse specifically), LLM guardrails, vulnerability scanning, and data residency.

Two adaptations are needed before this material fits the London session: it is **in French**, and it was written for a developer conference audience — Devoxx — rather than the insurance practitioners the abstract targets.

## Related projects in this workspace

The parent workspace guide at `/Users/laumayp/Development/CLAUDE.md` covers these; they are relevant here as demo sources:

- `snowcamp2026_talk_demo/` — SvelteKit app demonstrating LLM security (prompt injection), using Langfuse for observability. Directly matches the DEVOXX deck's demo and observability sections.
- `copilot/` — Moody's YOU 2.0 enterprise GenAI platform, a production system in the space the talk describes.
