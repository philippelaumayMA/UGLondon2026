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
    # --- spine ---
    ("Bringing Trustworthy GenAI to Insurance", "Cover 1"),  # 0  title
    ("Process Automations", None),        # 1  C-1  intro: agentic solutions
    ("AI Solution Suite", None),          # 2  C-5  intro: the three tiers
    ("Every Stage of AI Maturity", None), # 3  C-4  section 2
    ("in three layers", None),            # 4  C-6  section 3
    ("OUR APPROACH", None),               # 5  C-8  section 4
    ("OWASP", "Title Only"),              # 6  security landscape
    ("RESPONSIBLE AI", None),             # 7  C-9  section 5b
    ("one foundation", "Title Only"),     # 8  three use cases
    ("RiskIntegrity for IFRS 17 Navigator", "Title Only"),  # 9
    ("DEMO", "Divider 3 - Short title"),                    # 10
    # All three demo cards carry the same on-slide word, so "DEMO" is not a
    # unique signature -- two swapped demo cards would pass. The unique context
    # slide on either side pins the ordering, which is enough here.
    ("Upgrade Validation Assistant", "Title Only"),         # 11
    ("DEMO", "Divider 3 - Short title"),                    # 12
    ("MCP Server", "Title Only"),                           # 13
    ("DEMO", "Divider 3 - Short title"),                    # 14
    ("Trustworthy, in production", "Title Only"),           # 15 conclusion
    # --- appendix ---
    ("Driven by Perspective", None),      # 16 C-2
    ("Content Foundation", None),         # 17 C-3
    ("ACTUARIAL AI CURVE", None),         # 18 C-7
    ("WHY MOODY", None),                  # 19 C-10
    # --- closing ---
    ("Thank you", "Back Cover 1"),        # 20
    # The disclaimer's signature is None deliberately: that slide has zero
    # shapes of its own, because the legal copy is a fixed shape on the
    # Disclaimer layout. Slide-level text extraction returns "", so any
    # signature would fail; the layout name is the real assertion.
    (None, "Disclaimer"),                 # 21
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
