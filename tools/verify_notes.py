#!/usr/bin/env python3
"""Assert every main-line slide carries a timing budget and a cut mark."""
import sys

from pptx import Presentation

DECK = "deck/UG_London_2026.pptx"
MAIN_LINE = 17   # slides 17-22 are appendix, back cover and disclaimer


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
