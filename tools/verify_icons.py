#!/usr/bin/env python3
"""Assert every name in tools/icons.py resolves against the official icon set.

An unresolved name does not raise -- it degrades silently to a bright-blue
circle placeholder on the slide, which is only visible in a render. This check
catches it at build time instead.

Needs both deps and the assets dir:

    export SKILL=.../skills/implementation-moodys-pptx
    export MOODYS_ASSETS_DIR=.../moodys-assets
    uv run --with python-pptx --with cairosvg python tools/verify_icons.py
"""
import os
import sys

sys.path.insert(0, os.environ["SKILL"] + "/scripts")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import moodys_deck as md  # noqa: E402

from icons import ICONS  # noqa: E402


def main():
    md.ensure_icons()
    bad = [k for k, v in ICONS.items() if md.find_icon(v) is None]
    if bad:
        print(f"FAIL - {len(bad)} unresolved icon name(s)")
        for k in bad:
            print(f"  - {k}: {ICONS[k]!r}")
        return 1
    print(f"PASS - all {len(ICONS)} icon names resolve")
    return 0


if __name__ == "__main__":
    sys.exit(main())
