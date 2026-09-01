#!/usr/bin/env python3
"""Assert every name in tools/icons.py resolves AND rasterizes.

Two distinct failure modes, both silent on the slide:

1. A wrong name -- `find_icon` returns None and the slide gets a bright-blue
   circle placeholder.
2. A right name that cannot be rasterized -- `icon_png` returns None (cairosvg
   missing, or a malformed SVG) and the slide gets the same placeholder.

Checking `find_icon` alone passes while every icon on the deck is a circle, so
this rasterizes each one for real.

Needs both deps, the assets dir, and the native cairo library:

    export SKILL=.../skills/implementation-moodys-pptx
    export MOODYS_ASSETS_DIR=.../moodys-assets
    export DYLD_FALLBACK_LIBRARY_PATH=/opt/homebrew/lib:/usr/local/lib:/usr/lib
    uv run --with python-pptx --with cairosvg python tools/verify_icons.py

`DYLD_FALLBACK_LIBRARY_PATH` is needed on macOS: cairocffi dlopens `libcairo`
by bare name and does not search Homebrew's prefix, so without it the import
raises `OSError: no library called "cairo-2" was found`.
"""
import os
import sys

sys.path.insert(0, os.environ["SKILL"] + "/scripts")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import moodys_deck as md  # noqa: E402

from icons import ICONS  # noqa: E402


def main():
    md.ensure_icons()
    errors = []
    for key, name in ICONS.items():
        svg = md.find_icon(name)
        if svg is None:
            errors.append(f"{key}: name {name!r} is not in the official set")
            continue
        if md.icon_png(svg, 256) is None:
            errors.append(f"{key}: {name!r} resolves but will not rasterize")
    if errors:
        print(f"FAIL - {len(errors)} icon problem(s); these render as blue circles")
        for e in errors:
            print("  -", e)
        return 1
    print(f"PASS - all {len(ICONS)} icons resolve and rasterize")
    return 0


if __name__ == "__main__":
    sys.exit(main())
