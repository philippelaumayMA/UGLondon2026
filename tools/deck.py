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
