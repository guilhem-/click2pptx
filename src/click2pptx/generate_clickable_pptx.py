#!/usr/bin/env python3
"""
generate_clickable_pptx.py
-------------------------

Convert a Freeplane HTML export containing an image map into a PowerPoint
presentation where each active area is recreated as a fully transparent
rectangle linking to the corresponding external URL.

Usage
=====
    python generate_clickable_pptx.py [-i SOURCE.html] [-o DESTINATION.pptx]

    • -i / --input   : source HTML file. If omitted, the *first* ``*.html``
                       found in the current directory is used.
    • -o / --output : path of the PPTX file to create. If omitted, the
                       script creates an ``output`` folder (if needed) and
                       writes ``mind_map_clickable_YYYYMMDD_HHMMSS.pptx``
                       inside it.

Dependencies: ``beautifulsoup4``, ``python-pptx``.
"""

import argparse
import glob
import os
import re
import sys
from datetime import datetime

from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_FILL
EMU_PER_PX = 9_525  # 1 px → 9 525 EMU (pptx)

# ---------------------------------------------------------------------------#
# I/O utilities
# ---------------------------------------------------------------------------#

def find_html(path: str | None) -> str:
    """Return the path of the HTML file to process."""
    if path:
        if not os.path.isfile(path):
            sys.exit(f"Source file not found: {path}")
        return path

    matches = glob.glob("*.html") + glob.glob("*.htm")
    if not matches:
        sys.exit("No *.html file found and no --input provided.")
    return matches[0]


def make_output_path(path: str | None) -> str:
    """Build the output path (create ./output if needed)."""
    if path:
        out_dir = os.path.dirname(path)
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)
        return path

    out_dir = "output"
    os.makedirs(out_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(out_dir, f"mind_map_clickable_{ts}.pptx")


# ---------------------------------------------------------------------------#
# Freeplane HTML parsing
# ---------------------------------------------------------------------------#

def parse_html(html_text: str):
    """
    Extracts:
      • the list of clickable areas ``([x1, y1, x2, y2], url)``
      • the optional path to the background image
    """
    soup = BeautifulSoup(html_text, "html.parser")

    imagemap = soup.find("map", id="fm_imagemap")
    if not imagemap:
        return [], None

    # internal id → external http(s) link
    id2url = {}
    for a_internal in soup.find_all("a", id=re.compile(r"FMID_\d+FM")):
        external = a_internal.find_next("a", href=re.compile(r"^https?://"))
        if external:
            id2url[a_internal["id"]] = external["href"]

    clickables = []
    for area in imagemap.find_all("area"):
        coords = list(map(int, area["coords"].split(",")))
        internal_id = area["href"].lstrip("#")
        url = id2url.get(internal_id)
        if url:
            clickables.append((coords, url))

    img_tag = soup.find("img", {"usemap": "#fm_imagemap"})
    img_src = img_tag["src"] if img_tag else None
    return clickables, img_src


# ---------------------------------------------------------------------------#
# PPTX generation
# ---------------------------------------------------------------------------#

def create_pptx(clickables, img_src, dest_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    if img_src and os.path.isfile(img_src):
        slide.shapes.add_picture(img_src, Emu(0), Emu(0))  # background image

    for coords, url in clickables:
        x1, y1, x2, y2 = coords
        left, top = Emu(x1 * EMU_PER_PX), Emu(y1 * EMU_PER_PX)
        width, height = Emu((x2 - x1) * EMU_PER_PX), Emu((y2 - y1) * EMU_PER_PX)

        rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        # MAKE THE RECTANGLE INVISIBLE
        rect.fill.background()
        rect.line.fill.background()

        rect.click_action.hyperlink.address = url

    prs.save(dest_path)


# ---------------------------------------------------------------------------#
# Main program
# ---------------------------------------------------------------------------#

def main():
    ap = argparse.ArgumentParser(
        description="Convert a Freeplane HTML export into a clickable PPTX."
    )
    ap.add_argument("-i", "--input", help="Source HTML file")
    ap.add_argument("-o", "--output", help="Destination PPTX file")
    args = ap.parse_args()

    html_path = find_html(args.input)
    output_path = make_output_path(args.output)

    with open(html_path, "r", encoding="utf-8") as f:
        html_text = f.read()

    clickables, img_src = parse_html(html_text)
    if not clickables:
        sys.exit("No clickable area with external link found in the HTML.")

    create_pptx(clickables, img_src, output_path)
    print(f"Generated PPTX: {output_path}")


if __name__ == "__main__":
    main()
