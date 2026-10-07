#!/usr/bin/env python
"""
Extract icons from a PowerPoint "icon sheet" (pictures laid out on slides, e.g. the ArcGIS Visio
toolkit exported to .pptx) into transparent PNG files for add_icon / list_icons.

* Pictures are taken slide by slide in reading order (rows top to bottom, left to right) and named
  from a JSON file: {"1": ["Name 1", "Name 2", ...], "2": [...]}. The number of names must match
  the number of pictures on that slide.
* Labels baked into the pictures (black text under the icon) are cut off, so you can place your own
  labels. Use --keep-label to keep them.
* WMF/EMF pictures are rendered through GDI, so this needs Windows.

Usage:
    python scripts/extract_pptx_icons.py SHEET.pptx --names names.json [--dest icons/arcgis]
        [--prefix arcgis-] [--dpi 400] [--keep-label] [--list]

Notes:
  * Vendor icon sets have their own licence terms: keep the output local (icons/ is gitignored).
  * Single-colour line icons are converted to that colour on a transparent background; icons with
    several colours keep their colours with transparency derived from brightness.
"""
import argparse
import collections
import io
import json
import re
import sys
from pathlib import Path

from PIL import Image, ImageChops
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

EMU_PER_IN = 914400
ROW_GAP_IN = 0.8  # pictures whose tops differ by more than this start a new row


def reading_order(pictures):
    pics = sorted(pictures, key=lambda s: s.top)
    rows, row, row_top = [], [], None
    for p in pics:
        if row_top is None or (p.top - row_top) / EMU_PER_IN > ROW_GAP_IN:
            if row:
                rows.append(row)
            row, row_top = [], p.top
        row.append(p)
    if row:
        rows.append(row)
    return [p for r in rows for p in sorted(r, key=lambda s: s.left)]


def render(blob, dpi):
    im = Image.open(io.BytesIO(blob))
    if im.format in ("WMF", "EMF"):
        im.load(dpi=dpi)
    return im.convert("RGB")


def is_text(r, g, b):
    return max(r, g, b) < 60 and max(r, g, b) - min(r, g, b) < 25


def cut_label(im):
    """Remove the baked-in black label: everything from the first row with black text pixels."""
    w, h = im.size
    data = im.tobytes()
    first = None
    for y in range(h):
        row = data[y * w * 3:(y + 1) * w * 3]
        if any(is_text(row[i], row[i + 1], row[i + 2]) for i in range(0, len(row), 3)):
            first = y
            break
    return im.crop((0, 0, w, first)) if first and first > h * 0.25 else im


def to_transparent(im):
    """White background -> transparent. Single-colour icons become that colour with alpha."""
    r, g, b = im.split()
    minc = ImageChops.darker(ImageChops.darker(r, g), b)
    data = im.tobytes()
    buckets = collections.Counter()
    for i in range(0, len(data), 3):
        if min(data[i], data[i + 1], data[i + 2]) < 160:
            buckets[(data[i] // 24, data[i + 1] // 24, data[i + 2] // 24)] += 1
    total = sum(buckets.values())
    if total == 0:
        return im.convert("RGBA")
    if buckets.most_common(1)[0][1] / total >= 0.8:
        # solid colour = the darkest-channel average of pixels in the dominant bucket
        key = buckets.most_common(1)[0][0]
        px = [(data[i], data[i + 1], data[i + 2]) for i in range(0, len(data), 3)
              if (data[i] // 24, data[i + 1] // 24, data[i + 2] // 24) == key]
        col = tuple(sum(p[c] for p in px) // len(px) for c in range(3))
        scale = 255 - min(col)
        alpha = minc.point(lambda v: min(255, int((255 - v) * 255 / max(scale, 1))))
        out = Image.new("RGBA", im.size, col + (0,))
        out.putalpha(alpha)
        return out
    out = im.convert("RGBA")
    out.putalpha(minc.point(lambda v: 255 - v))
    return out


def safe_name(name):
    name = re.sub(r'[\\/:*?"<>|\r\n\t%]+', " ", name).strip()
    return re.sub(r"\s+", " ", name)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx")
    ap.add_argument("--names", required=True, help='JSON: {"1": ["name", ...], "2": [...]}')
    ap.add_argument("--dest", default=str(Path(__file__).resolve().parent.parent / "icons" / "icons-from-pptx"))
    ap.add_argument("--prefix", default="")
    ap.add_argument("--dpi", type=int, default=400)
    ap.add_argument("--keep-label", action="store_true")
    ap.add_argument("--list", action="store_true", help="Only list slide/picture/name mapping")
    args = ap.parse_args()

    names = json.loads(Path(args.names).read_text(encoding="utf-8"))
    prs = Presentation(args.pptx)
    dest = Path(args.dest)
    if not args.list:
        dest.mkdir(parents=True, exist_ok=True)
    done, used, problems = 0, set(), []
    for number, slide in enumerate(prs.slides, start=1):
        pics = reading_order([s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE])
        slide_names = names.get(str(number))
        if slide_names is None:
            problems.append(f"slide {number}: no names given ({len(pics)} pictures) - skipped")
            continue
        if len(slide_names) != len(pics):
            problems.append(f"slide {number}: {len(pics)} pictures but {len(slide_names)} names - skipped")
            continue
        for pic, name in zip(pics, slide_names):
            final = args.prefix + safe_name(name)
            if final.lower() in used:
                problems.append(f"slide {number}: duplicate name '{final}' - skipped")
                continue
            used.add(final.lower())
            if args.list:
                print(f"slide {number}: {name}")
                continue
            try:
                im = render(pic.image.blob, args.dpi)
                if not args.keep_label:
                    im = cut_label(im)
                out = to_transparent(im)
                out = out.crop(out.getchannel("A").getbbox() or (0, 0) + out.size)
                if max(out.size) > 512:
                    out.thumbnail((512, 512), Image.LANCZOS)
                out.save(dest / f"{final}.png")
                done += 1
            except Exception as e:
                problems.append(f"slide {number}: {name}: {e}")
    if not args.list:
        print(f"Wrote {done} icons to {dest}")
    for p in problems:
        print("NOTE", p)
    return 0


if __name__ == "__main__":
    sys.exit(main())
