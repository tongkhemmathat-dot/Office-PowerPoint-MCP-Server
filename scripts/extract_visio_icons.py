#!/usr/bin/env python
"""
Extract icons from Visio files (.vsdx diagrams and .vssx stencils) into PNG files that the
add_icon / list_icons tools can use.

Two methods, combined automatically:
  1. Masters whose artwork is a single embedded picture (EMF / PNG / BMP / JPEG) are read straight
     from the file's zip. No Visio needed; EMF conversion needs Windows (Pillow renders via GDI).
  2. Masters drawn from Visio's own vector geometry (e.g. the built-in "Server", "Switch",
     "Web Server") are exported by Visio itself through COM. Needs Windows, Microsoft Visio and
     pywin32 (`pip install pywin32`). Without them these masters are only listed as skipped.

Usage:
    python scripts/extract_visio_icons.py FILE_OR_DIR [...] [--dest icons/visio] [--size 512]
                                          [--prefix visio-] [--only "Server,Firewall"]
                                          [--visio auto|off|only] [--list]

Notes:
  * Output files are named <prefix><master name>.png (default prefix "visio-", so they never
    clash with other icon sets, e.g. visio-Server.png).
  * Vendor stencils carry their own licence terms. Keep the output local (icons/ is gitignored)
    and do not commit it.
  * Connectors / lines (1-D shapes) are skipped. Identical pictures are saved once; different
    masters with the same name get -2, -3, ...
  * Backgrounds are made transparent by flood-filling white from the corners, so artwork that is
    white and touches the border may lose a little.
  * Visio is attached to if already running (your session is left open); otherwise a hidden
    instance is started and quit at the end.
"""
import argparse
import hashlib
import io
import os
import posixpath
import re
import shutil
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from PIL import Image, ImageDraw

NS = {"v": "http://schemas.microsoft.com/office/visio/2012/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
RID = "{%s}id" % NS["r"]


# ---------- method 1: pictures embedded in the zip ----------
def read_rels(z, part):
    d, b = posixpath.split(part)
    p = posixpath.join(d, "_rels", b + ".rels")
    if p not in z.namelist():
        return {}
    return {e.get("Id"): posixpath.normpath(posixpath.join(d, e.get("Target")))
            for e in ET.fromstring(z.read(p)).findall("rel:Relationship", NS)}


def iter_masters(z):
    if "visio/masters/masters.xml" not in z.namelist():
        return
    root = ET.fromstring(z.read("visio/masters/masters.xml"))
    mrels = read_rels(z, "visio/masters/masters.xml")
    for m in root.findall("v:Master", NS):
        rel = m.find("v:Rel", NS)
        if rel is not None and rel.get(RID) in mrels:
            yield (m.get("Name") or m.get("NameU") or "master"), mrels[rel.get(RID)]


def classify(z, part):
    """('image', media_path) for a single-picture master, else ('skip', reason)."""
    root = ET.fromstring(z.read(part))
    rels = read_rels(z, part)
    pics, vectors = [], 0

    def walk(el):
        nonlocal vectors
        for s in el.findall("v:Shape", NS):
            fd = s.find("v:ForeignData", NS)
            if fd is not None:
                rel = fd.find("v:Rel", NS)
                if rel is not None and rel.get(RID) in rels:
                    pics.append(rels[rel.get(RID)])
            elif any(x.get("N") == "Geometry" for x in s.findall("v:Section", NS)):
                vectors += 1
            sub = s.find("v:Shapes", NS)
            if sub is not None:
                walk(sub)

    shapes = root.find("v:Shapes", NS)
    if shapes is not None:
        walk(shapes)
    if len(pics) == 1 and vectors == 0:
        return "image", pics[0]
    if pics and vectors:
        return "skip", "picture + vector parts"
    if len(pics) > 1:
        return "skip", "several pictures"
    return "skip", "vector shape"


# ---------- image post-processing ----------
def clear_white_background(im):
    im = im.convert("RGBA")
    w, h = im.size
    for pt in ((0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1)):
        px = im.getpixel(pt)
        if px[3] > 0 and min(px[:3]) > 240:
            ImageDraw.floodfill(im, pt, (255, 255, 255, 0), thresh=14)
    return im


def finish(im, size):
    bbox = im.getchannel("A").getbbox() or (0, 0) + im.size
    im = im.crop(bbox)
    scale = size / max(im.size)
    if scale < 1 or max(im.size) < size // 2:
        im = im.resize((max(1, round(im.width * scale)), max(1, round(im.height * scale))), Image.LANCZOS)
    return im


def picture_to_png(data, ext, size):
    im = Image.open(io.BytesIO(data))
    if ext in (".emf", ".wmf"):
        im.load(dpi=600)  # the nominal EMF size is often tiny: render large, then scale
        im = im.convert("RGB")
        if max(im.size) > 2400:
            im.thumbnail((2400, 2400), Image.LANCZOS)
    return finish(clear_white_background(im), size)


# ---------- method 2: Visio COM ----------
def visio_export(src, names, size, emit):
    """Export the named masters of `src` through Visio. `emit(name, PIL image)` receives each one.
    Returns names that failed or were skipped as connectors."""
    import win32com.client
    owns = False
    try:
        app = win32com.client.GetActiveObject("Visio.Application")
    except Exception:
        app = win32com.client.Dispatch("Visio.Application")
        owns = True
        try:
            app.Visible = False
        except Exception:
            pass
    problems = []
    tmp = tempfile.mkdtemp(prefix="visio_icons_")
    scratch = doc = None
    try:
        scratch = app.Documents.Add("")
        page = scratch.Pages.Item(1)
        doc = app.Documents.OpenEx(os.path.abspath(src), 2 + 64)  # read-only + hidden
        wanted = set(names)
        for i in range(1, doc.Masters.Count + 1):
            master = doc.Masters.Item(i)
            name = master.Name
            if name not in wanted:
                continue
            shp = None
            try:
                shp = page.Drop(master, 5, 5)
                if shp.OneD:  # connector / line, not an icon
                    problems.append(f"{name} (connector)")
                    continue
                w, h = shp.Cells("Width").Result("in"), shp.Cells("Height").Result("in")
                k = 4.0 / max(w, h, 0.01)  # bitmap export is ~96 dpi: enlarge to get enough pixels
                shp.Cells("Width").ResultIU = w * k
                shp.Cells("Height").ResultIU = h * k
                out = os.path.join(tmp, f"m{i}.png")
                shp.Export(out)
                with Image.open(out) as im:
                    emit(name, finish(clear_white_background(im.copy()), size))
            except Exception as e:
                problems.append(f"{name} ({e})")
            finally:
                if shp is not None:
                    try:
                        shp.Delete()
                    except Exception:
                        pass
    finally:
        for d in (doc, scratch):
            if d is not None:
                try:
                    d.Saved = True
                    d.Close()
                except Exception:
                    pass
        if owns:
            try:
                app.Quit()
            except Exception:
                pass
        shutil.rmtree(tmp, ignore_errors=True)
    return problems


def safe_name(name):
    name = re.sub(r'[\\/:*?"<>|\r\n\t%]+', " ", name).strip()
    return re.sub(r"\s+", " ", name) or "icon"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("inputs", nargs="+")
    ap.add_argument("--dest", default=str(Path(__file__).resolve().parent.parent / "icons" / "visio"))
    ap.add_argument("--size", type=int, default=512, help="Longest side in pixels (default 512)")
    ap.add_argument("--prefix", default="visio-", help="File name prefix (default 'visio-')")
    ap.add_argument("--only", help="Comma-separated master names to export (default: all)")
    ap.add_argument("--visio", choices=("auto", "off", "only"), default="auto",
                    help="auto: pictures from the zip, vector masters via Visio; off: zip only; "
                         "only: everything via Visio")
    ap.add_argument("--list", action="store_true", help="Only list what would be exported")
    args = ap.parse_args()

    files = []
    for item in args.inputs:
        p = Path(item)
        files += sorted(p.rglob("*")) if p.is_dir() else [p]
    files = [f for f in files if f.suffix.lower() in (".vsdx", ".vssx", ".vstx")]
    only = {s.strip() for s in args.only.split(",")} if args.only else None
    dest = Path(args.dest)
    if not args.list:
        dest.mkdir(parents=True, exist_ok=True)

    seen_hash, used = {}, set()
    counts = {"zip": 0, "visio": 0}
    notes = []

    def save(src_name, im, method):
        digest = hashlib.sha1(im.tobytes()).hexdigest()
        if digest in seen_hash:
            return
        base = args.prefix + safe_name(src_name)
        final, n = base, 1
        while final.lower() in used:
            n += 1
            final = f"{base}-{n}"
        used.add(final.lower())
        seen_hash[digest] = final
        im.save(dest / f"{final}.png")
        counts[method] += 1

    for f in files:
        try:
            z = zipfile.ZipFile(f)
        except zipfile.BadZipFile:
            notes.append(f"{f.name}: not a .vsdx/.vssx zip (old binary format?): open it in Visio and save as .vssx/.vsdx")
            continue
        todo = []
        for name, part in iter_masters(z):
            if only and name not in only:
                continue
            kind, info = classify(z, part)
            if kind == "image" and args.visio != "only":
                if args.list:
                    print(f"zip    {f.stem}: {name}")
                    continue
                try:
                    save(name, picture_to_png(z.read(info), posixpath.splitext(info)[1].lower(), args.size), "zip")
                except Exception as e:
                    notes.append(f"{f.stem}: {name}: {e}")
            else:
                todo.append(name)
        if not todo or args.visio == "off":
            if todo:
                notes.append(f"{f.stem}: {len(todo)} vector masters skipped (--visio off): " + ", ".join(todo[:8]))
            continue
        if args.list:
            for name in todo:
                print(f"visio  {f.stem}: {name}")
            continue
        try:
            problems = visio_export(f, todo, args.size, lambda n, im: save(n, im, "visio"))
            if problems:
                notes.append(f"{f.stem}: not exported: " + ", ".join(problems[:10]) + (" ..." if len(problems) > 10 else ""))
        except ImportError:
            notes.append(f"{f.stem}: {len(todo)} vector masters skipped: pywin32 not installed")
        except Exception as e:
            notes.append(f"{f.stem}: Visio export failed ({e}); {len(todo)} vector masters skipped")

    if not args.list:
        print(f"Exported {counts['zip']} from embedded pictures + {counts['visio']} via Visio to {dest}")
    for n in notes:
        print("NOTE", n)
    return 0


if __name__ == "__main__":
    sys.exit(main())
