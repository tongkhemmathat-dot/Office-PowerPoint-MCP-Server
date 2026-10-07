#!/usr/bin/env python
"""
Build the system-engineering PowerPoint template (16:9 master + layouts, logo placeholder, footer
with date and page number, theme colours and fonts for Latin AND Thai).

Usage:
    python scripts/build_se_template.py [--out templates/SE_Template.pptx] [--font Tahoma] [--scale 1.0]

Examples:
    # default (Tahoma)
    python scripts/build_se_template.py
    # TH Sarabun New: narrower than Tahoma (~0.67 x width, ~0.62 x x-height at the same pt), so
    # scale text sizes up (1.3 keeps layouts about as wide as the Tahoma version)
    python scripts/build_se_template.py --font "TH Sarabun New" --scale 1.3 --out templates/SE_Template_Sarabun.pptx

--font sets the theme's Latin, complex-script (cs, which PowerPoint uses for Thai) and Thai fonts.
--scale multiplies the master/layout text sizes (title, body, footer, title/section slides).
"""
import argparse, sys, uuid, os
sys.path.insert(0, os.getcwd())
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.dml import MSO_LINE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.shapes.shapetree import SlideShapes
from lxml import etree
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import utils.presentation_utils as pu

_ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
_ap.add_argument("--out", default=os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "templates", "SE_Template.pptx"))
_ap.add_argument("--font", default="Tahoma")
_ap.add_argument("--scale", type=float, default=1.0)
_args = _ap.parse_args()
OUT, FONT, SC = _args.out, _args.font, _args.scale
DK, ACC, LIGHT, GREY = "1F3A5F", "0B5CAD", "EAF1FA", "8A94A3"
rgb = lambda h: RGBColor.from_string(h)
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

p = pu.create_presentation(slide_size="16:9")
master = p.slide_master

# ---------- theme: colours + fonts (Latin + Thai) ----------
tp = master.part.part_related_by(RT.THEME)
th = etree.fromstring(tp.blob)
th.set("name", "SE Template")
clr = th.find(f".//{{{A}}}clrScheme"); clr.set("name", "SE Template")
for tag, val in dict(dk2=DK, lt2=LIGHT, accent1=ACC, accent2="ED7014", accent3="2E9E5B",
                     accent4="7B4EA3", accent5="17A2B8", accent6="C0392B").items():
    el = clr.find(f"{{{A}}}{tag}")
    for c in list(el):
        el.remove(c)
    etree.SubElement(el, f"{{{A}}}srgbClr").set("val", val)
fs = th.find(f".//{{{A}}}fontScheme"); fs.set("name", "SE Template")
for grp in (fs.find(f"{{{A}}}majorFont"), fs.find(f"{{{A}}}minorFont")):
    grp.find(f"{{{A}}}latin").set("typeface", FONT)
    grp.find(f"{{{A}}}cs").set("typeface", FONT)  # Thai is a complex script: PowerPoint uses <a:cs>
    thai = [f for f in grp.findall(f"{{{A}}}font") if f.get("script") == "Thai"]
    if thai:
        thai[0].set("typeface", FONT)
    else:
        etree.SubElement(grp, f"{{{A}}}font", script="Thai", typeface=FONT)
tp._blob = etree.tostring(th, xml_declaration=True, encoding="UTF-8", standalone=True)


# ---------- helpers ----------
def shapes_of(container):  # add_* on master/layout spTree
    return SlideShapes(container._element.cSld.spTree, container)


def to_back(shape):
    tree = shape._element.getparent()
    tree.remove(shape._element)
    tree.insert(2, shape._element)


def rect(cont, l, t, w, h, fill=None, line=None, lw=1, dash=False, name=None, back=False):
    s = shapes_of(cont).add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(fill)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(lw)
    else:
        s.line.fill.background()
    if dash:
        s.line.dash_style = MSO_LINE.DASH
    s.shadow.inherit = False
    if name:
        s.name = name
    if back:
        to_back(s)
    return s


def label(cont, text, l, t, w, h, size=None, color="FFFFFF", bold=False, align=PP_ALIGN.LEFT, name=None):
    size = size or int(round(10 * SC))
    s = shapes_of(cont).add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = s.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = True
    para = tf.paragraphs[0]
    para.alignment = align
    r = para.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = rgb(color)
    if name:
        s.name = name
    return s


def field(cont, kind, l, t, w, h, align, name):
    s = label(cont, "x", l, t, w, h, align=align, name=name)
    para = s.text_frame.paragraphs[0]._p
    for r in para.findall(qn("a:r")):
        para.remove(r)
    fld = etree.SubElement(para, qn("a:fld"))
    fld.set("id", "{%s}" % str(uuid.uuid4()).upper())
    fld.set("type", kind)
    rpr = etree.SubElement(fld, qn("a:rPr"))
    rpr.set("lang", "en-US")
    rpr.set("sz", str(int(round(1000 * SC))))
    sf = etree.SubElement(rpr, qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr")).set("val", "FFFFFF")
    etree.SubElement(fld, qn("a:t")).text = "#" if kind == "slidenum" else "date"
    return s


def set_ph(ph, l=None, t=None, w=None, h=None, size=None, bold=None, color=None, align=None,
           anchor=None, nobullet=False):
    if l is not None:
        ph.left, ph.top, ph.width, ph.height = Inches(l), Inches(t), Inches(w), Inches(h)
    if anchor is not None:
        ph.text_frame.vertical_anchor = anchor
    if size or bold is not None or color or align or nobullet:
        ls = ph._element.txBody.find(qn("a:lstStyle"))
        for c in list(ls):
            ls.remove(c)
        lvl = etree.SubElement(ls, qn("a:lvl1pPr"))
        if align:
            lvl.set("algn", align)
        if nobullet:
            lvl.set("marL", "0")
            lvl.set("indent", "0")
            etree.SubElement(lvl, qn("a:buNone"))
        d = etree.SubElement(lvl, qn("a:defRPr"))
        if size:
            d.set("sz", str(int(size * 100)))
        if bold is not None:
            d.set("b", "1" if bold else "0")
        if color:
            sf = etree.SubElement(d, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr")).set("val", color)


def inherit_pos(ph):
    x = ph._element.spPr.find(qn("a:xfrm"))
    if x is not None:
        ph._element.spPr.remove(x)


def drop_footer_phs(cont):
    for sh in list(cont.placeholders):
        if sh.placeholder_format.type in (PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER):
            sh._element.getparent().remove(sh._element)


# ---------- master ----------
drop_footer_phs(master)
for ph in master.placeholders:
    t = ph.placeholder_format.type
    if t == PP_PLACEHOLDER.TITLE:
        set_ph(ph, 0.5, 0.25, 10.9, 0.8, anchor=MSO_ANCHOR.MIDDLE)
    elif t == PP_PLACEHOLDER.BODY:
        set_ph(ph, 0.5, 1.3, 12.33, 5.6)
ts = master._element.find(qn("p:txStyles"))
l1 = ts.find(qn("p:titleStyle")).find(qn("a:lvl1pPr"))
l1.set("algn", "l")
d = l1.find(qn("a:defRPr"))
d.set("sz", str(int(round(2600 * SC))))
d.set("b", "1")
for c in d.findall(qn("a:solidFill")):
    d.remove(c)
sf = etree.Element(qn("a:solidFill"))
etree.SubElement(sf, qn("a:srgbClr")).set("val", DK)
d.insert(0, sf)
for i, sz in enumerate((2000, 1800, 1600, 1400, 1400), start=1):
    ts.find(qn("p:bodyStyle")).find(qn(f"a:lvl{i}pPr")).find(qn("a:defRPr")).set("sz", str(int(round(sz * SC))))

rect(master, 0.5, 1.1, 12.33, 0.04, fill=ACC, name="Title Rule")
rect(master, 0, 7.15, 13.333, 0.35, fill=DK, name="Footer Bar")
field(master, "datetime1", 0.4, 7.17, 2.5, 0.3, PP_ALIGN.LEFT, "Date")
label(master, "ชื่อโครงการ / Project Name", 3.2, 7.17, 6.9, 0.3, align=PP_ALIGN.CENTER, name="Footer Text")
field(master, "slidenum", 11.9, 7.17, 1.0, 0.3, PP_ALIGN.RIGHT, "Slide Number")
logo = rect(master, 11.6, 0.2, 1.3, 0.75, line=GREY, lw=1, dash=True, name="Logo Placeholder")
logo.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
logo.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
r = logo.text_frame.paragraphs[0].add_run()
r.text = "LOGO"
r.font.size = Pt(10)
r.font.color.rgb = rgb(GREY)

# ---------- layouts ----------
for lay in list(master.slide_layouts)[7:]:  # keep indexes 0-6 identical to the default template
    master.slide_layouts.remove(lay)
lays = list(master.slide_layouts)
for lay in lays:
    drop_footer_phs(lay)


def dark_layout(lay, title_box, sub_box, tsize):
    lay._element.set("showMasterSp", "0")
    rect(lay, 0, 0, 13.333, 7.5, fill=DK, name="Background", back=True)
    rect(lay, 0, 0, 0.35, 7.5, fill=ACC, name="Accent Band")
    lg = rect(lay, 11.6, 0.3, 1.3, 0.75, line="FFFFFF", lw=1, dash=True, name="Logo Placeholder")
    lg.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    lg.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    rr = lg.text_frame.paragraphs[0].add_run()
    rr.text = "LOGO"
    rr.font.size = Pt(10)
    rr.font.color.rgb = rgb("FFFFFF")
    for ph in lay.placeholders:
        if ph.placeholder_format.idx == 0:
            set_ph(ph, *title_box, size=tsize, bold=True, color="FFFFFF", align="l", anchor=MSO_ANCHOR.BOTTOM)
        else:
            set_ph(ph, *sub_box, size=int(round(20 * SC)), bold=False, color="DCE6F2", align="l", anchor=MSO_ANCHOR.TOP,
                   nobullet=True)


dark_layout(lays[0], (0.9, 2.2, 10.5, 1.8), (0.9, 4.15, 10.5, 1.2), int(round(40 * SC)))  # Title Slide
dark_layout(lays[2], (0.9, 2.6, 10.5, 1.3), (0.9, 4.05, 10.5, 1.0), int(round(36 * SC)))  # Section Header

for i in (1, 5):  # inherit master title/body positions
    for ph in lays[i].placeholders:
        inherit_pos(ph)

for ph in lays[3].placeholders:  # Two Content
    k = ph.placeholder_format.idx
    if k == 0:
        inherit_pos(ph)
    elif k == 1:
        set_ph(ph, 0.5, 1.3, 6.0, 5.6)
    elif k == 2:
        set_ph(ph, 6.83, 1.3, 6.0, 5.6)

for ph in lays[4].placeholders:  # Comparison
    k = ph.placeholder_format.idx
    if k == 0:
        inherit_pos(ph)
    elif k == 1:
        set_ph(ph, 0.5, 1.3, 6.0, 0.5, size=int(round(18 * SC)), bold=True, color=ACC, anchor=MSO_ANCHOR.MIDDLE)
    elif k == 2:
        set_ph(ph, 0.5, 1.85, 6.0, 5.05)
    elif k == 3:
        set_ph(ph, 6.83, 1.3, 6.0, 0.5, size=int(round(18 * SC)), bold=True, color=ACC, anchor=MSO_ANCHOR.MIDDLE)
    elif k == 4:
        set_ph(ph, 6.83, 1.85, 6.0, 5.05)

# diagram/table canvas: frame on "Title Only"
rect(lays[5], 0.4, 1.3, 12.53, 5.7, line="C9D3E0", lw=1, name="Content Frame")
lays[5]._element.cSld.set("name", "Title Only (Diagram / Table)")

def renumber_ids(cont):
    """python-pptx numbers new shapes from the largest id in the whole part (incl. layout ids
    near 2^31); PowerPoint rejects that, so renumber shape ids sequentially."""
    tree = cont._element.cSld.spTree
    for n, el in enumerate(tree.iter(qn("p:cNvPr")), start=1):
        el.set("id", str(n))


for cont in [master] + list(master.slide_layouts):
    renumber_ids(cont)

p.save(OUT)
print("saved", OUT)
for i, lay in enumerate(p.slide_master.slide_layouts):
    print(i, lay.name, [(ph.placeholder_format.idx, str(ph.placeholder_format.type).split('.')[-1].split(' ')[0])
                        for ph in lay.placeholders])
