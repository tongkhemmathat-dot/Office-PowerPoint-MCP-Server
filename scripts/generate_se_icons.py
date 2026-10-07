#!/usr/bin/env python
"""
Generate the built-in system-engineering icon set (servers, VMs, storage, network, users).

The icons are drawn from scratch with Pillow (original artwork, MIT like the rest of the repo),
in a flat/soft-gradient style similar to common infrastructure diagrams. Output: 512x512 PNGs
with transparent backgrounds in assets/icons/se/, named `se-<name>.png`.

Usage:
    python scripts/generate_se_icons.py [--dest assets/icons/se]
"""
import argparse
import math
from pathlib import Path

from PIL import Image, ImageDraw

SIZE = 512
S = 4  # supersampling factor


# ---------- drawing helpers (coordinates are in 512-space) ----------
def new_canvas():
    return Image.new("RGBA", (SIZE * S, SIZE * S), (0, 0, 0, 0))


def sc(v):
    return int(round(v * S))


def box_sc(b):
    return tuple(sc(v) for v in b)


def vgrad(w, h, c1, c2):
    img = Image.new("RGBA", (w, h))
    d = ImageDraw.Draw(img)
    for y in range(h):
        t = y / max(h - 1, 1)
        d.line([(0, y), (w, y)], fill=tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(4)))
    return img


def rgba(c, a=255):
    return (c[0], c[1], c[2], a)


def grad_rrect(im, box, radius, c1, c2, outline=None, ow=3):
    x0, y0, x1, y1 = box_sc(box)
    w, h = x1 - x0, y1 - y0
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, w - 1, h - 1), sc(radius), fill=255)
    im.paste(vgrad(w, h, rgba(c1), rgba(c2)), (x0, y0), mask)
    if outline:
        ImageDraw.Draw(im).rounded_rectangle((x0, y0, x1, y1), sc(radius), outline=rgba(outline), width=sc(ow))


def poly(im, pts, fill, outline=None, ow=2):
    d = ImageDraw.Draw(im)
    p = [(sc(x), sc(y)) for x, y in pts]
    d.polygon(p, fill=rgba(fill))
    if outline:
        d.line(p + [p[0]], fill=rgba(outline), width=sc(ow), joint="curve")


def ellipse(im, box, fill=None, outline=None, ow=2):
    ImageDraw.Draw(im).ellipse(box_sc(box), fill=rgba(fill) if fill else None,
                               outline=rgba(outline) if outline else None, width=sc(ow))


def line(im, pts, color, w):
    ImageDraw.Draw(im).line([(sc(x), sc(y)) for x, y in pts], fill=rgba(color), width=sc(w), joint="curve")


def shadow(im, box):
    x0, y0, x1, y1 = box
    ellipse(im, (x0, y1 - 6, x1, y1 + 14), fill=(0, 0, 0, 40))


def finish(im):
    return im.resize((SIZE, SIZE), Image.LANCZOS)


GREY1, GREY2, GREY_LINE = (165, 170, 176), (112, 118, 125), (80, 85, 92)
BLUE1, BLUE2, BLUE_LINE = (120, 175, 230), (40, 100, 170), (25, 70, 125)
ORANGE1, ORANGE2, ORANGE_LINE = (250, 170, 90), (225, 110, 25), (170, 75, 10)
GREEN1, GREEN2 = (150, 205, 100), (90, 150, 50)


# ---------- reusable parts ----------
def server_unit(im, x, y, w, h):
    shadow(im, (x + 10, y, x + w - 10, y + h))
    grad_rrect(im, (x, y, x + w, y + h), h * 0.14, GREY1, GREY2, outline=GREY_LINE, ow=3)
    r = h * 0.11
    for i in range(3):
        cx = x + h * 0.32 + i * h * 0.3
        ellipse(im, (cx - r, y + h / 2 - r, cx + r, y + h / 2 + r), fill=(250, 250, 250))
    grad_rrect(im, (x + w - w * 0.3, y + h * 0.38, x + w - h * 0.2, y + h * 0.62), h * 0.12,
               (250, 250, 250), (225, 228, 232))


def cylinder(im, cx, top, w, h, c_body=(150, 156, 162), c_body2=(105, 110, 116), c_top=GREEN1, c_top2=GREEN2):
    x0, x1 = cx - w / 2, cx + w / 2
    eh = w * 0.28
    grad_rrect(im, (x0, top + eh / 2, x1, top + h - eh / 2), 2, c_body, c_body2)
    ellipse(im, (x0, top + h - eh, x1, top + h), fill=c_body2)
    ellipse(im, (x0, top, x1, top + eh), fill=c_top2, outline=None)
    ellipse(im, (x0 + 4, top + 3, x1 - 4, top + eh - 3), fill=c_top)


def iso_cube(im, cx, cy, r, top=(255, 190, 110), left=(235, 130, 40), right=(190, 90, 15), edge=ORANGE_LINE):
    h = r * 0.5
    t = [(cx, cy - r), (cx + r * 0.87, cy - h), (cx, cy), (cx - r * 0.87, cy - h)]
    l = [(cx - r * 0.87, cy - h), (cx, cy), (cx, cy + r), (cx - r * 0.87, cy + h)]
    rr = [(cx + r * 0.87, cy - h), (cx + r * 0.87, cy + h), (cx, cy + r), (cx, cy)]
    poly(im, l, left, edge, 2)
    poly(im, rr, right, edge, 2)
    poly(im, t, top, edge, 2)


def layers(im, cx, cy, w, color=ORANGE1, color2=ORANGE2):
    for i, c in enumerate((color2, color2, color)):
        y = cy + (1 - i) * w * 0.28
        poly(im, [(cx, y - w * 0.3), (cx + w / 2, y), (cx, y + w * 0.3), (cx - w / 2, y)], c, (255, 255, 255), 2)


def globe(im, cx, cy, r):
    x0, y0, x1, y1 = cx - r, cy - r, cx + r, cy + r
    ellipse(im, (x0, y0, x1, y1), fill=BLUE2, outline=BLUE_LINE, ow=3)
    ellipse(im, (x0 + 5, y0 + 5, x1 - 5, y1 - 5), fill=(80, 150, 215))
    for k in (0.35, 0.7):
        ellipse(im, (cx - r * k, y0 + 4, cx + r * k, y1 - 4), outline=(255, 255, 255), ow=2)
    line(im, [(x0 + 6, cy), (x1 - 6, cy)], (255, 255, 255), 2)
    line(im, [(cx, y0 + 4), (cx, y1 - 4)], (255, 255, 255), 2)


def person(im, cx, base, scale=1.0, color=(115, 120, 128), color2=(80, 85, 92)):
    r = 40 * scale
    ellipse(im, (cx - r, base - 190 * scale, cx + r, base - 190 * scale + 2 * r), fill=color)
    grad_rrect(im, (cx - 70 * scale, base - 135 * scale, cx + 70 * scale, base), 55 * scale, color, color2)


def arrow_poly(p0, p1, shaft, head):
    dx, dy = p1[0] - p0[0], p1[1] - p0[1]
    n = math.hypot(dx, dy)
    ux, uy = dx / n, dy / n
    px, py = -uy, ux
    hb = (p1[0] - ux * head, p1[1] - uy * head)
    return [(p0[0] + px * shaft, p0[1] + py * shaft), (hb[0] + px * shaft, hb[1] + py * shaft),
            (hb[0] + px * head * 0.8, hb[1] + py * head * 0.8), p1,
            (hb[0] - px * head * 0.8, hb[1] - py * head * 0.8), (hb[0] - px * shaft, hb[1] - py * shaft),
            (p0[0] - px * shaft, p0[1] - py * shaft)]


def vm_card(im, x, y, w, h, tile=True):
    grad_rrect(im, (x, y, x + w, y + h), h * 0.1, (250, 252, 255), (215, 225, 238), outline=BLUE_LINE, ow=3)
    grad_rrect(im, (x, y, x + w, y + h * 0.26), h * 0.1, BLUE1, BLUE2)
    if tile:
        gw, gh = w * 0.34, h * 0.26
        for i in range(2):
            for j in range(2):
                tx = x + w * 0.12 + i * (gw + w * 0.08)
                ty = y + h * 0.36 + j * (gh + h * 0.07)
                grad_rrect(im, (tx, ty, tx + gw, ty + gh), 4, ORANGE1 if (i + j) % 2 == 0 else BLUE1,
                           ORANGE2 if (i + j) % 2 == 0 else BLUE2)


# ---------- icons ----------
def icon_server():
    im = new_canvas()
    server_unit(im, 36, 190, 440, 132)
    return im


def icon_server_rack():
    im = new_canvas()
    for i in range(3):
        server_unit(im, 56, 96 + i * 108, 400, 92)
    return im


def icon_app_server():
    im = new_canvas()
    server_unit(im, 30, 220, 380, 120)
    iso_cube(im, 400, 205, 70)
    return im


def icon_web_server():
    im = new_canvas()
    server_unit(im, 30, 220, 380, 120)
    globe(im, 395, 215, 62)
    return im


def icon_gis_server():
    im = new_canvas()
    server_unit(im, 30, 220, 380, 120)
    layers(im, 395, 200, 130)
    return im


def icon_db_server():
    im = new_canvas()
    server_unit(im, 30, 230, 360, 118)
    cylinder(im, 405, 150, 100, 150)
    return im


def icon_database():
    im = new_canvas()
    cylinder(im, 256, 90, 230, 330, c_body=(155, 162, 170), c_body2=(105, 112, 120))
    return im


def icon_nas():
    im = new_canvas()
    shadow(im, (40, 250, 370, 380))
    grad_rrect(im, (36, 190, 380, 380), 22, (90, 130, 185), (35, 70, 120), outline=BLUE_LINE, ow=3)
    for i in range(3):
        grad_rrect(im, (62, 214 + i * 54, 330, 246 + i * 54), 8, (235, 242, 250), (190, 205, 225))
        ellipse(im, (344, 224 + i * 54, 362, 242 + i * 54), fill=(120, 230, 120) if i < 2 else (240, 190, 70))
    cylinder(im, 420, 150, 90, 130)
    return im


def icon_vm():
    im = new_canvas()
    vm_card(im, 96, 96, 320, 320)
    return im


def icon_esx_host():
    im = new_canvas()
    for i in range(3):
        vm_card(im, 48 + i * 152, 40, 136, 128)
    poly(im, [(60, 190), (452, 190), (352, 290), (160, 290)], (170, 205, 240), BLUE_LINE, 3)
    server_unit(im, 96, 320, 320, 100)
    return im


def icon_switch():
    im = new_canvas()
    top = [(40, 235), (256, 140), (472, 235), (256, 330)]
    left = [(40, 235), (256, 330), (256, 410), (40, 315)]
    right = [(256, 330), (472, 235), (472, 315), (256, 410)]
    poly(im, left, (45, 105, 175), BLUE_LINE, 3)
    poly(im, right, (28, 78, 148), BLUE_LINE, 3)
    poly(im, top, (125, 180, 235), BLUE_LINE, 3)
    poly(im, arrow_poly((170, 262), (300, 205), 9, 34), (255, 255, 255))
    poly(im, arrow_poly((342, 262), (212, 205), 9, 34), (255, 255, 255))
    return im


def icon_firewall():
    wall = Image.new("RGBA", (300 * S, 360 * S), (0, 0, 0, 0))
    d = ImageDraw.Draw(wall)
    w, h = 300 * S, 360 * S
    d.rectangle((0, 0, w, h), fill=(150, 40, 20, 255))
    rows, bh = 6, h // 6
    for r in range(rows):
        y0 = r * bh
        off = (r % 2) * (w // 6)
        n = 3
        bw = w // n
        for c in range(-1, n + 1):
            x0 = c * bw + off
            col = (235, 105, 55, 255) if (r + c) % 2 == 0 else (225, 85, 40, 255)
            d.rectangle((x0 + 5 * S, y0 + 5 * S, x0 + bw - 5 * S, y0 + bh - 5 * S), fill=col)
            d.rectangle((x0 + 5 * S, y0 + 5 * S, x0 + bw - 5 * S, y0 + 16 * S), fill=(250, 150, 90, 255))
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rectangle((0, 0, w - 1, h - 1), fill=255)
    wall.putalpha(mask)
    d.rectangle((0, 0, w - 1, h - 1), outline=(110, 25, 10, 255), width=4 * S)
    shear = 0.28
    new_w = int(w + h * shear)
    sheared = wall.transform((new_w, h), Image.AFFINE, (1, shear, -h * shear, 0, 1, 0), Image.BICUBIC)
    im = new_canvas()
    shadow_im = ImageDraw.Draw(im)
    shadow_im.ellipse(box_sc((90, 440, 430, 470)), fill=(0, 0, 0, 45))
    im.paste(sheared, ((SIZE * S - new_w) // 2, 60 * S), sheared)
    return im


def icon_internet():
    im = new_canvas()
    m = Image.new("L", im.size, 0)
    d = ImageDraw.Draw(m)
    for (x0, y0, x1, y1) in ((60, 250, 220, 410), (130, 150, 330, 350), (250, 190, 450, 400),
                             (330, 260, 470, 400), (110, 330, 430, 410)):
        d.ellipse(box_sc((x0, y0, x1, y1)), fill=255)
    d.rectangle(box_sc((130, 340, 400, 400)), fill=255)
    g = vgrad(im.size[0], im.size[1], (205, 232, 252, 255), (70, 150, 225, 255))
    im.paste(g, (0, 0), m)
    return im


def icon_user():
    im = new_canvas()
    person(im, 256, 440, 1.3)
    return im


def icon_user_web():
    im = new_canvas()
    person(im, 170, 440, 1.15)
    grad_rrect(im, (250, 200, 480, 350), 14, (120, 190, 120), (60, 140, 60), outline=(40, 100, 40), ow=3)
    grad_rrect(im, (264, 214, 466, 336), 8, (245, 250, 255), (215, 230, 245))
    globe(im, 365, 275, 42)
    grad_rrect(im, (340, 350, 390, 380), 4, (130, 130, 135), (95, 95, 100))
    grad_rrect(im, (300, 378, 430, 396), 8, (130, 130, 135), (95, 95, 100))
    return im


def icon_user_mobile():
    im = new_canvas()
    person(im, 170, 440, 1.15)
    grad_rrect(im, (300, 170, 440, 400), 24, ORANGE1, ORANGE2, outline=ORANGE_LINE, ow=3)
    grad_rrect(im, (314, 196, 426, 372), 10, (252, 252, 255), (225, 232, 240))
    ellipse(im, (358, 380, 382, 396), fill=(250, 250, 250))
    return im


def icon_workstation():
    im = new_canvas()
    grad_rrect(im, (40, 120, 330, 330), 18, ORANGE1, ORANGE2, outline=ORANGE_LINE, ow=3)
    grad_rrect(im, (58, 138, 312, 312), 8, (250, 252, 255), (220, 228, 240))
    grad_rrect(im, (160, 330, 210, 372), 4, (140, 140, 145), (100, 100, 105))
    grad_rrect(im, (110, 370, 260, 392), 10, (140, 140, 145), (100, 100, 105))
    grad_rrect(im, (360, 150, 470, 392), 14, GREY1, GREY2, outline=GREY_LINE, ow=3)
    grad_rrect(im, (376, 178, 454, 198), 5, (235, 235, 238), (200, 200, 205))
    ellipse(im, (404, 340, 428, 364), fill=(235, 235, 238))
    return im


def icon_container():
    im = new_canvas()
    shadow(im, (60, 150, 452, 380))
    grad_rrect(im, (56, 140, 456, 380), 18, (90, 160, 225), (30, 85, 155), outline=BLUE_LINE, ow=4)
    for i in range(11):
        x = 92 + i * 33
        line(im, [(x, 160), (x, 360)], (255, 255, 255, 70)[:3], 3)
    grad_rrect(im, (56, 140, 456, 164), 12, (150, 195, 240), (80, 140, 205))
    grad_rrect(im, (56, 356, 456, 380), 12, (35, 90, 160), (20, 60, 115))
    return im


def icon_container_image():
    im = new_canvas()
    shadow(im, (60, 380, 452, 400))
    for i, (c1, c2) in enumerate(((BLUE2, BLUE_LINE), (BLUE2, BLUE_LINE), (BLUE1, BLUE2))):
        y = 330 - i * 80
        poly(im, [(256, y + 40), (456, y), (256, y - 40), (56, y)], c2, (255, 255, 255), 3)
        poly(im, [(56, y), (256, y + 40), (256, y + 68), (56, y + 28)], c1, (255, 255, 255), 3)
        poly(im, [(256, y + 40), (456, y), (456, y + 28), (256, y + 68)], c2, (255, 255, 255), 3)
        poly(im, [(256, y + 40), (456, y), (256, y - 40), (56, y)], c1 if i == 2 else c2, (255, 255, 255), 3)
    return im


def icon_registry():
    im = icon_internet()
    iso_cube(im, 256, 300, 78, top=(150, 195, 240), left=(60, 120, 190), right=(30, 85, 150), edge=BLUE_LINE)
    return im


ICONS = {
    "container": icon_container, "container-image": icon_container_image, "registry": icon_registry,
    "server": icon_server, "server-rack": icon_server_rack, "app-server": icon_app_server,
    "web-server": icon_web_server, "gis-server": icon_gis_server, "db-server": icon_db_server,
    "database": icon_database, "nas": icon_nas, "vm": icon_vm, "esx-host": icon_esx_host,
    "network-switch": icon_switch, "firewall": icon_firewall, "internet": icon_internet,
    "user": icon_user, "user-web": icon_user_web, "user-mobile": icon_user_mobile,
    "workstation": icon_workstation,
}


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    root = Path(__file__).resolve().parent.parent
    parser.add_argument("--dest", default=str(root / "assets" / "icons" / "se"))
    dest = Path(parser.parse_args().dest)
    dest.mkdir(parents=True, exist_ok=True)
    for name, fn in ICONS.items():
        finish(fn()).save(dest / f"se-{name}.png")
    print(f"Wrote {len(ICONS)} icons to {dest}")


if __name__ == "__main__":
    main()
