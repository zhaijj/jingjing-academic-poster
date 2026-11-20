#!/usr/bin/env python3
"""Build a conference poster (.pptx + .pdf + preview .png) from a YAML spec.

House style (distilled from Jingjing Zhai's 2023-2026 posters):
  Arial everywhere; panels = light-blue (#DEEBF7) header bar + white body,
  both outlined in gray (#AFABAB, ~6 pt on a 45-in poster); green-free,
  black text; check-mark bullets; small right-aligned citations under figures.

Usage:
  python build_poster.py spec.yaml -o out_dir [--no-pdf]

See ../references/spec.md for the full spec format.
"""
import argparse, os, re, subprocess, sys, shutil, tempfile, json

import yaml
from PIL import Image, ImageChops, ImageFont
from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOGO_DIR = os.path.join(SKILL_DIR, "assets", "logos")
EMU_PER_PT = 12700
EMU_PER_IN = 914400
LH = 1.2  # line height factor used for measuring (PowerPoint single spacing for Arial ~1.15)

COLORS = {
    "header_fill": "DEEBF7", "border": "AFABAB", "box": "E7E6E6",
    "r": "C00000", "red": "C00000", "b": "2E75B6", "blue": "2E75B6",
    "g": "38761D", "green": "38761D", "o": "C55A11", "orange": "C55A11",
    "gray": "595959", "purple": "7030A0",
    # callout fills
    "yellow_box": "FFE599", "green_box": "B6D7A8", "blue_box": "C9DAF8",
    "gray_box": "F2F2F2", "orange_box": "F8CBAD", "pink_box": "F4CCCC",
}

PRESETS = {  # sizes in inches, fonts in pt — measured from the historical posters
    "landscape": dict(size=(45, 40), title=96, authors=44, aff=28, header=72, sub=54, body=44, caption=28),
    "portrait":  dict(size=(33, 45), title=66, authors=36, aff=23, header=60, sub=44, body=38, caption=24),
    "square":    dict(size=(42, 42), title=80, authors=48, aff=36, header=70, sub=48, body=46, caption=28),
}

WARN = []


def warn(msg):
    WARN.append(msg)
    print("WARNING:", msg, file=sys.stderr)


# ----------------------------------------------------------------------------- fonts / measuring
def _find_font(names):
    dirs = ["/usr/share/fonts/truetype/liberation", "/usr/share/fonts/liberation",
            "/Library/Fonts", "/System/Library/Fonts/Supplemental", "C:/Windows/Fonts",
            os.path.expanduser("~/Library/Fonts")]
    for d in dirs:
        for n in names:
            p = os.path.join(d, n)
            if os.path.exists(p):
                return p
    raise SystemExit("No Arial-metric font found (install fonts-liberation)")


FONT_FILES = {
    (False, False): _find_font(["Arial.ttf", "arial.ttf", "LiberationSans-Regular.ttf"]),
    (True, False): _find_font(["Arial Bold.ttf", "arialbd.ttf", "LiberationSans-Bold.ttf"]),
    (False, True): _find_font(["Arial Italic.ttf", "ariali.ttf", "LiberationSans-Italic.ttf"]),
    (True, True): _find_font(["Arial Bold Italic.ttf", "arialbi.ttf", "LiberationSans-BoldItalic.ttf"]),
}
_font_cache = {}


def _font(bold, italic, size_pt):
    key = (bold, italic, round(size_pt * 10))
    if key not in _font_cache:
        _font_cache[key] = ImageFont.truetype(FONT_FILES[(bold, italic)], key[2])
    return _font_cache[key]


def text_width_pt(text, bold, italic, size_pt):
    return _font(bold, italic, size_pt).getlength(text) / 10.0


# ----------------------------------------------------------------------------- inline markup
# **bold**  *italic*  ^superscript^  ~subscript~  {r:red text}  {b:...} {g:...} {o:...} {gray:...}
TOKEN = re.compile(r"\\\*|\*\*|\*|\^|~|\{(\w+):|\}")


def parse_markup(s, base_bold=False):
    s = str(s)
    runs, buf = [], []
    st = dict(bold=base_bold, italic=False, sup=False, sub=False)
    colors = []

    def flush():
        if buf:
            runs.append(dict(text="".join(buf), color=colors[-1] if colors else None, **st))
            buf.clear()

    pos = 0
    for m in TOKEN.finditer(s):
        buf.append(s[pos:m.start()])
        tok = m.group(0)
        if tok == "\\*":
            buf.append("*")
        else:
            flush()
            if tok == "**":
                st["bold"] = not st["bold"]
            elif tok == "*":
                st["italic"] = not st["italic"]
            elif tok == "^":
                st["sup"] = not st["sup"]
            elif tok == "~":
                st["sub"] = not st["sub"]
            elif tok == "}":
                if colors:
                    colors.pop()
                else:
                    buf.append("}")
            else:
                colors.append(COLORS.get(m.group(1), m.group(1)))
        pos = m.end()
    buf.append(s[pos:])
    flush()
    return [r for r in runs if r["text"]]


def count_lines(runs, size_pt, max_w_pt):
    """Greedy word-wrap over styled runs; returns number of lines."""
    if max_w_pt <= 0:
        return 99
    # split into glued words separated by spaces
    words, cur = [], 0.0
    pending_space = 0.0
    lines, line_w = 1, 0.0
    for r in runs:
        sz = size_pt * (0.65 if (r["sup"] or r["sub"]) else 1.0)
        parts = re.split(r"( +|\n)", r["text"])
        for part in parts:
            if part == "":
                continue
            if part == "\n":
                lines += 1
                line_w, pending_space, cur = 0.0, 0.0, 0.0
                continue
            if part.strip() == "":
                pending_space += text_width_pt(part, r["bold"], r["italic"], sz)
                cur = 0.0
                continue
            w = text_width_pt(part, r["bold"], r["italic"], sz)
            if cur > 0:  # glued to previous fragment (no space between)
                line_w += w
                cur += w
                if line_w > max_w_pt and cur < max_w_pt:
                    lines += 1
                    line_w = cur
                continue
            if line_w > 0 and line_w + pending_space + w > max_w_pt:
                lines += 1
                line_w = w
            else:
                line_w += (pending_space if line_w > 0 else 0) + w
            pending_space = 0.0
            cur = w
    return lines


# ----------------------------------------------------------------------------- pptx helpers
def rgb(hexs):
    return RGBColor.from_string(hexs.upper())


def emu(pt):
    return Emu(int(round(pt * EMU_PER_PT)))


def add_runs(paragraph, runs, size_pt, default_color="000000"):
    for r in runs:
        chunks = r["text"].split("\n")
        for ci, chunk in enumerate(chunks):
            if ci:
                paragraph.add_line_break()
            if chunk:
                _add_run(paragraph, dict(r, text=chunk), size_pt, default_color)


def _add_run(paragraph, r, size_pt, default_color):
    run = paragraph.add_run()
    run.text = r["text"]
    f = run.font
    f.name = "Arial"
    f.size = Pt(size_pt)
    f.bold = r["bold"]
    f.italic = r["italic"]
    f.color.rgb = rgb(r["color"] or default_color)
    if r["sup"] or r["sub"]:
        run._r.get_or_add_rPr().set("baseline", "30000" if r["sup"] else "-25000")


BULLET_CHARS = {"check": ("Wingdings", "\u00fc"), "diamond": ("Wingdings", "v"),
                "arrow": ("Wingdings", "\u00d8"), "dot": ("Arial", "\u2022")}


def set_bullet(paragraph, kind, size_pt):
    font, char = BULLET_CHARS[kind]
    indent = int(size_pt * 1.15 * EMU_PER_PT)
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(indent))
    pPr.set("indent", str(-indent))
    for tag in ("a:buNone", "a:buFont", "a:buChar"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    bf = etree.SubElement(pPr, qn("a:buFont"))
    bf.set("typeface", font)
    bc = etree.SubElement(pPr, qn("a:buChar"))
    bc.set("char", char)


def textbox(slide, x, y, w, h, paras, anchor="top"):
    """paras: list of dict(runs, size, align, bullet, space_before)"""
    tb = slide.shapes.add_textbox(emu(x), emu(y), emu(w), emu(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, side, 0)
    tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
                          "bottom": MSO_ANCHOR.BOTTOM}[anchor]
    for i, pd in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                       "right": PP_ALIGN.RIGHT}[pd.get("align", "left")]
        if pd.get("space_before"):
            p.space_before = Pt(pd["space_before"])
        if pd.get("bullet"):
            set_bullet(p, pd["bullet"], pd["size"])
        add_runs(p, pd["runs"], pd["size"], pd.get("color", "000000"))
    return tb


def rect(slide, x, y, w, h, fill=None, line=None, line_w=0, shape=MSO_SHAPE.RECTANGLE, radius=None):
    s = slide.shapes.add_shape(shape, emu(x), emu(y), emu(w), emu(h))
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(fill)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(line_w)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    if s.has_text_frame:
        s.text_frame.text = ""
    return s


# ----------------------------------------------------------------------------- images
class Img:
    def __init__(self, path, trim=True, workdir=None):
        self.src = path
        p = path
        ext = os.path.splitext(path)[1].lower()
        if ext == ".pdf":
            out = os.path.join(workdir, os.path.basename(path) + ".png")
            subprocess.run(["pdftoppm", "-r", "300", "-png", "-singlefile", path, out[:-4]], check=True)
            p = out
        elif ext == ".svg":
            out = os.path.join(workdir, os.path.basename(path) + ".png")
            subprocess.run(["rsvg-convert", "-d", "300", "-p", "300", "-o", out, path], check=True)
            p = out
        im = Image.open(p)
        if ext not in (".png", ".jpg", ".jpeg") or trim:
            im = im.convert("RGBA") if im.mode in ("P", "LA", "RGBA") else im.convert("RGB")
            if trim:
                im = trim_white(im)
            p = os.path.join(workdir, f"img{abs(hash(path)) % 10**8}.png")
            im.save(p)
        self.path = p
        self.w, self.h = im.size
        self.aspect = self.w / self.h  # width / height


def trim_white(im, pad_frac=0.01):
    rgb_im = im.convert("RGB")
    if im.mode == "RGBA":  # treat transparent as white
        bg = Image.new("RGB", im.size, "white")
        bg.paste(im, mask=im.split()[3])
        rgb_im = bg
    diff = ImageChops.difference(rgb_im, Image.new("RGB", im.size, "white"))
    diff = diff.convert("L").point(lambda v: 255 if v > 12 else 0)
    box = diff.getbbox()
    if not box:
        return im
    pad = int(max(im.size) * pad_frac)
    box = (max(0, box[0] - pad), max(0, box[1] - pad), min(im.size[0], box[2] + pad), min(im.size[1], box[3] + pad))
    return im.crop(box)


def resolve_logo(name):
    if os.path.exists(name):
        return name
    p = os.path.join(LOGO_DIR, name if name.endswith(".png") else name + ".png")
    if os.path.exists(p):
        return p
    warn(f"logo not found: {name} (put a PNG in assets/logos/ or pass a file path) — skipped")
    return None


# ----------------------------------------------------------------------------- layout engine
class Ctx:
    def __init__(self, F, scale, base_dir, workdir):
        self.F, self.scale, self.base_dir, self.workdir = F, scale, base_dir, workdir
        self.imgs = {}

    def img(self, path, trim=True):
        full = path if os.path.isabs(path) else os.path.join(self.base_dir, path)
        if not os.path.exists(full):
            raise SystemExit(f"image not found: {full}")
        key = (full, trim)
        if key not in self.imgs:
            self.imgs[key] = Img(full, trim, self.workdir)
        return self.imgs[key]


def text_paras(block, F, k):
    """Return list of paragraph dicts for a text-ish block, with font factor k."""
    body = F["body"] * k
    paras = []
    if "h" in block:
        paras.append(dict(runs=parse_markup(block["h"], True), size=F["sub"] * k, align=block.get("align", "left")))
    if "text" in block:
        for line in str(block["text"]).split("\n\n"):
            paras.append(dict(runs=parse_markup(line, block.get("bold", False)),
                              size=body * block.get("scale", 1), align=block.get("align", "left"),
                              space_before=body * 0.25 if paras else 0))
    if "bullets" in block:
        kind = block.get("bullet", "check")
        for i, item in enumerate(block["bullets"]):
            paras.append(dict(runs=parse_markup(item), size=body * block.get("scale", 1),
                              bullet=None if kind == "none" else kind,
                              space_before=body * 0.3 if (paras or i) else 0))
    if "caption" in block and "image" not in block:
        paras.append(dict(runs=parse_markup(block["caption"]), size=F["caption"] * k,
                          align=block.get("align", "right"), color=COLORS["gray"]))
    return paras


def paras_height(paras, width):
    h = 0.0
    for p in paras:
        avail = width - (p["size"] * 1.15 if p.get("bullet") else 0)
        n = count_lines(p["runs"], p["size"], avail)
        h += n * p["size"] * LH + (p.get("space_before") or 0)
    return h


def is_text_block(b):
    return any(key in b for key in ("h", "text", "bullets")) and not any(
        key in b for key in ("image", "row", "box", "callout", "logos"))


def measure(block, width, ctx, k):
    """Return (fixed_height, flex) where flex is None or dict(weight, max_h, min_h)."""
    F, s = ctx.F, ctx.scale
    if is_text_block(block) or ("caption" in block and "image" not in block and "row" not in block):
        return paras_height(text_paras(block, F, k), width), None
    if "image" in block:
        im = ctx.img(block["image"], block.get("trim", True))
        w = width * block.get("width", 1.0)
        cap_h = F["caption"] * k * LH * count_lines(parse_markup(block["caption"]), F["caption"] * k, width) if block.get("caption") else 0
        if "height_in" in block:
            return block["height_in"] * 72 + cap_h, None
        return cap_h, dict(weight=block.get("weight", 1.0), max_h=w / im.aspect, min_h=0)
    if "row" in block:
        cols = block["row"]
        ws = [c.get("w", 1.0) for c in cols]
        gap = F["body"] * 0.6
        tot = sum(ws)
        inner = width - gap * (len(cols) - 1)
        fixed, maxh, flex = 0, 0, False
        for c, wf in zip(cols, ws):
            cw = inner * wf / tot
            fh, fl = measure(c, cw, ctx, k)
            if fl:
                flex = True
                maxh = max(maxh, fh + fl["max_h"])
            fixed = max(fixed, fh)
        if flex:
            return 0, dict(weight=block.get("weight", 1.0), max_h=max(maxh, fixed), min_h=fixed)
        return fixed, None
    if "box" in block or "callout" in block:
        pad = F["body"] * 0.5
        inner = dict(block)
        if "box" in block:
            inner = {kk: v for kk, v in block.items() if kk != "box"}
            if block["box"]:
                inner["h"] = block["box"]
            inner_f = dict(F)
        else:
            inner = dict(text=block["callout"], align=block.get("align", "center"), bold=block.get("bold", False))
        paras = text_paras(inner, {**F, "sub": F["body"]}, k)  # box titles use body-size bold
        return paras_height(paras, width - 2 * pad) + 2 * pad, None
    if "stack" in block:
        gap = F["body"] * 0.45 * k
        ms = [measure(c, width, ctx, k) for c in block["stack"]]
        fixed = sum(m[0] for m in ms) + gap * (len(ms) - 1)
        fl = [m[1] for m in ms if m[1]]
        if fl:
            return 0, dict(weight=block.get("weight", 1.0), max_h=fixed + sum(f["max_h"] for f in fl),
                           min_h=fixed + sum(f["min_h"] for f in fl))
        return fixed, None
    if "logos" in block:
        return block.get("height_in", 1.4) * 72 * s, None
    if "spacer" in block:
        return block["spacer"] * 72, None
    raise SystemExit(f"unknown block: {block}")


def render(block, slide, x, y, width, height, ctx, k):
    F, s = ctx.F, ctx.scale
    if is_text_block(block) or ("caption" in block and "image" not in block and "row" not in block):
        textbox(slide, x, y, width, height + 2, text_paras(block, F, k))
        return
    if "image" in block:
        im = ctx.img(block["image"], block.get("trim", True))
        cap_h = 0
        if block.get("caption"):
            cap_h = F["caption"] * k * LH * count_lines(parse_markup(block["caption"]), F["caption"] * k, width)
        bw = width * block.get("width", 1.0)
        ih = height - cap_h
        iw = min(bw, ih * im.aspect)
        ih = iw / im.aspect
        align = block.get("align", "center")
        ix = x + {"left": 0, "center": (width - iw) / 2, "right": width - iw}[align]
        iy = y + (height - cap_h - ih) / 2
        slide.shapes.add_picture(im.path, emu(ix), emu(iy), emu(iw), emu(ih))
        dpi = im.w / (iw / 72)
        if dpi < 100:
            warn(f"{os.path.basename(im.src)} printed at {dpi:.0f} dpi (<100): supply a higher-resolution export")
        if ih < 2.0 * 72 * s:
            warn(f"{os.path.basename(im.src)} squeezed to {ih/72:.1f} in tall: give its panel more room or cut text")
        if cap_h:
            textbox(slide, x, iy + ih + 2, width, cap_h + 2,
                    [dict(runs=parse_markup(block["caption"]), size=F["caption"] * k,
                          align=block.get("caption_align", "right"), color=COLORS["gray"])])
        return
    if "row" in block:
        cols = block["row"]
        ws = [c.get("w", 1.0) for c in cols]
        gap = F["body"] * 0.6
        inner = width - gap * (len(cols) - 1)
        cx = x
        for c, wf in zip(cols, ws):
            cw = inner * wf / sum(ws)
            fh, fl = measure(c, cw, ctx, k)
            if fl:
                ch = min(height, fh + fl["max_h"]) if fl["max_h"] else height
            else:
                ch = fh
            render(c, slide, cx, y + (height - ch) / 2, cw, ch, ctx, k)
            cx += cw + gap
        return
    if "box" in block or "callout" in block:
        pad = F["body"] * 0.5
        if "box" in block:
            inner = {kk: v for kk, v in block.items() if kk != "box"}
            if block["box"]:
                inner["h"] = block["box"]
            fill = COLORS.get(block.get("fill", "box"), block.get("fill", "box"))
            shape = MSO_SHAPE.RECTANGLE
        else:
            inner = dict(text=block["callout"], align=block.get("align", "center"), bold=block.get("bold", False))
            fill = COLORS.get(block.get("fill", "gray_box"), block.get("fill", "gray_box"))
            shape = MSO_SHAPE.ROUNDED_RECTANGLE
        line = COLORS["border"] if "callout" in block else None
        rect(slide, x, y, width, height, fill=fill, line=line, line_w=1.5 * s, shape=shape, radius=0.15)
        paras = text_paras(inner, {**F, "sub": F["body"]}, k)
        textbox(slide, x + pad, y + pad, width - 2 * pad, height - 2 * pad, paras,
                anchor="middle" if "callout" in block else "top")
        return
    if "logos" in block:
        draw_logos(slide, block["logos"], x, y, width, height, block.get("align", "right"), s)
        return
    if "stack" in block:
        layout_blocks(block["stack"], x, y, width, height, ctx, slide, "stack", k=k, spread=False)
        return
    if "spacer" in block:
        return


def draw_logos(slide, names, x, y, width, height, align, s):
    items = []
    for n in names:
        p = resolve_logo(n)
        if not p:
            continue
        with Image.open(p) as im:
            items.append((p, im.size[0] / im.size[1]))
    if not items:
        return
    gap = 0.3 * 72 * s
    h = height
    total = sum(h * a for _, a in items) + gap * (len(items) - 1)
    if total > width:  # shrink to fit
        h *= width / total
        total = width
    cx = x + {"left": 0, "center": (width - total) / 2, "right": width - total}[align]
    for p, a in items:
        slide.shapes.add_picture(p, emu(cx), emu(y + (height - h) / 2), emu(h * a), emu(h))
        cx += h * a + gap


def layout_blocks(blocks, x, y, width, height, ctx, slide, label, k=None, spread=True):
    """Stack blocks vertically inside a panel body; shrink fonts if they overflow."""
    F = ctx.F
    fixed_k = k is not None
    k = k or 1.0
    while True:
        gap = F["body"] * 0.45 * k
        ms = [measure(b, width, ctx, k) for b in blocks]
        fixed = sum(m[0] for m in ms) + gap * (len(blocks) - 1)
        mins = sum(m[1]["min_h"] for m in ms if m[1])
        if fixed_k or fixed + mins <= height or k <= 0.72:
            break
        k -= 0.04
    if k < 1.0 and not fixed_k:
        warn(f"panel '{label}': text too long, fonts shrunk to {k:.0%} — consider trimming text")
    if fixed + mins > height + 1 and not fixed_k:
        warn(f"panel '{label}': content overflows by {(fixed + mins - height)/72:.1f} in even after shrinking")
    # distribute flexible space
    avail = max(0.0, height - fixed)
    alloc = {}
    flex = {i: m[1] for i, m in enumerate(ms) if m[1]}
    open_ = dict(flex)
    while open_:
        wsum = sum(f["weight"] for f in open_.values())
        capped = False
        for i, f in list(open_.items()):
            share = avail * f["weight"] / wsum
            if share >= f["max_h"]:
                alloc[i] = f["max_h"]
                avail -= f["max_h"]
                del open_[i]
                capped = True
        if not capped:
            for i, f in open_.items():
                alloc[i] = max(f["min_h"], avail * f["weight"] / wsum)
            avail = 0
            break
    leftover = max(0.0, avail)
    # spread leftover as extra gaps (capped), remainder split above/below
    # a lone subheading stays glued to the block after it
    glue = [set(b) - {"align"} == {"h"} for b in blocks]
    n_gaps = max(1, sum(1 for i in range(len(blocks) - 1) if not glue[i]))
    extra = min(leftover / (n_gaps + 2), F["body"] * 1.5) if spread else 0
    top = (leftover - extra * n_gaps) / 2 if len(blocks) > 1 else leftover / 2
    top = min(top, F["body"]) if spread else top
    cy = y + top
    for i, (b, m) in enumerate(zip(blocks, ms)):
        h = m[0] + alloc.get(i, 0)
        render(b, slide, x, cy, width, h, ctx, k)
        cy += h + gap + (0 if glue[i] else extra)


def draw_panel(slide, panel, x, y, w, h, header_h, ctx):
    F, s = ctx.F, ctx.scale
    lw = 6 * s
    pad = 0.35 * 72 * s
    rect(slide, x, y, w, h, fill="FFFFFF", line=COLORS["border"], line_w=lw)
    rect(slide, x, y, w, header_h, fill=COLORS["header_fill"], line=COLORS["border"], line_w=lw)
    textbox(slide, x + pad * 0.6, y, w - pad * 1.2, header_h,
            [dict(runs=parse_markup(panel["title"], True), size=F["header"])], anchor="middle")
    layout_blocks(panel.get("blocks", []), x + pad, y + header_h + pad * 0.8, w - 2 * pad,
                  h - header_h - pad * 1.6, ctx, slide, panel["title"])


def header_height(title, w, ctx):
    F, s = ctx.F, ctx.scale
    pad = 0.35 * 72 * s
    n = count_lines(parse_markup(title, True), F["header"], w - pad * 1.2)
    return n * F["header"] * LH + 0.3 * 72 * s


# ----------------------------------------------------------------------------- top / bottom bands
def draw_title_block(slide, spec, ctx, W, m):
    F, s = ctx.F, ctx.scale
    y = m * 0.7
    tb_w = W - 2 * m
    title_runs = parse_markup(spec["title"], spec.get("title_bold", True))
    n = count_lines(title_runs, F["title"], tb_w)
    th = n * F["title"] * LH
    textbox(slide, m, y, tb_w, th, [dict(runs=title_runs, size=F["title"])])
    y += th + F["authors"] * 0.25
    # authors
    auth_parts = []
    authors = spec.get("authors", [])
    for i, a in enumerate(authors):
        if isinstance(a, str):
            a = {"name": a}
        name = a["name"]
        is_first = a.get("presenting", i == 0)
        sup = ",".join(str(v) for v in a.get("aff", []))
        if a.get("equal"):
            sup += "*" if sup else "*"
        piece = (f"**{name}**" if is_first else name) + (f"^{sup}^" if sup else "")
        sep = "" if i == 0 else (" and " if i == len(authors) - 1 and spec.get("and_last", True) else ", ")
        auth_parts.append(sep + piece)
    email_w = 0
    if spec.get("email"):
        email_w = text_width_pt(spec["email"], False, False, F["aff"]) + F["aff"] * 1.8
    a_w = W - 2 * m - email_w - 0.4 * 72 * s
    a_runs = parse_markup("".join(auth_parts))
    ah = count_lines(a_runs, F["authors"], a_w) * F["authors"] * LH
    textbox(slide, m, y, a_w, ah, [dict(runs=a_runs, size=F["authors"])])
    if spec.get("email"):
        icon = F["aff"] * 1.3
        ex = W - m - email_w
        if resolve_logo("email_icon"):
            slide.shapes.add_picture(resolve_logo("email_icon"), emu(ex), emu(y + (F["authors"] * LH - icon) / 2),
                                     emu(icon), emu(icon))
        textbox(slide, ex + icon * 1.25, y, email_w - icon * 1.25, F["authors"] * LH,
                [dict(runs=[dict(text=spec["email"], bold=False, italic=False, sup=False, sub=False, color=None)],
                      size=F["aff"])], anchor="middle")
    y += ah + F["aff"] * 0.2
    affs = spec.get("affiliations", [])
    if affs:
        txt = "; ".join(f"{i+1} {a}" for i, a in enumerate(affs))
        runs = parse_markup(txt)
        ah2 = count_lines(runs, F["aff"], W - 2 * m) * F["aff"] * LH
        textbox(slide, m, y, W - 2 * m, ah2, [dict(runs=runs, size=F["aff"])])
        y += ah2
    return y + 0.45 * 72 * s


def draw_footer(slide, footer, ctx, W, H, m):
    """Bottom strip: funding text left, logos right. Returns top y of the strip."""
    F, s = ctx.F, ctx.scale
    logo_h = footer.get("logo_height_in", 1.25) * 72 * s
    logos = [n for n in footer.get("logos", []) if resolve_logo(n)]
    logo_w = 0
    if logos:
        for n in logos:
            with Image.open(resolve_logo(n)) as im:
                logo_w += logo_h * im.size[0] / im.size[1]
        logo_w += 0.3 * 72 * s * (len(logos) - 1)
    text_w = W - 2 * m - logo_w - (0.5 * 72 * s if logos else 0)
    size = footer.get("size", F["caption"] * 1.1)
    runs = parse_markup(footer.get("text", ""))
    th = count_lines(runs, size, text_w) * size * LH if runs else 0
    strip_h = max(logo_h, th)
    y = H - m * 0.6 - strip_h
    if runs:
        textbox(slide, m, y, text_w, strip_h, [dict(runs=runs, size=size)], anchor="middle")
    if logos:
        draw_logos(slide, logos, W - m - logo_w, y, logo_w, logo_h, "right", s)
    return y


# ----------------------------------------------------------------------------- main
def build(spec_path, out_dir, make_pdf=True, preview_dpi=24):
    with open(spec_path) as fh:
        spec = yaml.safe_load(fh)
    base_dir = os.path.dirname(os.path.abspath(spec_path))
    os.makedirs(out_dir, exist_ok=True)
    workdir = tempfile.mkdtemp(prefix="poster_")

    preset = PRESETS[spec.get("preset", "landscape")]
    Wi, Hi = spec.get("size_in", preset["size"])
    scale = Wi / preset["size"][0] if "size_in" in spec else 1.0
    F = {kk: v * scale for kk, v in preset.items() if kk != "size"}
    F.update(spec.get("fonts", {}))
    W, H = Wi * 72, Hi * 72
    ctx = Ctx(F, Wi / 45 if "size_in" in spec else preset["size"][0] / 45, base_dir, workdir)
    ctx.scale = max(ctx.scale, 0.73)

    prs = Presentation()
    prs.slide_width, prs.slide_height = emu(W), emu(H)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    m = spec.get("margin_in", 0.6) * 72 * (Wi / 45)
    gap = spec.get("gap_in", 0.45) * 72 * (Wi / 45)

    top = draw_title_block(slide, spec, ctx, W, m)
    bottom = H - m * 0.7
    if spec.get("footer"):
        bottom = draw_footer(slide, spec["footer"], ctx, W, H, m) - gap

    if "rows" in spec:  # grid: rows of panels, headers aligned within a row
        rows = spec["rows"]
        rw = [r.get("height", 1.0) for r in rows]
        avail_h = bottom - top - gap * (len(rows) - 1)
        y = top
        for r, wt in zip(rows, rw):
            rh = avail_h * wt / sum(rw)
            pw = [p.get("width", 1.0) for p in r["panels"]]
            avail_w = W - 2 * m - gap * (len(pw) - 1)
            widths = [avail_w * v / sum(pw) for v in pw]
            hh = max(header_height(p["title"], w, ctx) for p, w in zip(r["panels"], widths))
            x = m
            for p, w in zip(r["panels"], widths):
                draw_panel(slide, p, x, y, w, rh, hh, ctx)
                x += w + gap
            y += rh + gap
    else:  # columns of stacked panels
        cols = spec["columns"]
        cw = [c.get("width", 1.0) for c in cols]
        avail_w = W - 2 * m - gap * (len(cols) - 1)
        x = m
        for c, wt in zip(cols, cw):
            w = avail_w * wt / sum(cw)
            ph = [p.get("height", 1.0) for p in c["panels"]]
            avail_h = bottom - top - gap * (len(ph) - 1)
            y = top
            for p, hwt in zip(c["panels"], ph):
                h = avail_h * hwt / sum(ph)
                draw_panel(slide, p, x, y, w, h, header_height(p["title"], w, ctx), ctx)
                y += h + gap
            x += w + gap

    stem = spec.get("output_name", os.path.splitext(os.path.basename(spec_path))[0])
    pptx_path = os.path.join(out_dir, stem + ".pptx")
    prs.save(pptx_path)
    result = {"pptx": pptx_path, "size_in": [Wi, Hi], "warnings": WARN}
    if make_pdf:
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        mac = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if not soffice and os.path.exists(mac):
            soffice = mac
        if soffice:
            subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, pptx_path],
                           check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            pdf = os.path.join(out_dir, stem + ".pdf")
            result["pdf"] = pdf
            prev = os.path.join(out_dir, stem + "_preview")
            subprocess.run(["pdftoppm", "-r", str(preview_dpi), "-png", "-singlefile", pdf, prev], check=True)
            result["preview"] = prev + ".png"
        else:
            warn("LibreOffice not found: open the .pptx in PowerPoint and File > Export > PDF")
    shutil.rmtree(workdir, ignore_errors=True)
    print(json.dumps(result, indent=2))
    return result


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("-o", "--out", default=".")
    ap.add_argument("--no-pdf", action="store_true")
    ap.add_argument("--preview-dpi", type=int, default=24)
    a = ap.parse_args()
    build(a.spec, a.out, not a.no_pdf, a.preview_dpi)
