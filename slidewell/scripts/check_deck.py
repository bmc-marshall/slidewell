#!/usr/bin/env python3
"""Check a .pptx against the measurable Slidewell design rules.

Usage:
    python check_deck.py deck.pptx [--theme theme.json] [--slides 3,5-7] [--json]

Reports ERRORS (must fix) and WARNINGS (look at the rendered slide and decide)
per slide, then deck-level findings. Exit code 1 if there are errors.

What it checks: font sizes against the tiers, number of fonts, banned fonts,
italics/underline/shadow, likely text overflow, headline length, palette,
red/green use, near-misses in alignment and size, consistency of equivalent
boxes, margins, recurring-element positions across slides, page numbers,
leftover placeholders, flattened slides, chart axes/titles/types, numbers that
probably belong in a chart, and over-precise numbers.

What it can't check (do these on the rendered slides): story, hierarchy,
balance, label-to-mark alignment, and anything needing judgement.

Optional theme.json (all keys optional):
{
  "palette": ["#1F3A5F", "#E4572E", ...],   # brand colours; neutrals always allowed
  "fonts": ["Inter", "Georgia"],            # allowed font families
  "page_number_exempt": [1, 7, 15]          # extra slides that may omit page numbers
}
"""
import argparse
import colorsys
import json
import os
import re
import subprocess
import sys
from collections import Counter, defaultdict
from functools import lru_cache

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
EMU_IN = 914400
NEAR = int(0.06 * EMU_IN)      # "almost aligned" window: (0, 0.06 in]
EXACT = int(0.005 * EMU_IN)    # treat as aligned below this
BANNED_FONTS = {"calibri", "papyrus", "comic sans ms", "algerian", "aharoni"}
SOURCE_RE = re.compile(r"^\s*(source|sources|note|notes|\[?\d+\]|\d\s|\*)", re.I)


# ----------------------------------------------------------------- helpers

def inch(v):
    return round(v / EMU_IN, 2)


class Finding:
    def __init__(self, level, rule, msg, slide=None):
        self.level, self.rule, self.msg, self.slide = level, rule, msg, slide

    def as_dict(self):
        return {"level": self.level, "rule": self.rule, "slide": self.slide, "message": self.msg}


def iter_shapes(shapes, xf=None):
    """Yield (shape, (left, top, width, height)) with absolute positions, recursing into groups."""
    for sh in shapes:
        l, t, w, h = sh.left or 0, sh.top or 0, sh.width or 0, sh.height or 0
        if xf:
            l, t, w, h = xf(l, t, w, h)
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            x = sh._element.grpSpPr.find("a:xfrm", NS)
            if x is not None and x.find("a:chOff", NS) is not None:
                off, ext = x.find("a:off", NS), x.find("a:ext", NS)
                coff, cext = x.find("a:chOff", NS), x.find("a:chExt", NS)
                ox, oy = int(off.get("x")), int(off.get("y"))
                ex, ey = int(ext.get("cx")) or 1, int(ext.get("cy")) or 1
                cx, cy = int(coff.get("x")), int(coff.get("y"))
                cex, cey = int(cext.get("cx")) or 1, int(cext.get("cy")) or 1
                sx, sy = ex / cex, ey / cey

                def child(ll, tt, ww, hh, ox=ox, oy=oy, cx=cx, cy=cy, sx=sx, sy=sy, parent=xf):
                    r = (ox + (ll - cx) * sx, oy + (tt - cy) * sy, ww * sx, hh * sy)
                    return parent(*r) if parent else r
                yield from iter_shapes(sh.shapes, child)
            else:
                yield from iter_shapes(sh.shapes, xf)
        else:
            rot = getattr(sh, "rotation", 0) or 0
            if round(rot) % 180 == 90:      # rotated a quarter turn: the visible box swaps width and height
                cx, cy = l + w / 2, t + h / 2
                l, t, w, h = cx - h / 2, cy - w / 2, h, w
            yield sh, (int(l), int(t), int(w), int(h))


def hex_of(el, theme_colors):
    """Resolve the colour of a fill/line/run element to #RRGGBB, or None."""
    if el is None:
        return None
    s = el.find(".//a:srgbClr", NS)
    if s is not None:
        return "#" + s.get("val").upper()
    sc = el.find(".//a:schemeClr", NS)
    if sc is not None:
        base = theme_colors.get(sc.get("val"))
        if base is None:
            return None
        # apply luminance modifiers (lumMod/lumOff, e.g. "Accent 4, lighter 40%")
        mod, off = sc.find("a:lumMod", NS), sc.find("a:lumOff", NS)
        if mod is None and off is None:
            return base
        r, g, b = (int(base[i:i + 2], 16) / 255 for i in (1, 3, 5))
        h, l, sat = colorsys.rgb_to_hls(r, g, b)
        l = l * (int(mod.get("val")) / 100000 if mod is not None else 1) + (int(off.get("val")) / 100000 if off is not None else 0)
        r, g, b = colorsys.hls_to_rgb(h, max(0, min(1, l)), sat)
        return "#%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))
    return None


def is_neutral(hx):
    r, g, b = (int(hx[i:i + 2], 16) / 255 for i in (1, 3, 5))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    return s < 0.12 or l > 0.93 or l < 0.08


def hue_family(hx):
    r, g, b = (int(hx[i:i + 2], 16) / 255 for i in (1, 3, 5))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    if s < 0.25 or l < 0.15 or l > 0.9:
        return None
    deg = h * 360
    if deg < 15 or deg >= 345:
        return "red"
    if 85 <= deg < 160:
        return "green"
    return None


def color_dist(a, b):
    pa = [int(a[i:i + 2], 16) for i in (1, 3, 5)]
    pb = [int(b[i:i + 2], 16) for i in (1, 3, 5)]
    return sum((x - y) ** 2 for x, y in zip(pa, pb)) ** 0.5


@lru_cache(maxsize=None)
def font_file(family, bold=False):
    """(path, family actually used), via fontconfig or the OS font folders (sw_fonts.py)."""
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from sw_fonts import find_font
    return find_font(family, bold)


@lru_cache(maxsize=None)
def pil_font(family, bold, size_pt):
    from PIL import ImageFont
    path, _ = font_file(family or "Arial", bold)
    try:
        return ImageFont.truetype(path, max(1, int(round(size_pt * 10))))
    except Exception:
        return ImageFont.load_default()


def text_width_pt(text, family, bold, size_pt):
    f = pil_font(family, bold, size_pt)
    try:
        return f.getlength(text) / 10.0
    except Exception:
        return len(text) * size_pt * 0.5


def wrap_lines(text, width_pt, family, bold, size_pt):
    if width_pt <= 0:
        return 1
    lines = 0
    for para in text.split("\n"):
        words = para.split(" ")
        if not para.strip():
            lines += 1
            continue
        cur = ""
        n = 1
        for w in words:
            trial = (cur + " " + w).strip()
            if text_width_pt(trial, family, bold, size_pt) <= width_pt or not cur:
                cur = trial
            else:
                n += 1
                cur = w
        lines += n
    return lines


# ----------------------------------------------------------- deck context

class Deck:
    def __init__(self, path, theme):
        self.prs = Presentation(path)
        self.W, self.H = self.prs.slide_width, self.prs.slide_height
        self.theme = theme or {}
        self.theme_colors, self.major, self.minor = self._theme()
        self.default_size = self._default_size()

    def _theme(self):
        master = self.prs.slide_masters[0]
        colors, major, minor = {}, None, None
        try:
            tpart = master.part.part_related_by(RT.THEME)
            root = etree.fromstring(tpart.blob)
            scheme = root.find(".//a:clrScheme", NS)
            for child in scheme:
                tag = etree.QName(child).localname
                v = hex_of(child, {})
                if v is None:
                    sc = child.find("a:sysClr", NS)
                    if sc is not None:
                        v = "#" + sc.get("lastClr", "000000").upper()
                colors[tag] = v
            colors.update({"tx1": colors.get("dk1"), "bg1": colors.get("lt1"),
                           "tx2": colors.get("dk2"), "bg2": colors.get("lt2")})
            mj = root.find(".//a:majorFont/a:latin", NS)
            mn = root.find(".//a:minorFont/a:latin", NS)
            major = mj.get("typeface") if mj is not None else None
            minor = mn.get("typeface") if mn is not None else None
        except Exception:
            pass
        return colors, major, minor

    def _default_size(self):
        el = self.prs.part._element.find(".//p:defaultTextStyle/a:lvl1pPr/a:defRPr", NS)
        if el is not None and el.get("sz"):
            return int(el.get("sz")) / 100
        return 18.0

    def resolve_font(self, name):
        if not name:
            return self.minor
        if name.startswith("+mj"):
            return self.major
        if name.startswith("+mn"):
            return self.minor
        return name


def master_style_size(shape, level=0):
    """Size from the slide master's title/body/other style for placeholders."""
    try:
        master = shape.part.slide_layout.slide_master
        tx = master._element.find("p:txStyles", NS)
        kind = "p:bodyStyle"
        if shape.is_placeholder and shape.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
            kind = "p:titleStyle"
        el = tx.find(f"{kind}/a:lvl{level + 1}pPr/a:defRPr", NS)
        if el is not None and el.get("sz"):
            return int(el.get("sz")) / 100
    except Exception:
        pass
    return None


def layout_ph_size(shape, level=0):
    try:
        idx = shape.placeholder_format.idx
        for ph in shape.part.slide_layout.placeholders:
            if ph.placeholder_format.idx == idx:
                for el in ph._element.iter(f"{{{NS['a']}}}defRPr", f"{{{NS['a']}}}rPr"):
                    if el.get("sz"):
                        return int(el.get("sz")) / 100
    except Exception:
        pass
    return None


def effective_size(run, para, shape, deck):
    """(size_pt, inherited?)"""
    if run.font.size:
        return run.font.size.pt, False
    pPr = para._p.find("a:pPr", NS)
    if pPr is not None:
        d = pPr.find("a:defRPr", NS)
        if d is not None and d.get("sz"):
            return int(d.get("sz")) / 100, True
    lvl = para.level or 0
    lst = shape.text_frame._txBody.find("a:lstStyle", NS)
    if lst is not None:
        d = lst.find(f"a:lvl{lvl + 1}pPr/a:defRPr", NS)
        if d is not None and d.get("sz"):
            return int(d.get("sz")) / 100, True
    if shape.is_placeholder:
        s = layout_ph_size(shape, lvl) or master_style_size(shape, lvl)
        if s:
            return s, True
    return deck.default_size, True


def has_effect(el, tag):
    return el.find(f".//a:effectLst/a:{tag}", NS) is not None


# --------------------------------------------------------------- slide info

class ShapeInfo:
    def __init__(self, sh, box, deck):
        self.sh, self.box, self.deck = sh, box, deck
        self.l, self.t, self.w, self.h = box
        self.r, self.b = self.l + self.w, self.t + self.h
        self.name = sh.name
        self.text = sh.text_frame.text.strip() if getattr(sh, "has_text_frame", False) and sh.has_text_frame else ""
        self.is_pic = sh.shape_type == MSO_SHAPE_TYPE.PICTURE
        self.is_chart = getattr(sh, "has_chart", False) and sh.has_chart
        self.is_table = getattr(sh, "has_table", False) and sh.has_table
        self.is_line = sh.shape_type == MSO_SHAPE_TYPE.LINE or sh._element.tag.endswith("cxnSp")
        self.area_frac = (self.w * self.h) / (deck.W * deck.H) if deck.W and deck.H else 0
        self.placeholder_type = sh.placeholder_format.type if sh.is_placeholder else None
        spPr = sh._element.find(".//p:spPr", NS)
        self.fill = None
        self.line = None
        if spPr is not None:
            sf = spPr.find("a:solidFill", NS)
            self.fill = hex_of(sf, deck.theme_colors) if sf is not None else None
            ln = spPr.find("a:ln", NS)
            if ln is not None and ln.find("a:noFill", NS) is None:
                self.line = (hex_of(ln.find("a:solidFill", NS), deck.theme_colors), ln.get("w"))
        prst = sh._element.find(".//a:prstGeom", NS)
        self.geom = prst.get("prst") if prst is not None else None
        self.adj = tuple(g.get("fmla") for g in prst.findall(".//a:gd", NS)) if prst is not None else ()

    @property
    def is_bleed(self):
        return self.area_frac > 0.6 and (self.is_pic or self.fill)

    def first_run_style(self):
        if not self.text:
            return None
        for p in self.sh.text_frame.paragraphs:
            for r in p.runs:
                if r.text.strip():
                    size, _ = effective_size(r, p, self.sh, self.deck)
                    return (round(size, 1), bool(r.font.bold), str(p.alignment))
        return None

    def body_props(self):
        if not self.text:
            return None
        bp = self.sh.text_frame._txBody.find("a:bodyPr", NS)
        return tuple(bp.get(k) for k in ("lIns", "tIns", "rIns", "bIns", "anchor"))


def classify(deck, slide_idx, infos):
    """Find headline, breadcrumb, source box and page number on a slide."""
    H = deck.H
    roles = {}
    titles = [i for i in infos if i.placeholder_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE) and i.text]
    if titles and not (titles[0].text.upper() == titles[0].text and len(titles[0].text) < 50):
        roles["headline"] = titles[0]
    else:
        if titles:  # an all-caps title placeholder is being used as a breadcrumb
            roles["breadcrumb"] = titles[0]
        top = [i for i in infos if i.text and i.t < H * 0.25 and len(i.text) > 12 and i is not roles.get("breadcrumb")
               and not i.name.startswith("Photo placeholder")
               and i.text.upper() != i.text]
        best, best_size = None, 0
        for i in top:
            st = i.first_run_style()
            if st and st[0] > best_size:
                best, best_size = i, st[0]
        if best:
            roles["headline"] = best
    hl = roles.get("headline")
    if hl and "breadcrumb" not in roles:
        above = [i for i in infos if i.text and i is not hl and i.t < hl.t and i.b <= hl.t + int(0.12 * EMU_IN) and i.t < H * 0.2
                 and len(i.text) < 70 and i.text.upper() == i.text]
        if above:
            roles["breadcrumb"] = above[0]
    srcs = [i for i in infos if i.text and i.t > H * 0.75 and SOURCE_RE.match(i.text)]
    if srcs:
        roles["source"] = min(srcs, key=lambda i: i.l)
    nums = []
    for i in infos:
        if i.placeholder_type == PP_PLACEHOLDER.SLIDE_NUMBER:
            nums.append(i)
        elif i.sh._element.find(".//a:fld[@type='slidenum']", NS) is not None:
            nums.append(i)
        elif i.text.strip() == str(slide_idx) and (i.t > H * 0.85 or i.t < H * 0.1):
            nums.append(i)
    if nums:
        roles["page_number"] = nums[0]
    return roles


def layout_has_slidenum(slide):
    try:
        lay = slide.slide_layout
        for el in (lay._element, lay.slide_master._element):
            for sp in el.iter(f"{{{NS['p']}}}sp"):
                is_ph = sp.find(".//p:nvPr/p:ph", NS) is not None
                if not is_ph and sp.find(".//a:fld[@type='slidenum']", NS) is not None:
                    return True
    except Exception:
        pass
    return False


# ------------------------------------------------------------------ checks

def check_text(deck, idx, infos, roles, out):
    fonts_used = Counter()
    long_numbers = []
    numeric_tokens = 0
    for i in infos:
        if not i.text:
            continue
        tf = i.sh.text_frame
        role = next((k for k, v in roles.items() if v is i), None)
        is_small_ok = (role in ("source", "breadcrumb", "page_number") or bool(SOURCE_RE.match(i.text)) or i.name.startswith("Photo placeholder")
                       or (i.t > deck.H * 0.8 and i.h < deck.H * 0.2))
        words = len(i.text.split())
        sizes = []
        size_issue = None   # (level, msg) worst per shape
        styles = Counter()
        for p in tf.paragraphs:
            for r in p.runs:
                if not r.text.strip():
                    continue
                size, inherited = effective_size(r, p, i.sh, deck)
                sizes.append(size)
                fam = deck.resolve_font(r.font.name) or "(theme default)"
                fonts_used[fam] += len(r.text)
                rPr = r._r.find("a:rPr", NS)
                if r.font.italic and size > 12:      # small italic is fine for legends and keys
                    styles["italic"] += 1
                if r.font.underline:
                    styles["underline"] += 1
                if rPr is not None and has_effect(rPr, "outerShdw"):
                    styles["shadow"] += 1
                if fam.lower() in BANNED_FONTS and fam not in (deck.major, deck.minor):   # a template's own theme font is the brand's choice
                    styles["banned:" + fam] += 1
                tag = " (inherited size, verify)" if inherited else ""
                cand = None
                if size < 10:
                    cand = (3, "error", f"{size:g}pt text in '{i.name}' (minimum 10pt for any text){tag}: {short(r.text)!r}")
                elif size < 14 and not is_small_ok and words > 12:
                    cand = (2, "error", f"{size:g}pt body text in '{i.name}' (main ≥16, supporting ≥14; 10–13 only for sources, notes, chart labels, short annotations){tag}: {short(i.text)!r}")
                elif size < 14 and not is_small_ok and not (size >= 10 and words <= 4 and (i.b < 0.1 * deck.H or i.t > 0.9 * deck.H)):
                    # (navigation labels and trackers at 10-13pt in the top or bottom band are allowed: finishing-standard §4)
                    cand = (1, "warning", f"{size:g}pt short text in '{i.name}': OK only if it's an annotation or label{tag}: {short(i.text)!r}")
                elif size < 16 and words > 25 and not i.fill and not i.line:
                    cand = (0, "warning", f"{size:g}pt for a long text block in '{i.name}': main text should be ≥16pt{tag}")
                if cand and (size_issue is None or cand[0] > size_issue[0]):
                    size_issue = cand
        if size_issue:
            out.append(Finding(size_issue[1], "font-size", size_issue[2], idx))
        for k, n in styles.items():
            if k.startswith("banned:"):
                continue  # reported once at deck level
            else:
                out.append(Finding("warning", "type", f"{k.capitalize()} used in '{i.name}' ({n} run(s)): fine only if the template uses it intentionally", idx))
        if not sizes:
            continue
        # numbers
        for m in re.finditer(r"(?<![\w.])[$£€]?\d[\d,]*\.\d{2,}%?", i.text):
            if re.fullmatch(r"[$£€]\d+\.\d{2}", m.group(0)):
                continue            # money in pence/cents (£7.10) is conventional, not over-precise
            long_numbers.append(m.group(0))
        if not (role in ("page_number",) or is_small_ok):
            nums = re.findall(r"(?<![\w])[$£€]?\d[\d,.]*", i.text)
            numeric_tokens += sum(1 for n in nums if not re.fullmatch(r"(19|20)\d\d", n) and len(i.text.split()) >= 8)
        # overflow / headline lines
        check_fit(deck, idx, i, role, sizes, out)
    for n in sorted(set(long_numbers))[:5]:
        out.append(Finding("warning", "rounding", f"Over-precise number {n!r}: round sensibly (charts.md §5)", idx))
    return fonts_used, numeric_tokens


def short(t, n=60):
    t = " ".join((t or "").split())
    return t if len(t) <= n else t[: n - 1] + "…"


def check_fit(deck, idx, i, role, sizes, out):
    tf = i.sh.text_frame
    bp = tf._txBody.find("a:bodyPr", NS)
    if bp is None:
        return
    if bp.find("a:normAutofit", NS) is not None:
        out.append(Finding("warning", "overflow", f"'{i.name}' uses shrink-text-on-overflow: text may be rendered smaller than its set size", idx))
    if bp.find("a:spAutoFit", NS) is not None:
        return  # box grows with text
    ins = lambda k, d: int(bp.get(k, d)) / 12700  # EMU -> pt
    width_pt = i.w / 12700 - ins("lIns", 91440) - ins("rIns", 91440)
    height_pt = i.h / 12700 - ins("tIns", 45720) - ins("bIns", 45720)
    wrap = bp.get("wrap", "square") != "none"
    total = 0.0
    lines_first = None
    last_line_slack = last_after = 0
    for p in tf.paragraphs:
        runs = [r for r in p.runs if r.text]
        txt = "".join(r.text for r in runs)
        if not runs:
            end = p._p.find("a:endParaRPr", NS)
            esz = int(end.get("sz")) / 100 if end is not None and end.get("sz") else (sizes[0] if sizes else 12)
            total += esz * 1.2
            continue
        size, _ = effective_size(runs[0], p, i.sh, deck)
        fam = deck.resolve_font(runs[0].font.name) or "Arial"
        bold = bool(runs[0].font.bold)
        indent = 18 if p.level else 0
        txt_m = txt.replace("\t", "    ")
        if wrap:
            n = wrap_lines(txt_m, width_pt - indent, fam, bold, size)
        else:
            n = txt.count("\n") + 1
            if (i.fill or i.line) and text_width_pt(txt_m, fam, bold, size) > width_pt * 1.03:
                out.append(Finding("error", "overflow", f"Text runs past the edge of the box '{i.name}' (wrap is off): {short(txt)!r}", idx))
        if lines_first is None:
            lines_first = n
        spacing = 1.2
        lnSpc = p._p.find("a:pPr/a:lnSpc/a:spcPct", NS)
        if lnSpc is not None:
            spacing = 1.2 * int(lnSpc.get("val")) / 100000
        total += n * size * spacing
        last_line_slack = size * spacing - size * 0.95
        for tag in ("spcBef", "spcAft"):
            sp = p._p.find(f"a:pPr/a:{tag}/a:spcPts", NS)
            if sp is not None:
                last_after = int(sp.get("val")) / 100 if tag == "spcAft" else 0
                total += int(sp.get("val")) / 100
            elif tag == "spcAft":
                last_after = 0
    total -= last_line_slack + last_after if tf.paragraphs else 0
    if role == "headline" and idx not in (1, len(deck.prs.slides)):
        all_txt = tf.text
        st = i.first_run_style()
        fam = deck.resolve_font(next((r.font.name for p in tf.paragraphs for r in p.runs if r.text.strip()), None)) or "Arial"
        n = wrap_lines(all_txt, width_pt, fam, bool(st and st[1]), st[0] if st else 24)
        if n > 2:
            out.append(Finding("error", "headline", f"Headline wraps to ~{n} lines (max 2): {short(all_txt, 90)!r}", idx))
    if not wrap or i.h < 0.1 * EMU_IN:
        return  # auto-sized labels: no vertical box to overflow
    if height_pt > 0 and total > height_pt * 1.12:
        visible = bool(i.fill or i.line)
        level = "error" if (total > height_pt * 1.4 and visible) else "warning"
        out.append(Finding(level, "overflow",
                           f"Text in '{i.name}' needs ~{total:.0f}pt of height but the box has {height_pt:.0f}pt: likely overflow (font substitution can affect this; confirm on render)", idx))


def check_geometry(deck, idx, infos, out):
    W, H = deck.W, deck.H
    content = [i for i in infos if not i.is_bleed and i.w > 0 and i.h > 0]
    photo = lambda i: i.is_pic or i.name.startswith(("Photo placeholder", "Overlay"))
    for i in content:
        if photo(i):
            continue            # photos may bleed off an edge deliberately (half-bleed, shaped crops)
        if i.l < -EXACT or i.t < -EXACT or i.r > W + EXACT or i.b > H + EXACT:
            out.append(Finding("warning", "layout", f"'{i.name}' extends beyond the slide edge", idx))
    # margins (skipped on slides with a full- or half-bleed panel or photo: those are composed against the edge)
    edges = lambda i: sum([i.l <= EXACT, i.t <= EXACT, i.r >= W - EXACT, i.b >= H - EXACT])
    half_bleed = any(edges(i) >= 3 and i.area_frac >= 0.2 for i in infos)
    inner = [i for i in content if i.l > EXACT and i.r < W - EXACT and i.area_frac < 0.6]
    if len(inner) >= 3 and not any(i.is_bleed for i in infos) and not half_bleed:
        left = min(i.l for i in inner)
        right = W - max(i.r for i in inner)
        if abs(left - right) > 0.3 * EMU_IN:
            out.append(Finding("warning", "margins", f"Outer margins unbalanced: left {inch(left)} in vs right {inch(right)} in", idx))
    # equivalent boxes: same geometry, similar size
    boxes = [i for i in content if i.geom and not i.is_line and not i.is_pic and (i.fill or i.line)]
    groups = defaultdict(list)
    for b in boxes:
        for key, g in groups.items():
            ref = g[0]
            if key[0] == b.geom and abs(ref.w - b.w) <= 0.05 * ref.w and abs(ref.h - b.h) <= 0.05 * ref.h:
                g.append(b)
                break
        else:
            groups[(b.geom, len(groups))].append(b)
    for key, g in groups.items():
        if len(g) < 2:
            continue
        names = ", ".join(f"'{x.name}'" for x in g[:6])
        for attr, label in (("fill", "fill colour"), ("line", "border"), ("adj", "corner/shape adjustment")):
            vals = {getattr(x, attr) for x in g}
            if len(vals) > 1 and attr != "fill":
                out.append(Finding("warning", "consistency", f"Equivalent boxes differ in {label}: {names}", idx))
        fills = Counter(x.fill for x in g)
        if len(fills) > 1 and len(g) >= 3 and fills.most_common(1)[0][1] < len(g) - 1:
            out.append(Finding("warning", "consistency", f"Equivalent boxes use {len(fills)} different fills: {names}. Fine only if colour carries meaning", idx))
        styles = {x.first_run_style() for x in g if x.text}
        if len(styles) > 1:
            out.append(Finding("warning", "consistency", f"Equivalent boxes differ in text size/weight/alignment: {names}", idx))
        props = {x.body_props() for x in g if x.text}
        if len(props) > 1:
            out.append(Finding("warning", "consistency", f"Equivalent boxes differ in internal padding or vertical alignment: {names}", idx))
        ws = [x.w for x in g]
        hs = [x.h for x in g]
        if EXACT < max(ws) - min(ws) <= 0.15 * EMU_IN or EXACT < max(hs) - min(hs) <= 0.15 * EMU_IN:
            out.append(Finding("warning", "consistency", f"Equivalent boxes are almost (not exactly) the same size: {names}", idx))
    # near-miss alignment among equivalent elements (same kind, similar size)
    cand = [i for i in content if not i.is_pic and not i.is_line and i.area_frac < 0.5 and (i.geom or i.text)]

    def similar(a, b):
        if a.geom != b.geom or bool(a.text) != bool(b.text):
            return False
        if a.text and a.first_run_style() != b.first_run_style():
            return False
        return (abs(a.w - b.w) <= 0.1 * max(a.w, b.w, 1) or abs(a.h - b.h) <= 0.1 * max(a.h, b.h, 1))

    found = []
    for x in range(len(cand)):
        for y in range(x + 1, len(cand)):
            a, b = cand[x], cand[y]
            if not similar(a, b):
                continue
            side_by_side = a.r <= b.l + EXACT or b.r <= a.l + EXACT
            stacked = a.b <= b.t + EXACT or b.b <= a.t + EXACT
            checks = []
            if side_by_side:
                checks += [("t", "tops"), ("b", "bottoms")]
            if stacked:
                checks += [("l", "left edges"), ("r", "right edges")]
            for edge, label in checks:
                d = abs(getattr(a, edge) - getattr(b, edge))
                if EXACT < d <= NEAR:
                    found.append(f"'{a.name}' / '{b.name}' {label} {inch(d)} in apart")
    if found:
        more = f" (+{len(found) - 5} more)" if len(found) > 5 else ""
        out.append(Finding("warning", "alignment", "Parallel elements almost (not exactly) aligned: " + "; ".join(found[:5]) + more, idx))


def check_colors(deck, idx, infos, out, deck_colors):
    fams = set()
    for i in infos:
        cols = []
        if i.fill:
            cols.append(i.fill)
        if i.line and i.line[0]:
            cols.append(i.line[0])
        if i.text:
            for el in i.sh._element.iter(f"{{{NS['a']}}}rPr"):
                c = hex_of(el.find("a:solidFill", NS), deck.theme_colors)
                if c:
                    cols.append(c)
        for c in cols:
            if not c:
                continue
            if not is_neutral(c):
                deck_colors[c] += 1
            f = hue_family(c)
            if f:
                fams.add(f)
        pal = deck.theme.get("palette")
        if pal:
            for c in cols:
                if c and not is_neutral(c) and min(color_dist(c, p.upper()) for p in pal) > 40:
                    out.append(Finding("warning", "palette", f"'{i.name}' uses {c}, which isn't in the theme palette", idx))
    if {"red", "green"} <= fams:
        out.append(Finding("warning", "color", "Red and green both used on this slide: OK only for labelled RAG ratings (every cell also carries a word/symbol)", idx))


def check_effects(deck, idx, infos, out):
    """Shadows (explicit, or inherited from the theme's effect styles) on non-picture shapes."""
    shadowed = []
    for i in infos:
        el = i.sh._element
        explicit = el.find(".//p:spPr/a:effectLst/a:outerShdw", NS) is not None
        ref = el.find(".//p:style/a:effectRef", NS)
        inherited = ref is not None and ref.get("idx", "0") != "0"   # some viewers ignore an empty effect list
        if (explicit or inherited) and not i.is_pic:
            shadowed.append(i.name)
    if shadowed:
        more = f" (+{len(shadowed) - 4} more)" if len(shadowed) > 4 else ""
        out.append(Finding("warning", "effects", "Shadow on " + ", ".join(f"'{n}'" for n in shadowed[:4]) + more +
                           ": remove it (shadows only on clearly external images). Inherited theme shadows: set the shape style's effectRef idx to 0", idx))


def check_bullets(deck, idx, infos, out):
    """Typed bullet characters, missing hanging indents, and orphaned last words."""
    for i in infos:
        if not i.text:
            continue
        for p in i.sh.text_frame.paragraphs:
            t = "".join(r.text for r in p.runs)
            if re.match(r"^\s*[•\-–▪●]\s", t):
                out.append(Finding("warning", "bullets", f"Typed bullet character in '{i.name}': use a real bullet with a hanging indent so wrapped lines align with the text: {short(t)!r}", idx))
                break
        # orphan check on wrapped paragraphs
        bp = i.sh.text_frame._txBody.find("a:bodyPr", NS)
        if bp is None or bp.get("wrap") == "none":
            continue
        ins = lambda k, d: int(bp.get(k, d)) / 12700
        width_pt = i.w / 12700 - ins("lIns", 91440) - ins("rIns", 91440)
        for p in i.sh.text_frame.paragraphs:
            runs = [r for r in p.runs if r.text]
            # manual line breaks (<a:br/>): only the text after the last break can orphan
            segs, cur_seg, wb, cur_wb = [], "", [], []     # wb: (word, bold) so mixed-weight lines measure right
            for child in p._p:
                tag = etree.QName(child).localname
                if tag == "br":
                    segs.append(cur_seg); cur_seg = ""; cur_wb = []
                elif tag == "r":
                    t_ = "".join(t.text or "" for t in child.iter(f"{{{NS['a']}}}t"))
                    rp = child.find("a:rPr", NS)
                    bb = rp is not None and rp.get("b") in ("1", "true")
                    joined = (cur_seg + t_)
                    cur_seg = joined
                    cur_wb = [(w_, bb if w_ in t_.split() else b0) for w_, b0 in
                              zip(joined.split(), [x[1] for x in cur_wb] + [bb] * len(joined.split()))]
            segs.append(cur_seg)
            txt = segs[-1]
            words = txt.split()
            if len(words) < 4 and len(segs) == 1:
                continue
            if len(segs) > 1 and len(words) >= 2:
                continue
            if len(segs) > 1 and len(words) < 2:
                out.append(Finding("warning", "wrapping", f"Manual line break leaves {txt.strip()!r} alone on the last line of '{i.name}'", idx))
                continue
            size, _ = effective_size(runs[0], p, i.sh, deck)
            fam = deck.resolve_font(runs[0].font.name) or "Arial"
            pPr = p._p.find("a:pPr", NS)
            indent_pt = int(pPr.get("marL", "0")) / 12700 if pPr is not None else 0
            bolds = dict(enumerate(b for _, b in cur_wb)) if len(cur_wb) == len(words) else {}
            wpt = lambda k, w: text_width_pt(w, fam, bolds.get(k, bool(runs[0].font.bold)), size)
            sp_pt = text_width_pt(" ", fam, False, size)
            lines, cur = [[]], 0.0
            for k, w in enumerate(words):
                trial = cur + (sp_pt if cur else 0) + wpt(k, w)
                if trial <= width_pt - indent_pt or not cur:
                    cur = trial; lines[-1].append(w)
                else:
                    lines.append([w]); cur = wpt(k, w)
            if len(lines) > 1 and len(lines[-1]) == 1:
                out.append(Finding("warning", "wrapping", f"Orphaned last word in '{i.name}' ({lines[-1][0]!r} alone on the last line): reword or break earlier at a phrase boundary", idx))


def text_block_bounds(deck, i):
    """Approximate top/bottom (EMU) of the visible text inside a text shape, honouring vertical anchoring."""
    tf = i.sh.text_frame
    bp = tf._txBody.find("a:bodyPr", NS)
    ins = lambda k, d: int(bp.get(k, d)) if bp is not None and bp.get(k) is not None else d
    width_pt = (i.w - ins("lIns", 91440) - ins("rIns", 91440)) / 12700
    h = 0.0
    for p in tf.paragraphs:
        runs = [r for r in p.runs if r.text]
        if not runs:
            continue
        size, _ = effective_size(runs[0], p, i.sh, deck)
        fam = deck.resolve_font(runs[0].font.name) or "Arial"
        segs, cur_s = [], ""
        for child in p._p:
            tag = etree.QName(child).localname
            if tag == "br":
                segs.append(cur_s); cur_s = ""
            elif tag == "r":
                cur_s += "".join(t.text or "" for t in child.iter(f"{{{NS['a']}}}t"))
        segs.append(cur_s)
        n = sum(wrap_lines(sg, width_pt, fam, bool(runs[0].font.bold), size) for sg in segs if sg) or 1
        h += n * size * 1.2
        pPr = p._p.find("a:pPr", NS)
        for tag in ("spcBef", "spcAft"):          # paragraph spacing (e.g. gaps between bullets)
            el = pPr.find(f"a:{tag}/a:spcPts", NS) if pPr is not None else None
            if el is not None and not (tag == "spcAft" and p is tf.paragraphs[-1]) and not (tag == "spcBef" and p is tf.paragraphs[0]):
                h += int(el.get("val", 0)) / 100
    h_emu = int(h * 12700)
    anchor = bp.get("anchor", "t") if bp is not None else "t"
    top_in, bot_in = i.t + ins("tIns", 45720), i.b - ins("bIns", 45720)
    if anchor == "b":
        return bot_in - h_emu, bot_in
    if anchor == "ctr":
        mid = (top_in + bot_in) // 2
        return mid - h_emu // 2, mid + h_emu // 2
    return top_in, top_in + h_emu


def check_balance(deck, idx, infos, roles, out):
    """White space above the content (below the headline) vs below it (above the note) should be similar."""
    hl, src = roles.get("headline"), roles.get("source")
    if not hl or any(i.is_bleed for i in infos):
        return
    _, head_bottom = text_block_bounds(deck, hl)
    # no notes: distribute between the headline and the bottom of the slide
    footer = [i for i in infos if i.text and i.t > deck.H * 0.9 and i.h < int(0.45 * EMU_IN)]   # bottom tracker, page number
    if src:
        note_top = text_block_bounds(deck, src)[0]
    elif footer:
        note_top = min(text_block_bounds(deck, i)[0] for i in footer)
    else:
        note_top = deck.H
    skip = set(id(v) for v in roles.values())
    content = [i for i in infos if id(i) not in skip and i.w > 0 and i.h > 0 and i.t >= head_bottom - EXACT
               and i.b <= note_top + EXACT and i.t > deck.H * 0.12
               and not (i.text and i.t > deck.H * 0.9 and i.h < int(0.45 * EMU_IN))]   # footer-zone navigation
    if len(content) < 1:
        return
    tops, bots = [], []
    for i in content:
        if i.text and not (i.fill or i.line) and not i.is_chart:
            t, b = text_block_bounds(deck, i)
        elif i.is_chart:
            t, b = i.t, i.b
            ml = i.sh.chart._chartSpace.find(".//c:plotArea/c:layout/c:manualLayout", NS)
            if ml is not None and ml.find("c:y", NS) is not None and ml.find("c:h", NS) is not None:
                fy, fh = float(ml.find("c:y", NS).get("val")), float(ml.find("c:h", NS).get("val"))
                plot_top = i.t + int(fy * i.h)
                plot_bottom = i.t + int((fy + fh) * i.h)
                t = plot_top - int(0.15 * EMU_IN)
                b = min(i.b, plot_bottom + int(0.6 * EMU_IN))   # tick labels + axis title below the plot
        else:
            t, b = i.t, i.b
        tops.append(t); bots.append(b)
    above, below = min(tops) - head_bottom, note_top - max(bots)
    if above < int(0.6 * EMU_IN):
        out.append(Finding("warning", "breathing-room",
                           f"Only {inch(above)} in between the headline and the content: leave at least ~0.6–0.8 in (make the content group smaller, then distribute it)", idx))
    if abs(above - below) > int(0.25 * EMU_IN):
        out.append(Finding("warning", "balance",
                           f"Content isn't vertically distributed: {inch(above)} in of space between headline and content vs {inch(below)} in between content and notes. Move the content group so the two are equal (within ~0.25 in)", idx))


def check_charts(deck, idx, infos, out):
    has_chart_or_table = False
    for i in infos:
        if i.is_table:
            has_chart_or_table = True
        if not i.is_chart:
            continue
        has_chart_or_table = True
        cs = i.sh.chart._chartSpace
        plot = cs.find(".//c:plotArea", NS)
        kinds = [etree.QName(el).localname for el in plot] if plot is not None else []
        if any(k in ("pieChart", "pie3DChart", "ofPieChart") for k in kinds):
            out.append(Finding("error", "charts", f"Pie chart in '{i.name}': use a bar or 100% stacked bar instead", idx))
        if any("3D" in k for k in kinds) or cs.find(".//c:view3D", NS) is not None and any("3D" in k for k in kinds):
            out.append(Finding("error", "charts", f"3D chart in '{i.name}'", idx))
        if any(k in ("doughnutChart", "radarChart") for k in kinds):
            out.append(Finding("warning", "charts", f"Donut/radar chart in '{i.name}': rarely the best choice (charts.md §3)", idx))
        has_title = cs.find(".//c:chart/c:title", NS) is not None
        auto_del = cs.find(".//c:chart/c:autoTitleDeleted", NS)
        titled_above = any(o.text and o is not i and o.b <= i.t + int(0.1 * EMU_IN) and i.t - o.b < int(0.6 * EMU_IN)
                           and o.l < i.r and o.r > i.l for o in infos)
        if not has_title and not titled_above:
            out.append(Finding("warning", "charts", f"Chart '{i.name}' has no built-in title: make sure a descriptive chart title sits above it", idx))
        dl_on = any(el.get("val") == "1" for el in cs.findall(".//c:dLbls/c:showVal", NS))
        for ax in cs.findall(".//c:valAx", NS):
            pos = ax.find("c:axPos", NS)
            if "scatterChart" in kinds and pos is not None and pos.get("val") in ("b", "t"):
                continue  # horizontal value axis of a scatter: not the measure axis
            deleted = ax.find("c:delete", NS)
            sparkline = "lineChart" in kinds and i.h < 1.7 * EMU_IN and all(
                c.get("val", "1") in ("1", "true") for c in cs.findall(".//c:catAx/c:delete", NS))
            if deleted is not None and deleted.get("val", "1") in ("1", "true") and not sparkline:
                out.append(Finding("error", "charts", f"Value axis hidden in '{i.name}': keep the axis line and tick labels", idx))
                continue
            ln = ax.find("c:spPr/a:ln", NS)
            if ln is not None and ln.find("a:noFill", NS) is not None:
                out.append(Finding("error", "charts", f"Value axis line is invisible in '{i.name}'", idx))
            if ax.find("c:title", NS) is None:
                out.append(Finding("warning", "charts", f"No value-axis title in '{i.name}': the unit belongs there", idx))
            mn = ax.find("c:scaling/c:min", NS)
            if mn is not None and float(mn.get("val")) != 0:
                out.append(Finding("warning", "charts", f"Value axis in '{i.name}' starts at {mn.get('val')}: allowed only if clearly labelled", idx))
            if ax.find("c:majorGridlines", NS) is not None and dl_on:
                out.append(Finding("warning", "charts", f"Gridlines AND data labels in '{i.name}': drop the gridlines", idx))
        for gw in cs.findall(".//c:gapWidth", NS):
            if int(gw.get("val")) > 100:
                out.append(Finding("warning", "charts", f"Bars in '{i.name}' are thinner than the gaps (gap width {gw.get('val')}%): use ~50–80%", idx))
                break
        for m in cs.findall(".//c:lineChart//c:ser/c:marker/c:symbol", NS):
            if m.get("val") != "none":
                out.append(Finding("warning", "charts", f"Line markers on every point in '{i.name}': use a dot only on labelled points", idx))
                break
    return has_chart_or_table


ADEC_URI = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}"


def alt_info(sh):
    """(alt text, is_decorative) for a shape."""
    c = sh._element.xpath(".//*[local-name()='cNvPr']")
    if not c:
        return "", False
    c = c[0]
    deco = any(e.get("uri") == ADEC_URI for e in c.xpath("./*[local-name()='extLst']/*[local-name()='ext']"))
    return (c.get("descr") or "").strip(), deco


def check_accessibility(deck, idx, infos, out):
    """Charts and meaningful pictures need alt text; small decorative images should be marked decorative."""
    for i in infos:
        if not (i.is_chart or i.is_pic):
            continue
        alt, deco = alt_info(i.sh)
        if alt or deco:
            continue
        if i.is_chart:
            out.append(Finding("warning", "alt-text", f"Chart '{i.name}' has no alt text: say what it shows and its main finding (finishing-standard §11)", idx))
        elif i.area_frac >= 0.01:
            out.append(Finding("warning", "alt-text", f"Picture '{i.name}' has no alt text and isn't marked decorative", idx))


def check_data_source(deck, idx, infos, roles, out):
    """A slide that shows data (a chart, or a table of numbers) needs a source or note line."""
    if roles.get("source"):
        return
    data = [i for i in infos if i.is_chart]
    for i in infos:
        if i.is_table:
            cells = [c.text for r in i.sh.table.rows for c in r.cells]
            if sum(1 for t in cells if re.search(r"\d", t)) >= 4:
                data.append(i)
    if data:
        out.append(Finding("warning", "source", f"Data on this slide ('{data[0].name}') but no source or note line: say where the numbers come from, or mark them illustrative", idx))


def check_slide(deck, idx, slide, out, deck_colors, n_slides):
    infos = [ShapeInfo(sh, box, deck) for sh, box in iter_shapes(slide.shapes)]
    # leftover placeholders
    for i in infos:
        if i.sh.is_placeholder and getattr(i.sh, "has_text_frame", False) and i.sh.has_text_frame and not i.text \
                and i.placeholder_type not in (PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.DATE):
            out.append(Finding("error", "housekeeping", f"Empty placeholder '{i.name}': delete it or fill it", idx))
    # flattened slide
    pics = [i for i in infos if i.is_pic and i.area_frac > 0.85]
    if pics and not any(i.text for i in infos):
        out.append(Finding("error", "editability", "Slide is a single full-slide image: rebuild it as native, editable content", idx))
    roles = classify(deck, idx, infos)
    fonts, numeric = check_text(deck, idx, infos, roles, out)
    check_geometry(deck, idx, infos, out)
    check_colors(deck, idx, infos, out, deck_colors)
    check_effects(deck, idx, infos, out)
    if not SPECIAL_LAYOUT_RE.search(slide.slide_layout.name or ""):   # covers and dividers are composed differently
        check_balance(deck, idx, infos, roles, out)
    check_label_vs_headline(deck, idx, roles, out)
    check_corner_marks(deck, idx, infos, roles, out)
    check_arithmetic(deck, idx, infos, out)
    check_headline_wrap(deck, idx, roles, out)
    check_centred_labels(deck, idx, infos, out)
    check_word_fit(deck, idx, infos, out)
    check_text_on_images(deck, idx, slide, out)
    check_page_number_low(deck, idx, infos, roles, out)
    check_bullets(deck, idx, infos, out)
    check_accessibility(deck, idx, infos, out)
    check_data_source(deck, idx, infos, roles, out)
    has_chart = check_charts(deck, idx, infos, out)
    if numeric >= 4 and not has_chart and not any(i.is_pic for i in infos):
        out.append(Finding("warning", "chart-vs-text", f"{numeric} numbers in text and no chart or table: if 2+ of them are being compared, show them as a chart (charts.md §1)", idx))
    bleed = any(i.is_bleed for i in infos)
    return roles, fonts, bleed


class _SpecialLayout:
    """Cover, title, section-divider and closing layouts: designed differently from content slides on purpose."""
    SPECIAL = re.compile(r"cover|section|divider|closing|end slide|thank|title[\s_-]*(slide|dark|page)|^title$", re.I)
    CONTENT = re.compile(r"only|content|text|body|chart|table|two|comparison|blank", re.I)

    def search(self, name):
        return bool(self.SPECIAL.search(name)) and not self.CONTENT.search(name)


SPECIAL_LAYOUT_RE = _SpecialLayout()


def check_recurring(deck, per_slide, out, slides=None):
    """Recurring elements should sit in the same place, with the same style, on every slide.
    Slides on a template's cover, section or closing layouts are designed differently, so they're left out."""
    special = set()
    if slides is not None:
        for idx, sl in enumerate(slides, 1):
            if SPECIAL_LAYOUT_RE.search(sl.slide_layout.name or ""):
                special.add(idx)
    for role in ("headline", "breadcrumb", "source", "page_number"):
        pos = [(idx, r[role]) for idx, r in per_slide if role in r and idx not in special]
        if len(pos) < 3:
            continue
        lefts = Counter(round(i.l / EXACT) for _, i in pos)
        tops = Counter(round(i.t / EXACT) for _, i in pos)
        styles = Counter(i.first_run_style() for _, i in pos)
        ml, mt, ms = lefts.most_common(1)[0][0], tops.most_common(1)[0][0], styles.most_common(1)[0][0]
        for idx, i in pos:
            moved = []
            if abs(round(i.l / EXACT) - ml) > 1:
                moved.append(f"left {inch(i.l)} in")
            # sources may move vertically with content height only if anchored at the bottom; flag anyway
            if abs(round(i.t / EXACT) - mt) > 1 and role != "source":
                moved.append(f"top {inch(i.t)} in")
            if moved:
                out.append(Finding("warning", "recurring", f"{role.replace('_', ' ')} is not in its usual position ({', '.join(moved)})", idx))
            if i.first_run_style() and ms and i.first_run_style()[:2] != ms[:2] and role != "page_number":
                out.append(Finding("warning", "recurring", f"{role.replace('_', ' ')} style differs from the rest of the deck: {i.first_run_style()} vs {ms}", idx))


def aggregate(findings, limit=4):
    """Collapse many findings of the same rule+level on one slide into a single summary."""
    groups = defaultdict(list)
    for f in findings:
        groups[(f.slide, f.rule, f.level)].append(f)
    out = []
    for (slide, rule, level), fs in groups.items():
        if len(fs) <= limit or slide is None:
            out.extend(fs)
        else:
            ex = " | ".join(f.msg for f in fs[:2])
            out.append(Finding(level, rule, f"{len(fs)} similar findings, e.g. {ex}", slide))
    return out


def parse_slides(spec, n):
    if not spec:
        return set(range(1, n + 1))
    out = set()
    for part in spec.split(","):
        if "-" in part:
            a, b = part.split("-")
            out.update(range(int(a), int(b) + 1))
        elif part.strip():
            out.add(int(part))
    return out


def run_color(sh, deck):
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            if r.text.strip():
                rPr = r._r.find("a:rPr", NS)
                sf = rPr.find("a:solidFill", NS) if rPr is not None else None
                return hex_of(sf, deck.theme_colors) if sf is not None else None
    return None


def check_label_vs_headline(deck, idx, roles, out):
    """A topic label / breadcrumb above the headline must look different from it (finishing-standard §3)."""
    hl, bc = roles.get("headline"), roles.get("breadcrumb")
    if not hl or not bc or bc.fill:
        return
    c1, c2 = run_color(hl.sh, deck), run_color(bc.sh, deck)
    if c1 and c2 and color_dist(c1, c2) < 30:
        out.append(Finding("warning", "label-colour",
                           f"Topic label '{bc.text[:30]}' is the same colour as the headline: put it in a box, make it grey, or use a complementary brand colour", idx))


def check_corner_marks(deck, idx, infos, roles, out):
    """Small logos/motifs in a corner sit the same distance from both nearby edges (unless they bleed off one)."""
    skip = set(id(v) for v in roles.values())
    lim, small = int(1.0 * EMU_IN), int(1.3 * EMU_IN)
    for i in infos:
        if id(i) in skip or i.w > small or i.h > small or i.w <= 0 or (i.text and not i.is_pic):
            continue
        dx = min(i.l, deck.W - i.r); dy = min(i.t, deck.H - i.b)
        if dx > lim or dy > lim:
            continue
        if dx <= int(0.02 * EMU_IN) or dy <= int(0.02 * EMU_IN):
            continue          # deliberately bleeds off an edge
        if any(j is not i and j.l <= i.l and j.r >= i.r and j.t <= i.t and j.b >= i.t
               and (j.t <= int(0.02 * EMU_IN) or j.l <= int(0.02 * EMU_IN) or deck.W - j.r <= int(0.02 * EMU_IN))
               and j.w <= small for j in infos):
            continue          # part of a mark that bleeds (e.g. the notch of a hanging bookmark)
        if abs(dx - dy) > int(0.08 * EMU_IN):
            out.append(Finding("warning", "corner-mark",
                               f"'{i.name}' sits in a corner {inch(dx)} in from the side but {inch(dy)} in from the top/bottom: use the same gap to both edges", idx))


ROUNDED = {"roundRect", "round2SameRect", "round1Rect", "round2DiagRect"}


def check_corners(deck, slides, wanted, out):
    """One corner style per deck; subtle radius (finishing-standard §6)."""
    rounded, square = [], []
    for idx, sl in enumerate(slides, 1):
        if idx not in wanted:
            continue
        for sh in sl.shapes:
            prst = sh._element.find(".//a:prstGeom", NS)
            if prst is None or not sh.width or not sh.height:
                continue
            g = prst.get("prst")
            w_in, h_in = sh.width / EMU_IN, sh.height / EMU_IN
            if g in ROUNDED:
                rounded.append(idx)
                gd = prst.find(".//a:gd", NS)
                if gd is not None and gd.get("fmla", "").startswith("val "):
                    r_in = int(gd.get("fmla")[4:]) / 100000 * min(w_in, h_in)
                    if r_in > 0.32:
                        out.append(Finding("warning", "corners", f"'{sh.name}' corner radius is about {r_in:.2f} in: keep rounding subtle (~0.1–0.15 in, max ~0.3 in)", idx))
                elif g == "roundRect" and 0.1667 * min(w_in, h_in) > 0.32:
                    out.append(Finding("warning", "corners", f"'{sh.name}' uses the default corner radius (about {0.1667 * min(w_in, h_in):.2f} in): set a subtle radius", idx))
            elif g == "rect":
                sp = sh._element.find(".//p:spPr", NS)
                filled = sp is not None and sp.find("a:solidFill", NS) is not None
                bleed = (sh.width >= 0.95 * deck.W or (sh.top <= 0 and sh.height < int(0.7 * EMU_IN))
                         or sh.left <= 0 or sh.top <= 0 or sh.left + sh.width >= deck.W or sh.top + sh.height >= deck.H)   # runs off an edge: no corners to see
                if filled and not bleed and w_in >= 0.6 and h_in >= 0.4:
                    square.append(idx)
    if rounded and square:
        out.append(Finding("warning", "corners",
                           f"Deck mixes rounded boxes (slides {', '.join(map(str, sorted(set(rounded))))}) and square boxes (slides {', '.join(map(str, sorted(set(square))))}). "
                           "Use one corner style; square boxes are fine only as cells inside a rounded container"))


TIME_RE = re.compile(r"\b([01]?\d|2[0-3])[:.]([0-5]\d)\b")
MIN_RE = re.compile(r"\b(\d{1,3})\s*(?:minutes|mins?)\b", re.I)


def check_arithmetic(deck, idx, infos, out):
    """Cheap consistency checks for derived numbers. Not a substitute for the accuracy pass (SKILL.md QA step):
    durations stated in minutes must match a gap between clock times shown on the same slide."""
    text = " \n".join(i.text for i in infos if i.text)
    times = sorted({int(h) * 60 + int(m) for h, m in TIME_RE.findall(text)})
    if len(times) >= 2:
        diffs = {b - a for a in times for b in times if b > a}
        for m in MIN_RE.findall(text):
            if int(m) not in diffs and int(m) > 5:
                out.append(Finding("warning", "arithmetic",
                                   f"'{m} minutes' doesn't match any gap between the times on this slide ({', '.join(f'{t // 60}:{t % 60:02d}' for t in times)}): check the sum, and state the range it covers", idx))


def check_page_number_low(deck, idx, infos, roles, out):
    """The page number stays low in its corner. If something (e.g. a bottom tracker) sits below the notes,
    the page number moves down with it rather than staying level with the notes."""
    pn = roles.get("page_number")
    if not pn:
        return
    below = [i for i in infos if i is not pn and i.text and i.t >= pn.b - int(0.02 * EMU_IN)]
    if below:
        out.append(Finding("warning", "page-number",
                           f"Page number sits above '{below[0].text[:25]}': keep it at the lowest baseline in the corner, even if that breaks alignment with the notes", idx))
    elif deck.H - pn.b > int(0.55 * EMU_IN):
        out.append(Finding("warning", "page-number", f"Page number is {inch(deck.H - pn.b)} in from the bottom edge: keep it low in the corner", idx))


def dump_numbers(deck, slides, wanted):
    """--numbers: every number on every slide with its context, for the accuracy pass."""
    num = re.compile(r"[£$€]?\d[\d,.]*\s*(?:%|[KMB]\b|bn\b|m\b|k\b)?")
    for idx, sl in enumerate(slides, 1):
        if idx not in wanted:
            continue
        rows = []
        for sh, _ in iter_shapes(sl.shapes):
            if getattr(sh, "has_text_frame", False) and sh.has_text_frame and num.search(sh.text_frame.text):
                rows.append(("text", " ".join(sh.text_frame.text.split())))
            if getattr(sh, "has_chart", False) and sh.has_chart:
                for pl in sh.chart.plots:
                    cats = list(pl.categories)
                    for ser in pl.series:
                        rows.append(("chart", f"{ser.name}: " + ", ".join(f"{c}={v}" for c, v in zip(cats, ser.values))))
        if rows:
            print(f"Slide {idx}")
            for k, t in rows:
                print(f"  [{k}] {t}")


def check_fill_text_pairs(deck, slides, wanted, out):
    """Each fill colour carries one text colour across the deck (finishing-standard §3): e.g. text on gold is always
    white, never white on one slide and dark on the next. Covers text in shapes, text boxes sitting on a
    filled shape, and chart data labels on coloured bars."""
    C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
    A_ = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    pairs = defaultdict(lambda: defaultdict(set))       # fill -> text colour -> slides

    def first_text_color(sh):
        rc = run_color(sh, deck) if getattr(sh, "has_text_frame", False) and sh.has_text_frame and sh.text_frame.text.strip() else None
        return rc

    for idx, sl in enumerate(slides, 1):
        if idx not in wanted:
            continue
        shapes = [(sh, box) for sh, box in iter_shapes(sl.shapes)]
        filled = []
        for sh, (l, t, w, h) in shapes:
            sp = sh._element.find(".//p:spPr", NS)
            sf = sp.find("a:solidFill", NS) if sp is not None else None
            if sf is not None and w * h < deck.W * deck.H * 0.6:
                filled.append((hex_of(sf, deck.theme_colors), l, t, w, h, sh))
        for sh, (l, t, w, h) in shapes:
            col = first_text_color(sh)
            if not col:
                continue
            own = [f for f in filled if f[5] is sh]
            if own:
                pairs[own[0][0]][col].add(idx); continue
            under = [f for f in filled if f[1] <= l + w / 2 <= f[1] + f[3] and f[2] <= t + h / 2 <= f[2] + f[4]]
            if under:
                f = min(under, key=lambda f: f[3] * f[4])
                pairs[f[0]][col].add(idx)
            if getattr(sh, "has_chart", False) and sh.has_chart:
                pass
        for sh, _ in shapes:
            if not (getattr(sh, "has_chart", False) and sh.has_chart):
                continue
            for ser in sh.chart._chartSpace.iter(f"{C}ser"):
                dl = ser.find(f"{C}dLbls")
                if dl is None or dl.find(f"{C}showVal") is None or dl.find(f"{C}showVal").get("val") != "1":
                    continue
                pos = dl.find(f"{C}dLblPos")
                if pos is not None and pos.get("val") not in ("ctr", "inEnd", "inBase"):
                    continue                                   # label sits outside the bar
                ser_col = dl.find(f".//{C}txPr//{A_}solidFill")
                ser_col = hex_of(ser_col, deck.theme_colors) if ser_col is not None else None
                over = {}
                for d in dl.findall(f"{C}dLbl"):
                    i_ = int(d.find(f"{C}idx").get("val")); sf = d.find(f".//{C}txPr//{A_}solidFill")
                    if sf is not None:
                        over[i_] = hex_of(sf, deck.theme_colors)
                for dp in ser.findall(f"{C}dPt"):
                    i_ = int(dp.find(f"{C}idx").get("val")); sf = dp.find(f"{C}spPr/{A_}solidFill")
                    if sf is None:
                        continue
                    tc = over.get(i_, ser_col)
                    if tc:
                        pairs[hex_of(sf, deck.theme_colors)][tc].add(idx)
    for fill, cols in pairs.items():
        r_, g_, b_ = (int(fill.lstrip("#")[k:k + 2], 16) / 255 for k in (0, 2, 4))
        if colorsys.rgb_to_hls(r_, g_, b_)[1] > 0.82:
            continue          # pale panels host mixed content (labels, body text, session input); rule is for solid fills
        distinct = []
        for c in cols:
            if all(color_dist(c, d) > 30 for d in distinct):
                distinct.append(c)
        if len(distinct) > 1:
            desc = "; ".join(f"{c} on slides {', '.join(map(str, sorted(cols[c])))}" for c in distinct)
            out.append(Finding("warning", "fill-text",
                               f"Text on {fill} fills uses different colours ({desc}): pick one text colour per fill and use it everywhere"))


def check_headline_wrap(deck, idx, roles, out):
    """Headlines use as much of the top line as possible; a second line of only 1-2 words should be reworded
    or abbreviated (finishing-standard §2). Also flags a manual break that leaves room on the top line."""
    hl = roles.get("headline")
    if not hl:
        return
    p = hl.sh.text_frame.paragraphs[0]
    runs = [r for r in p.runs if r.text]
    if not runs:
        return
    size, _ = effective_size(runs[0], p, hl.sh, deck)
    fam = deck.resolve_font(runs[0].font.name) or "Arial"
    bold = bool(runs[0].font.bold)
    bp = hl.sh.text_frame._txBody.find("a:bodyPr", NS)
    ins = lambda k, d: int(bp.get(k, d)) if bp is not None and bp.get(k) is not None else d
    width_pt = (hl.w - ins("lIns", 91440) - ins("rIns", 91440)) / 12700
    text = hl.sh.text_frame.text.replace("\x0b", "\n")
    parts = text.split("\n")
    if len(parts) == 2:
        first, second = parts
        nxt = second.split()[0] if second.split() else ""
        if nxt and text_width_pt(first + " " + nxt, fam, bold, size) <= width_pt * 0.97 and len(second.split()) > 2:
            out.append(Finding("warning", "headline-wrap", f"Headline breaks early: '{nxt}' would still fit on the top line; use as much of the top line as possible", idx))
        words = second.split()
    else:
        words, cur, lines = text.split(), "", [[]]
        for w in words:
            trial = (cur + " " + w).strip()
            if text_width_pt(trial, fam, bold, size) <= width_pt or not cur:
                cur = trial; lines[-1].append(w)
            else:
                lines.append([w]); cur = w
        words = lines[-1] if len(lines) > 1 else []
    if words and len(words) <= 2:
        out.append(Finding("warning", "headline-wrap", f"Headline's second line is only {' '.join(words)!r}: reword or abbreviate so it fits one line (or fills more of line two)", idx))


def check_centred_labels(deck, idx, infos, out):
    """A centre-aligned label sitting on a bar-like filled shape must be centred on that shape: the safest build
    is a text box with exactly the shape's geometry (or the shape's own text frame)."""
    bars = [i for i in infos if i.fill and not i.text and i.h <= int(0.7 * EMU_IN) and i.w >= 2 * i.h and not i.is_bleed]
    for i in infos:
        if not i.text or i.fill:
            continue
        p = i.sh.text_frame.paragraphs[0]
        if p.alignment is None or "CENTER" not in str(p.alignment):
            continue
        cy = i.t + i.h / 2
        for b in bars:
            if b.t <= cy <= b.b and b.l <= i.l + i.w / 2 <= b.r and i.h <= b.h * 1.6:
                off = abs((i.l + i.w / 2) - (b.l + b.w / 2))
                if off > int(0.08 * EMU_IN):
                    out.append(Finding("warning", "centring", f"'{i.text[:25]}' is centre-aligned but sits {inch(off)} in off the centre of the bar it labels: give the label the bar's exact geometry", idx))
                break


def _lum(rgb):
    def ch(c):
        c = c / 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = rgb
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def photo_legibility(pic, pb, tb, text_hex):
    """Look at the actual pixels under the text (images are often pre-darkened, so an overlay shape isn't the
    only way to make text legible). Returns a problem description, or None if the text is legible."""
    try:
        from PIL import Image, ImageStat
        import io
        im = Image.open(io.BytesIO(pic.image.blob)).convert("RGB")
    except Exception:
        return None
    W_, H_ = im.size
    cl, cr, ct, cb = (pic.crop_left or 0), (pic.crop_right or 0), (pic.crop_top or 0), (pic.crop_bottom or 0)
    def to_px(x, y):
        fx = (x - pb[0]) / pb[2]; fy = (y - pb[1]) / pb[3]
        return (cl + fx * (1 - cl - cr)) * W_, (ct + fy * (1 - ct - cb)) * H_
    x0, y0 = to_px(max(tb[0], pb[0]), max(tb[1], pb[1]))
    x1, y1 = to_px(min(tb[0] + tb[2], pb[0] + pb[2]), min(tb[1] + tb[3], pb[1] + pb[3]))
    if x1 - x0 < 2 or y1 - y0 < 2:
        return None
    region = im.crop((int(x0), int(y0), int(x1), int(y1))).resize((60, 30))
    L = [_lum(px) for px in (region.get_flattened_data() if hasattr(region, "get_flattened_data") else region.getdata())]
    mean = sum(L) / len(L)
    t = text_hex.lstrip("#")
    tl = _lum((int(t[0:2], 16), int(t[2:4], 16), int(t[4:6], 16)))
    contrast = (max(tl, mean) + 0.05) / (min(tl, mean) + 0.05)
    worst = min((max(tl, l) + 0.05) / (min(tl, l) + 0.05) for l in L)
    if contrast < 3:
        return f"contrast with the picture behind it is only {contrast:.1f}:1"
    if worst < 1.5 and sum(1 for l in L if (max(tl, l) + 0.05) / (min(tl, l) + 0.05) < 2) > len(L) * 0.15:
        return "parts of the picture behind it are as light (or dark) as the text"
    return None


def check_text_on_images(deck, idx, slide, out):
    """Text placed over a photo needs an overlay (a gradient fade or a semi-transparent/solid panel) between
    the picture and the text, so it stays legible (visual-identity §5b)."""
    shapes = list(iter_shapes(slide.shapes))
    order = {id(sh): k for k, (sh, _) in enumerate(shapes)}
    pics = [(sh, b) for sh, b in shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE and b[2] * b[3] > (deck.W * deck.H) * 0.04]
    def covers(b, x, y):
        return b[0] <= x <= b[0] + b[2] and b[1] <= y <= b[1] + b[3]
    for sh, b in shapes:
        if not (getattr(sh, "has_text_frame", False) and sh.has_text_frame and sh.text_frame.text.strip()):
            continue
        sp = sh._element.find(".//p:spPr", NS)
        if sp is not None and (sp.find("a:solidFill", NS) is not None or sp.find("a:gradFill", NS) is not None):
            continue                     # the text has its own background
        cx, cy = b[0] + b[2] / 2, b[1] + b[3] / 2
        for pic, pb in pics:
            if order[id(pic)] > order[id(sh)] or not covers(pb, cx, cy):
                continue
            shielded = False
            for o, ob in shapes:
                if order[id(pic)] < order[id(o)] < order[id(sh)] and covers(ob, cx, cy):
                    osp = o._element.find(".//p:spPr", NS)
                    if osp is not None and (osp.find("a:solidFill", NS) is not None or osp.find("a:gradFill", NS) is not None):
                        shielded = True; break
            if not shielded:
                tc = run_color(sh, deck)
                verdict = photo_legibility(pic, pb, b, tc) if tc else None   # unknown (inherited) colour: skip
                if verdict:
                    out.append(Finding("warning", "text-on-image",
                                       f"'{sh.text_frame.text.strip()[:30]}' on a photo: {verdict}. Add a gradient fade or a semi-transparent panel behind it (or darken/lighten the image under the text)", idx))
            break


def check_transitions(deck, slides, wanted, out):
    """At most one transition style across the deck (visual-identity §6)."""
    kinds = {}
    for idx, sl in enumerate(slides, 1):
        if idx not in wanted:
            continue
        t = sl._element.find(f"{{{NS['p']}}}transition")
        if t is None:
            kinds[idx] = None; continue
        child = [etree.QName(c).localname for c in t if etree.QName(c).localname not in ("sndAc", "extLst")]
        kinds[idx] = child[0] if child else "plain"
    used = {k for k in kinds.values() if k}
    if len(used) > 1:
        out.append(Finding("warning", "transitions", f"Several transition styles in one deck ({', '.join(sorted(used))}): use one subtle style throughout, or none"))
    elif used and any(v is None for v in kinds.values()):
        n = sum(v is None for v in kinds.values())
        out.append(Finding("warning", "transitions", f"Transitions on some slides but not on {n} others: apply the deck's one transition everywhere, or nowhere"))


def check_word_fit(deck, idx, infos, out):
    """A single word wider than its text box gets broken mid-word by the renderer ("Transp / ort"). Labels that
    don't fit inside a bar segment or box go outside it, or the shape is made wider (charts.md §5b)."""
    for i in infos:
        if not i.text or i.w <= 0:
            continue
        tf = i.sh.text_frame
        bp = tf._txBody.find("a:bodyPr", NS)
        if bp is not None and bp.get("wrap") == "none":
            continue
        ins = lambda k, d: int(bp.get(k)) if bp is not None and bp.get(k) is not None else d
        width_pt = ((i.sh.width or i.w) - ins("lIns", 91440) - ins("rIns", 91440)) / 12700   # the text frame's own width (unrotated)
        for p in tf.paragraphs:
            for r in p.runs:
                if not r.text.strip():
                    continue
                size, _ = effective_size(r, p, i.sh, deck)
                fam = deck.resolve_font(r.font.name) or "Arial"
                for w in r.text.split():
                    if text_width_pt(w, fam, bool(r.font.bold), size) > width_pt * 1.0:
                        out.append(Finding("error", "word-break", f"'{w}' is wider than its box ({round(width_pt / 72, 2)} in) and will break mid-word: widen the box, shorten the label, or put it outside the shape", idx))
                        break


def check_positions(deck, slides, wanted, out):
    """Recurring elements that do the same job should sit in the same place on every slide, so content
    doesn't jump around as the viewer moves through the deck (finishing-standard §5): chart titles, charts, and
    full-width decision/discussion bars."""
    tol, jump = int(0.08 * EMU_IN), int(0.75 * EMU_IN)
    titles, charts, bars = [], [], []
    for idx, sl in enumerate(slides, 1):
        if idx not in wanted:
            continue
        shapes = [sh for sh in sl.shapes if sh.width and sh.height]
        for sh in shapes:
            if getattr(sh, "has_chart", False) and sh.has_chart:
                charts.append((idx, sh.left, sh.top))
                above = [t for t in shapes if getattr(t, "has_text_frame", False) and t.has_text_frame
                         and t.text_frame.text.strip() and not (getattr(t, "has_chart", False) and t.has_chart)
                         and 0 <= sh.top - (t.top + t.height) <= int(0.8 * EMU_IN)
                         and t.left < sh.left + sh.width and t.left + t.width > sh.left and t.top > deck.H * 0.12]
                if above:
                    t = max(above, key=lambda t: t.top)
                    titles.append((idx, t.left, t.top))
            sp = sh._element.find(".//p:spPr", NS)
            filled = sp is not None and sp.find("a:solidFill", NS) is not None
            if (filled and getattr(sh, "has_text_frame", False) and sh.has_text_frame and sh.text_frame.text.strip()
                    and sh.width >= 0.8 * deck.W and sh.height <= int(1.3 * EMU_IN) and sh.top > deck.H * 0.5):
                bars.append((idx, sh.left, sh.top))

    def spread(items, k, what):
        vals = sorted(set(v[k] for v in items))
        if len(items) < 2 or vals[-1] - vals[0] <= tol:
            return
        # differences small enough to look like jitter, not a deliberately different layout
        pairs = [(a, b) for a in items for b in items if a[0] < b[0] and tol < abs(a[k] - b[k]) <= jump]
        if pairs:
            desc = ", ".join(f"slide {i}: {inch(v[k])} in" for i, *_ in [items[0]] for v in items for i in [v[0]])
            out.append(Finding("warning", "positions",
                               f"{what} sit at different {'heights' if k == 2 else 'x positions'} ({desc}): put elements that do the same job in the same place on every slide. A shift of about half a line on a slide with a two-line headline is expected; shorten that headline to one line to remove it"))
    spread(titles, 2, "Chart titles"); spread(titles, 1, "Chart titles")
    spread(charts, 2, "Charts")
    spread(bars, 2, "Full-width decision/discussion bars")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx")
    ap.add_argument("--theme", help="theme.json with palette / fonts / page_number_exempt")
    ap.add_argument("--slides", help="e.g. 3,5-7 (default: all)")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--numbers", action="store_true", help="list every number on each slide (text and chart data) for the accuracy pass, then exit")
    args = ap.parse_args()

    theme = json.load(open(args.theme)) if args.theme else {}
    deck = Deck(args.pptx, theme)
    slides = list(deck.prs.slides)
    wanted = parse_slides(args.slides, len(slides))
    if args.numbers:
        dump_numbers(deck, slides, wanted)
        return
    findings, per_slide = [], []
    deck_fonts, deck_colors = Counter(), Counter()
    exempt = set(theme.get("page_number_exempt", []))

    for idx, s in enumerate(slides, 1):
        if idx not in wanted:
            continue
        roles, fonts, bleed = check_slide(deck, idx, s, findings, deck_colors, len(slides))
        deck_fonts.update(fonts)
        per_slide.append((idx, roles))
        if "page_number" not in roles and idx not in (1, len(slides)) and not bleed and idx not in exempt \
                and not layout_has_slidenum(s):
            findings.append(Finding("warning", "page-number", "No page number found (required except on title, closing, divider and full-bleed slides)", idx))

    check_recurring(deck, per_slide, findings, slides)
    check_positions(deck, slides, wanted, findings)
    check_corners(deck, slides, wanted, findings)
    check_fill_text_pairs(deck, slides, wanted, findings)
    check_transitions(deck, slides, wanted, findings)

    fams = [f for f in deck_fonts if f and f != "(theme default)"]
    allowed = theme.get("fonts")
    if len(set(fams)) > 2:
        findings.append(Finding("error", "fonts", f"{len(set(fams))} fonts used (max 2): {', '.join(sorted(set(fams)))}"))
    for f in sorted(set(fams)):
        if f.lower() in BANNED_FONTS and f not in (deck.major, deck.minor):
            findings.append(Finding("error", "fonts", f"Banned font {f} is used (it may come from the template's theme fonts): switch to the deck's chosen font"))
    if allowed:
        for f in set(fams) - set(allowed):
            findings.append(Finding("warning", "fonts", f"Font {f} isn't in the theme's fonts"))
    for f in set(fams):
        path, matched = font_file(f)
        METRIC_TWINS = {("arial", "liberation sans"), ("helvetica", "liberation sans"), ("times new roman", "liberation serif"),
                        ("courier new", "liberation mono"), ("cambria", "caladea"), ("calibri", "carlito")}
        if matched and (f.lower(), matched.lower()) in METRIC_TWINS:
            continue  # metric-compatible stand-in: widths match, so fit checks and renders are reliable
        if matched and matched.lower() != f.lower():
            findings.append(Finding("warning", "fonts", f"Font '{f}' isn't installed here (renders as {matched}): text fit checks and renders are approximate; confirm on the presenting machine"))
    brand = [c for c, n in deck_colors.most_common() if n >= 2]
    # count colour *families*: tints and shades of one brand colour (same hue) count once, and greys and
    # near-white/near-black neutrals don't count (finishing-standard §3: 1-2 main colours + up to 4 accents)
    def hls(c):
        c = c.lstrip("#"); return colorsys.rgb_to_hls(*(int(c[k:k + 2], 16) / 255 for k in (0, 2, 4)))
    distinct, hues = [], []
    for c in brand:
        h_, l_, s_ = hls(c)
        if s_ < 0.18 or l_ > 0.93 or l_ < 0.1:
            continue
        if all(min(abs(h_ - u), 1 - abs(h_ - u)) > 0.035 for u in hues):
            hues.append(h_); distinct.append(c)
    if len(distinct) > 6 and not theme.get("palette"):
        findings.append(Finding("warning", "palette", f"About {len(distinct)} distinct non-neutral colours in use (max 6 + labelled RAG): {', '.join(distinct[:10])}"))

    findings = aggregate(findings)
    if args.json:
        print(json.dumps([f.as_dict() for f in findings], indent=2))
    else:
        by = defaultdict(list)
        for f in findings:
            by[f.slide].append(f)
        for k in sorted(by, key=lambda x: (x is not None, x or 0)):
            print("DECK" if k is None else f"Slide {k}")
            for f in sorted(by[k], key=lambda f: f.level != "error"):
                print(f"  {'ERROR' if f.level == 'error' else 'warn '} [{f.rule}] {f.msg}")
        e = sum(f.level == "error" for f in findings)
        w = len(findings) - e
        print(f"\n{e} error(s), {w} warning(s) across {len(wanted)} slide(s).")
    sys.exit(1 if any(f.level == "error" for f in findings) else 0)


if __name__ == "__main__":
    main()
