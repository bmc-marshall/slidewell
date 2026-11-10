"""slidekit: build helpers that bake in the Slidewell finishing standard.

Import from a build script:
    import sys; sys.path.insert(0, "<skill>/scripts")
    from slidekit import Kit
    kit = Kit(theme)          # theme: dict, see DEFAULT_THEME
    s = kit.new_slide()
    kit.frame(s, n=2, tag="Executive summary", headline="...", source="Source: ...")
    top, bottom = kit.content_zone(s)   # balanced space between headline and note

Working inside the user's own template:
    kit = Kit(theme_overrides, template="their-template.pptx")   # reads colours, fonts, headline position, footer
    s = kit.new_slide("Title only")                                # any layout by name
    kit.frame(s, n, None, headline, source)                         # headline goes into the title placeholder
    kit.fill(s, "body", None, bullets=[("Point", 0), ("Detail", 1)]) # keeps the template's bullet styles
    kit.move(placeholder, y=..., h=...)                             # reposition a placeholder without collapsing it

What it handles for you (see references/finishing-standard.md for the reasons):
- text with real line breaks, phrase-aware wrapping, no orphans or short fragments (smart_break)
- real bullets with hanging indents and brand-coloured bullet characters (bullet_list)
- shapes and connectors with NO shadows, including theme-inherited ones (no_effects)
- small open arrowheads (arrow), opaque badges (badge), evenly placed markers (markers_between)
- fixed-size, left-aligned header tag; source + page number on one bottom baseline (frame)
- fixed deck-wide anchors (anchors): content top, content bottom, decision-bar position, so recurring
  elements sit in the same place on every slide; vertically balanced blocks (place_balanced)
- chart titles aligned with the axis (chart_title); takeaways beside a chart spread down the plot (side_points)
- imagery: photos cropped to fill or placeholder frames (photo), legibility overlays (overlay), icons (icon),
  one deck-wide transition (transition)
- chart helpers: transparent chart areas, fixed plot layout, axis styling, line labels centred on lines
Fonts are measured with the real metrics of the installed font (sw_fonts.py: fontconfig, or the OS font folders), so use fonts that are
installed here or have metric-compatible stand-ins (Arial, Times New Roman, Cambria...).
"""
import sys
import subprocess
from functools import lru_cache

from lxml import etree
from PIL import ImageFont
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData  # noqa: F401  (re-exported for build scripts)
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"

FUNCTION_WORDS = {"that", "and", "to", "for", "with", "of", "in", "by", "from", "which", "while", "but", "so",
                  "the", "a", "an", "at", "on", "as", "would", "will", "is", "are", "or", "into", "than"}


def rgb(hexstr):
    hexstr = hexstr.lstrip("#")
    return RGBColor(int(hexstr[0:2], 16), int(hexstr[2:4], 16), int(hexstr[4:6], 16))


ADEC_URI = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}"
ADEC = "http://schemas.microsoft.com/office/drawing/2017/decorative"


def _cnvpr(shape):
    return shape._element.xpath(".//*[local-name()='cNvPr']")[0]


def set_alt(shape, text):
    """Set a shape's alt text (what a screen reader says; PowerPoint: right-click > View Alt Text)."""
    _cnvpr(shape).set("descr", text)
    return shape


def mark_decorative(shape):
    """Mark a shape as decorative (PowerPoint's 'Mark as decorative'), so screen readers skip it."""
    c = _cnvpr(shape)
    ext_lst = c.find(f"{{{A}}}extLst")
    if ext_lst is None:
        ext_lst = etree.SubElement(c, f"{{{A}}}extLst")
    if not any(e.get("uri") == ADEC_URI for e in ext_lst):
        ext = etree.SubElement(ext_lst, f"{{{A}}}ext"); ext.set("uri", ADEC_URI)
        d = etree.SubElement(ext, f"{{{ADEC}}}decorative", nsmap={"adec": ADEC}); d.set("val", "1")
    return shape


def describe_chart(gf):
    """Plain-language alt text for a native chart: type, categories and series."""
    ch = gf.chart
    kind = str(ch.chart_type).split(".")[-1].split(" ")[0].replace("_", " ").lower() if ch.chart_type else "chart"
    kind = kind.replace("clustered", "").replace("xy scatter", "scatter").strip() or "chart"
    try:
        cats = [str(c) for c in ch.plots[0].categories]
    except Exception:
        cats = []
    series = [s.name for p in ch.plots for s in p.series if s.name]
    bits = [f"{kind.capitalize()} chart"]
    if ch.has_title and ch.chart_title.has_text_frame and ch.chart_title.text_frame.text.strip():
        bits[0] += f": {ch.chart_title.text_frame.text.strip()}"
    if cats:
        shown = ", ".join(cats[:8]) + ("…" if len(cats) > 8 else "")
        bits.append(f"categories {shown}")
    if series:
        bits.append("series " + ", ".join(series[:5]))
    return "; ".join(bits) + "."


ICONS_ZIP = __import__("os").path.join(__import__("os").path.dirname(__import__("os").path.abspath(__file__)), "..", "assets", "icons.zip")


@lru_cache(maxsize=1)
def icon_names():
    """The names of the bundled icons (also listed in assets/icons.txt)."""
    import zipfile
    with zipfile.ZipFile(ICONS_ZIP) as z:
        return tuple(sorted(n[4:-4] for n in z.namelist() if n.startswith("png/")))


def _icon_png(name, col):
    """Recolour a bundled black icon mask (assets/icons.zip, png/<name>.png) with Pillow: no SVG tools needed."""
    import io, os, tempfile, zipfile
    from PIL import Image
    out = os.path.join(tempfile.gettempdir(), "slidewell_icons", f"{name}-{col}.png")
    if not os.path.exists(out):
        os.makedirs(os.path.dirname(out), exist_ok=True)
        with zipfile.ZipFile(ICONS_ZIP) as z:
            a = Image.open(io.BytesIO(z.read(f"png/{name}.png"))).getchannel("A")
        im = Image.new("RGBA", a.size, tuple(int(col[i:i + 2], 16) for i in (0, 2, 4)) + (0,))
        im.putalpha(a)
        im.save(out)
    return out


def template_theme(prs):
    """Read a template's own design into theme keys: slide size, margins and headline position (from the title
    placeholder), headline size, fonts and colours (from the theme), and where the footer starts. Anything the
    template doesn't define falls back to DEFAULT_THEME. Override any key by passing it in theme=."""
    from pptx.util import Emu
    EMU = 914400
    t = {"W": prs.slide_width / EMU, "H": prs.slide_height / EMU}
    master = prs.slide_master
    ns = {"a": A, "p": P}
    # colours and fonts from the theme part
    theme_part = [r.target_part for r in master.part.rels.values() if r.reltype.endswith("/theme")][0]
    th = etree.fromstring(theme_part.blob)

    def clr(name):
        el = th.find(f".//a:clrScheme/a:{name}", ns)
        if el is None:
            return None
        c = el.find("a:srgbClr", ns)
        if c is not None:
            return c.get("val")
        c = el.find("a:sysClr", ns)
        return c.get("lastClr") if c is not None else None

    dk1, dk2, a1, a2 = clr("dk1"), clr("dk2"), clr("accent1"), clr("accent2")
    main = dk2 if dk2 and dk2.upper() not in ("000000", "FFFFFF") else a1
    if main:
        t["main"] = main
        r, g, b = (int(main[i:i + 2], 16) for i in (0, 2, 4))
        t["main_pale"] = "%02X%02X%02X" % tuple(int(c + (255 - c) * 0.88) for c in (r, g, b))
    if a2 or a1:
        t["accent"] = a2 if (a2 and a2 != main) else a1
    if dk1:
        t["ink"] = dk1
    maj = th.find(".//a:fontScheme/a:majorFont/a:latin", ns)
    mnr = th.find(".//a:fontScheme/a:minorFont/a:latin", ns)
    if mnr is not None:
        t["body"] = mnr.get("typeface")
    if maj is not None:
        t["display"] = maj.get("typeface")
    # headline position and size from the title placeholder (a 'title only' layout if there is one)
    title = None
    for lay in prs.slide_layouts:
        phs = [ph for ph in lay.placeholders if ph.placeholder_format.type not in (13, 15, 16)]  # not sldNum/footer/date
        if len(phs) == 1 and phs[0].placeholder_format.type == 1:
            title = phs[0]
            break
    if title is None:
        title = next((ph for ph in master.placeholders if ph.placeholder_format.type == 1), None)
    if title is not None and title.width:
        t["M"] = round(title.left / EMU, 3)
        t["head_top"] = round(title.top / EMU, 3)
        t["head_w"] = round(title.width / EMU, 3)
    sz = master._element.find(".//p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr", ns)
    if sz is not None and sz.get("sz"):
        t["head_pt"] = int(sz.get("sz")) / 100
        t["head_bold"] = sz.get("b") == "1"
    # the footer starts at the highest master element in the bottom 15% of the slide
    tops = [sh.top / EMU for sh in master.shapes if sh.top is not None and sh.top / EMU > t["H"] * 0.85]
    if tops:
        t["footer_top"] = round(min(tops), 3)
        t["foot_top"] = round(min(tops) - 0.12 - 0.3, 3)   # the note box (0.3 in, bottom-anchored) sits just above it
    t["template_tags"] = False      # templates have their own header; no Slidewell tag unless asked
    return t


DEFAULT_THEME = dict(
    corners="square", radius=0.12,
    W=13.333, H=7.5, M=0.6,
    body="Arial", display="Arial",
    main="1F3A5F", main_pale="E4EAF1", accent="E07A2E", ink="222222", grey="7A8088", light="C6CBD1", white="FFFFFF",
    muted_fill=None, muted_text=None, muted_main=None, muted_ink="A7ADB4",
    tag_h=0.42, head_top=0.62, head_pt=26, small_pt=10, tag_labels=("Executive summary", "Contents"),
    foot_top=6.95, gap=0.45,
)


@lru_cache(maxsize=None)
def _font_path(family, bold):
    import os as _os
    sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
    from sw_fonts import font_path
    return font_path(family, bold)


@lru_cache(maxsize=None)
def _font(family, bold, pt10):
    return ImageFont.truetype(_font_path(family, bold), pt10)


class Kit:
    def __init__(self, theme=None, prs=None, template=None):
        """template: path to the user's .pptx/.potx. Its slides are removed; its masters, layouts, theme colours
        and fonts are kept and read into the theme (template_theme). theme= overrides any key."""
        t = dict(DEFAULT_THEME)
        self.template = template
        if template:
            prs = Presentation(template)
            ids = prs.slides._sldIdLst
            for sid in list(ids):
                prs.part.drop_rel(sid.rId)
                ids.remove(sid)
            t.update(template_theme(prs))
        t.update(theme or {})
        for k, v in list(t.items()):
            if isinstance(v, str) and len(v) == 6 and all(ch in "0123456789abcdefABCDEF" for ch in v):
                t[k] = rgb(v)
        # fade colours along their own hue if not given
        t["muted_fill"] = t["muted_fill"] or self._tint(t["main_pale"], 0.55)
        t["muted_main"] = t["muted_main"] or self._tint(t["main"], 0.62)
        t["muted_text"] = t["muted_text"] or t["muted_main"]
        self.T = t
        self.CW = t["W"] - 2 * t["M"]
        self.prs = prs or Presentation()
        if not template:
            self.prs.slide_width, self.prs.slide_height = Inches(t["W"]), Inches(t["H"])
            self.blank = self.prs.slide_layouts[6]
        else:
            self.blank = self.layout("blank", fallback=True)
        self.tag_w = t["M"] + max(self.width_in(l.upper(), 11, True) for l in t["tag_labels"]) + 0.45

    # ------------------------------------------------------------------ colour
    @staticmethod
    def _tint(c, amount):
        r, g, b = c[0], c[1], c[2]
        return RGBColor(int(r + (255 - r) * amount), int(g + (255 - g) * amount), int(b + (255 - b) * amount))

    def tint(self, key_or_rgb, amount):
        c = self.T[key_or_rgb] if isinstance(key_or_rgb, str) else key_or_rgb
        return self._tint(c, amount)

    # ------------------------------------------------------------------ measuring
    def width_in(self, s, pt, bold=False, font=None):
        return _font(font or self.T["body"], bold, int(pt * 10)).getlength(s) / 10 / 72

    def lines(self, s, pt, w, bold=False, font=None):
        n, cur = 0, 0.0
        sp = self.width_in(" ", pt, False, font)
        for para in s.split("\v"):
            n += 1; cur = 0.0
            for word in para.split(" "):
                tw = self.width_in(word, pt, bold, font)
                if cur and cur + sp + tw > w * 0.995:
                    n += 1; cur = tw
                else:
                    cur = cur + (sp if cur else 0) + tw
        return n

    def text_height(self, s, pt, w, bold=False, font=None, spacing=1.2):
        return self.lines(s, pt, w, bold, font) * pt * spacing / 72

    def smart_break(self, tokens, width, pt, font=None, headline=False):
        """tokens: [(word, bold)]. Return the index where a manual line break should go, or None.
        Prevents orphans (last line >= 3 words, or 2 for short text) and pulls short fragments back to a
        nearby phrase boundary (before a function word).
        headline=True: fill the first line as far as it goes. A second line of only 1-2 words is flagged for
        rewording/abbreviating; a single word is moved down with one more as a fallback."""
        sp = self.width_in(" ", pt, False, font)
        lines, cur, starts = [[]], 0.0, [0]
        for k, (t, b) in enumerate(tokens):
            tw = self.width_in(t, pt, b, font)
            if cur and cur + sp + tw > width * 0.995:      # mirror the renderer's own wrapping
                lines.append([]); starts.append(k); cur = tw
            else:
                cur = cur + (sp if cur else 0) + tw
            lines[-1].append(k)
        if len(lines) < 2:
            return None
        min_last = 2 if headline else (3 if len(tokens) >= 6 else 2)
        last_w = self.width_in(" ".join(tokens[k][0] for k in lines[-1]), pt, False, font)
        fits_last = lambda k: self.width_in(" ".join(t for t, _ in tokens[k:]), pt, any(b for _, b in tokens[k:]), font) <= width * 0.97
        # the line left above the break must stay healthy too: >= 2 words and >= 40% of the width
        # (in a paragraph of 3+ lines a short middle line looks like a hole, so require 60% there)
        prev_min = 0.6 if len(lines) >= 3 else 0.4
        prev_ok = lambda k: (k - starts[-2] >= 2 and
                             self.width_in(" ".join(t for t, _ in tokens[starts[-2]:k]), pt, False, font) >= prev_min * width)
        fits = lambda k: fits_last(k) and prev_ok(k)
        if headline:   # use as much of the top line as possible
            k = starts[-1]
            short = len(tokens) - k
            if short <= 2:
                # a 1-2 word second line: the real fix is to reword or abbreviate (to fit one line, or to fill
                # line two); flag it. As a fallback, never leave a single word: move one word down.
                print(f"slidekit: headline leaves {short} word(s) on line 2 ('{' '.join(t for t, _ in tokens[k:])}'): "
                      "reword or abbreviate so it fits one line or fills more of line two", file=sys.stderr)
                if short == 1 and prev_ok(k - 1):
                    k -= 1
            return k
        if len(lines[-1]) >= min_last:
            if last_w >= 0.4 * width:
                return None
            for k in range(starts[-1] - 1, max(starts[-2], starts[-1] - 3), -1):
                if tokens[k][0].lower() in FUNCTION_WORDS and fits(k):
                    return k
            return None
        # an orphan is worse than a slightly short middle line: if the strict test finds no break, relax it to 40%
        for pm in sorted({prev_min, 0.4}, reverse=True):
            ok = lambda k: (k - starts[-2] >= 2 and
                            self.width_in(" ".join(t for t, _ in tokens[starts[-2]:k]), pt, False, font) >= pm * width)
            for need in sorted({min_last, 2}, reverse=True):     # try 3 words on the last line, then fall back to 2
                best = None
                for k in range(starts[-1] - 1, starts[-2], -1):
                    if len(tokens) - k < need:
                        continue
                    if not fits_last(k):
                        break
                    if not ok(k):
                        continue
                    if best is None:
                        best = k
                    if tokens[k][0].lower() in FUNCTION_WORDS:
                        return k
                    if starts[-1] - k >= 4:
                        break
                if best is not None:
                    return best
        return None

    def wrapped(self, s, pt, width, bold=False, font=None, headline=False):
        if "\v" in s:
            return s
        toks = [(w, bold) for w in s.split(" ")]
        k = self.smart_break(toks, width, pt, font, headline=headline)
        return s if k is None else " ".join(w for w, _ in toks[:k]) + "\v" + " ".join(w for w, _ in toks[k:])

    # ------------------------------------------------------------------ primitives
    @staticmethod
    def no_effects(shp):
        ref = shp._element.find(f".//{{{P}}}style/{{{A}}}effectRef")
        if ref is not None:
            ref.set("idx", "0")
        return shp

    def layout(self, name, fallback=False):
        """A slide layout by name (case-insensitive; exact match first, then 'contains')."""
        lays = list(self.prs.slide_layouts)
        for test in (lambda n: n == name.lower(), lambda n: name.lower() in n):
            for lay in lays:
                if test(lay.name.lower()):
                    return lay
        if fallback:
            return min(lays, key=lambda l: len(l.placeholders))
        raise KeyError(f"No layout named {name!r}. Layouts in this template: {[l.name for l in lays]}")

    def new_slide(self, layout=None):
        """A new slide on the blank layout, or on a named template layout (e.g. "Title only", "Section")."""
        slide = self.prs.slides.add_slide(self.layout(layout) if layout else self.blank)
        if self.template:
            self._slide_number(slide)
        return slide

    def placeholder(self, slide, kind):
        """The slide's placeholder of a kind: "title", "subtitle", "body", "picture" (or a placeholder idx)."""
        types = {"title": (1, 3), "subtitle": (4,), "body": (2, 7), "picture": (18,)}
        for ph in slide.placeholders:
            pf = ph.placeholder_format
            if (isinstance(kind, int) and pf.idx == kind) or (kind in types and pf.type in types[kind]):
                return ph
        return None

    def fill(self, slide, kind, text, bullets=None):
        """Put text in a template placeholder, keeping the template's own font, size, colour and bullets.
        text: a string ("\v" for a line break), or a list of strings for one paragraph each.
        bullets: list of (text, level) pairs for the body placeholder."""
        ph = self.placeholder(slide, kind)
        if ph is None:
            raise KeyError(f"This layout has no {kind} placeholder")
        tf = ph.text_frame
        items = bullets or ([(x, 0) for x in text] if isinstance(text, list) else [(text, 0)])
        for k, (t_, lvl) in enumerate(items):
            p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
            p.level = lvl
            for j, piece in enumerate(t_.split("\v")):
                if j:
                    p.add_line_break()
                p.add_run().text = piece
        return ph

    def move(self, shape, x=None, y=None, w=None, h=None):
        """Move or resize a shape, including a template placeholder. A placeholder that inherits its position
        has no size of its own, so setting only its top would collapse it to zero width: this copies the
        inherited position first, then applies the change."""
        cur = (shape.left, shape.top, shape.width, shape.height)
        shape.left, shape.top, shape.width, shape.height = cur
        if x is not None:
            shape.left = Inches(x)
        if y is not None:
            shape.top = Inches(y)
        if w is not None:
            shape.width = Inches(w)
        if h is not None:
            shape.height = Inches(h)
        return shape

    def _slide_number(self, slide):
        """Copy the layout's slide-number placeholder onto the slide (PowerPoint's own page number field)."""
        lay_ph = next((ph for ph in slide.slide_layout.placeholders if ph.placeholder_format.type == 13), None)
        if lay_ph is None or any(ph.placeholder_format.type == 13 for ph in slide.placeholders):
            return
        n = len(self.prs.slides)
        el = etree.fromstring(
            f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr><p:cNvPr id="{900 + n}" name="Slide Number"/>'
            f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="{lay_ph.placeholder_format.idx}"/>'
            f'</p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>'
            f'<a:fld id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" type="slidenum"><a:rPr lang="en-GB"/><a:t>{n}</a:t></a:fld>'
            f'</a:p></p:txBody></p:sp>')
        slide.shapes._spTree.append(el)

    def _runs(self, p, segs, pt, font):
        import re as _re
        for seg_t, seg_b, seg_c in segs:
            seg_t = _re.sub(r"(?<=\w)-(?=\w)", "\u2011", seg_t)   # keep hyphenated words together (non-breaking hyphen)
            for j, piece in enumerate(seg_t.split("\v")):
                if j:
                    p.add_line_break()
                if not piece:
                    continue
                r = p.add_run(); r.text = piece
                r.font.size, r.font.bold, r.font.name = Pt(pt), seg_b, font or self.T["body"]
                r.font.color.rgb = seg_c

    def text(self, slide, x, y, w, h, s, pt, color=None, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             font=None, wrap=True, spacing_after=None, smart=True, headline=False):
        """s: a string, a list of strings (paragraphs), or a list of paragraphs given as [(text, bold, color), ...]."""
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = wrap; tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        paras = s if isinstance(s, list) else [s]
        for i, para in enumerate(paras):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            if spacing_after is not None:
                p.space_after = Pt(spacing_after)
            if isinstance(para, list):   # mixed-style paragraph: caller inserts any manual breaks ("\v")
                segs = [(t, b, c or color or self.T["ink"]) for t, b, c in para]
            else:
                txt = self.wrapped(para, pt, w, bold, font, headline=headline) if (smart and wrap) else para
                segs = [(txt, bold, color or self.T["ink"])]
            self._runs(p, segs, pt, font)
        return tb

    def bullet_list(self, slide, x, y, w, h, items, pt=18, color=None, bullet_color=None, gap_pt=6,
                    anchor=MSO_ANCHOR.TOP, indent=0.28, bold=None):
        """Real bullets: hanging indent, brand-coloured bullet, orphan-free wrapping.
        items: strings, or (lead, rest) pairs: the lead is the key phrase, bold in the brand colour (use for
        rules, principles, and guidelines, so the eye catches the key words first).
        bold: optional list of booleans, one per string item (measured as bold when wrapping)."""
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        for i, item in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            b_ = bool(bold[i]) if bold else False
            segs = None
            if isinstance(item, tuple):
                lead, rest = item
                toks = [(t, True) for t in lead.split(" ")] + [(t, False) for t in rest.split(" ")]
                k = self.smart_break(toks, w - indent, pt)
                words = [t for t, _ in toks]
                if k is not None:
                    words[k] = "\v" + words[k]
                nl = len(lead.split(" "))
                lead_t = " ".join(words[:nl]).replace(" \v", "\v")
                rest_t = (" " + " ".join(words[nl:])).replace(" \v", "\v")
                segs = [(lead_t, True, self.T["main"]), (rest_t, False, color or self.T["ink"])]
            else:
                item = self.wrapped(item, pt, w - indent, b_)
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", str(int(Inches(indent)))); pPr.set("indent", str(-int(Inches(indent))))
            clr = etree.SubElement(etree.SubElement(pPr, f"{{{A}}}buClr"), f"{{{A}}}srgbClr")
            clr.set("val", str(bullet_color or self.T["main"]))
            etree.SubElement(pPr, f"{{{A}}}buFont").set("typeface", "Arial")
            etree.SubElement(pPr, f"{{{A}}}buChar").set("char", "•")
            p.space_after = Pt(gap_pt)
            self._runs(p, segs or [(item, b_, color or self.T["ink"])], pt, None)
        return tb

    def lead_points(self, slide, x, y, w, items, pt=18, gap=0.22, lead_color=None, target_h=None, max_gap=None):
        """Short takeaways: bold lead-in in a brand colour, then plain text. If target_h is given, the gap
        between points grows (up to max_gap, default one line) so the block approaches that height."""
        heights = []
        for lead, rest in items:
            toks = [(t, True) for t in lead.split(" ")] + [(t, False) for t in rest.split(" ")]
            k = self.smart_break(toks, w, pt)
            n_lead = len(lead.split(" "))
            if k is not None and k > n_lead:
                rw = rest.split(" "); kk = k - n_lead
                rest_t = " " + " ".join(rw[:kk]) + "\v" + " ".join(rw[kk:])
            else:
                rest_t = " " + rest
            full = lead + rest_t
            heights.append((lead, rest_t, self.lines(full.replace("\v", "\v"), pt, w) * pt * 1.2 / 72 + 0.03))
        text_h = sum(h for *_, h in heights)
        if target_h and len(items) > 1:
            max_gap = max_gap if max_gap is not None else pt * 1.2 / 72
            gap = min(max_gap, max(gap, (target_h - text_h) / (len(items) - 1)))
        for lead, rest_t, h in heights:
            self.text(slide, x, y, w, h, [[(lead, True, lead_color or self.T["main"]), (rest_t, False, self.T["ink"])]], pt, smart=False)
            y += h + gap
        return y - gap

    def shape(self, slide, kind, x, y, w, h, fill, line=None, lw=1.0):
        b = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
        b.shadow.inherit = False
        self.no_effects(b)
        if fill is None:
            b.fill.background()
        else:
            b.fill.solid(); b.fill.fore_color.rgb = fill
        if line is None:
            b.line.fill.background()
        else:
            b.line.color.rgb = line; b.line.width = Pt(lw)
        return b

    def line(self, slide, x1, y1, x2, y2, color=None, w=1.0, dash=None, head=False):
        """Straight connector; optional small OPEN arrowhead at the end. Never shadowed."""
        a = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        a.line.color.rgb = color or self.T["grey"]; a.line.width = Pt(w)
        if dash:
            a.line.dash_style = dash
        self.no_effects(a)
        if head:
            te = etree.SubElement(a.line._get_or_add_ln(), f"{{{A}}}tailEnd")
            te.set("type", "arrow"); te.set("w", "sm"); te.set("len", "sm")
        return a

    def arrow(self, slide, x1, y1, x2, y2, color=None, w=1.0):
        return self.line(slide, x1, y1, x2, y2, color, w, head=True)

    def two_tone_arrow(self, slide, x1, y1, x_split, y_split, x2, y2, color1, color2, w=1.25):
        """One straight arrow that changes colour where it crosses a background. The split point must lie
        on the straight line from start to end, so both segments share one angle."""
        self.line(slide, x1, y1, x_split, y_split, color1, w)
        self.line(slide, x_split, y_split, x2, y2, color2, w, head=True)

    def badge(self, slide, cx, cy, d, label, pt=14, on=True, ring_color=None):
        """Opaque numbered badge: white-filled outer ring, coloured inner disc on top (hides anything beneath)."""
        col = ring_color or (self.T["main"] if on else self.T["muted_main"])
        self.shape(slide, MSO_SHAPE.OVAL, cx - d / 2, cy - d / 2, d, d, self.T["white"], line=col, lw=1.5)
        inner = d * 0.76
        c = self.shape(slide, MSO_SHAPE.OVAL, cx - inner / 2, cy - inner / 2, inner, inner, col)
        tf = c.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        if label:
            r = p.add_run(); r.text = label
            r.font.size, r.font.bold, r.font.name = Pt(pt), True, self.T["body"]; r.font.color.rgb = self.T["white"]

    def connector_word(self, slide, cx, cy, word="or", d=0.62, pt=16, fill=None):
        """'or', 'vs', 'and' between two things: ONE solid circle (not the double-ring badge, which is for
        numbers, dates and steps), centred vertically on the things either side of it."""
        c = self.shape(slide, MSO_SHAPE.OVAL, cx - d / 2, cy - d / 2, d, d, fill or self.T["main"])
        tf = c.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = word
        r.font.size, r.font.bold, r.font.name = Pt(pt), True, self.T["body"]; r.font.color.rgb = self.T["white"]
        return c

    def box(self, slide, x, y, w, h, fill, corners=None, line=None, lw=1.0, r=None):
        """Every box in the deck goes through here so the corner style is uniform (finishing-standard §6).
        corners: None = the theme's style ("square" or "rounded"); "top" / "bottom" round only that pair, so a
        header box (rounded top) and its content box (rounded bottom) meet on a straight, parallel edge."""
        style = self.T.get("corners", "square")
        r = self.T.get("radius", 0.12) if r is None else r
        if style == "square" and corners is None:
            return self.shape(slide, MSO_SHAPE.RECTANGLE, x, y, w, h, fill, line=line, lw=lw)
        kind = {None: MSO_SHAPE.ROUNDED_RECTANGLE, "all": MSO_SHAPE.ROUNDED_RECTANGLE,
                "top": MSO_SHAPE.ROUND_2_SAME_RECTANGLE, "bottom": MSO_SHAPE.ROUND_2_SAME_RECTANGLE}[corners]
        shp = self.shape(slide, kind, x, y, w, h, fill, line=line, lw=lw)
        a = min(0.5, r / min(w, h))
        if corners == "bottom":
            shp.adjustments[0] = 0; shp.adjustments[1] = a
        elif corners == "top":
            shp.adjustments[0] = a; shp.adjustments[1] = 0
        else:
            shp.adjustments[0] = a
        return shp

    def vertical_label(self, slide, x_edge, cy, s, pt=10, color=None, side="left", bold=True):
        """Small rotated label running along the side of an outlined element (e.g. RECOMMENDED beside an
        outlined row), centred vertically on it. side="left": reads bottom-to-top, ending at x_edge;
        side="right": reads top-to-bottom, starting at x_edge. Use instead of a key the reader must decode."""
        w = self.width_in(s, pt, bold) + 0.1
        h = pt * 1.2 / 72
        cx = x_edge - h / 2 if side == "left" else x_edge + h / 2
        tb = self.text(slide, cx - w / 2, cy - h / 2, w, h, s, pt, color, bold=bold, align=PP_ALIGN.CENTER,
                       anchor=MSO_ANCHOR.MIDDLE, smart=False)
        tb.rotation = 270 if side == "left" else 90
        return tb

    def distribute_x(self, widths, x0=0.0, x1=None):
        """Left x of each item so the gaps between items AND between the outer items and x0/x1 (default: the
        slide edges) are all equal; PowerPoint's 'distribute horizontally' with an off-slide shape at each edge."""
        x1 = self.T["W"] if x1 is None else x1
        gap = (x1 - x0 - sum(widths)) / (len(widths) + 1)
        xs, x = [], x0 + gap
        for w in widths:
            xs.append(x); x += w + gap
        return xs

    def animate_on_click(self, slide, shapes, dim_previous=False, dim_to=0.3, dur_ms=400):
        """Native PowerPoint animation: each shape (or group of shapes) fades in on its own click, in order.
        shapes: list of shapes or lists of shapes (a list appears on the same click).
        dim_previous: on each click, the items already shown fade to dim_to opacity, so attention moves to
        the new one. Only use when the user asked for native animation (visual-identity §6); the default is
        builds as successive slides. Renders and PDFs show the final state, so check the sequence in PowerPoint."""
        P_ = "http://schemas.openxmlformats.org/presentationml/2006/main"
        groups = [g if isinstance(g, (list, tuple)) else [g] for g in shapes]
        ids = iter(range(3, 10000))
        def fade_in(spid):
            a, b = next(ids), next(ids)
            return (f'<p:set><p:cBhvr><p:cTn id="{a}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
                    f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
                    f'</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                    f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{b}" dur="{dur_ms}"/>'
                    f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>')
        def dim(spid):
            a = next(ids)
            return (f'<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="{a}" dur="{dur_ms}" fill="hold"/>'
                    f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst>'
                    f'</p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:fltVal val="1"/></p:val></p:tav>'
                    f'<p:tav tm="100000"><p:val><p:fltVal val="{dim_to}"/></p:val></p:tav></p:tavLst></p:anim>')
        clicks = []
        for gi, g in enumerate(groups):
            effects = []
            for k, shp in enumerate(g):
                c = next(ids)
                effects.append(f'<p:par><p:cTn id="{c}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" '
                               f'grpId="0" nodeType="{"clickEffect" if k == 0 else "withEffect"}"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
                               f'<p:childTnLst>{fade_in(shp.shape_id)}</p:childTnLst></p:cTn></p:par>')
            if dim_previous and gi:
                for prev in groups[gi - 1]:
                    c = next(ids)
                    effects.append(f'<p:par><p:cTn id="{c}" presetID="9" presetClass="emph" presetSubtype="0" fill="hold" '
                                   f'grpId="1" nodeType="withEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
                                   f'<p:childTnLst>{dim(prev.shape_id)}</p:childTnLst></p:cTn></p:par>')
            o, i_ = next(ids), next(ids)
            clicks.append(f'<p:par><p:cTn id="{o}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
                          f'<p:par><p:cTn id="{i_}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                          + "".join(effects) + '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
        bld = "".join(f'<p:bldP spid="{shp.shape_id}" grpId="0"/>' for g in groups for shp in g)
        if dim_previous:
            bld += "".join(f'<p:bldP spid="{shp.shape_id}" grpId="1"/>' for g in groups[:-1] for shp in g)
        xml = (f'<p:timing xmlns:p="{P_}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
               f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
               + "".join(clicks) +
               '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
               '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
               f'</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>{bld}</p:bldLst></p:timing>')
        el = etree.fromstring(xml)
        sld = slide._element
        old = sld.find(f"{{{P_}}}timing")
        if old is not None:
            sld.remove(old)
        ext = sld.find(f"{{{P_}}}extLst")
        (ext.addprevious(el) if ext is not None else sld.append(el))

    # ------------------------------------------------------------------ imagery (visual-identity §5b)
    def overlay(self, slide, x, y, w, h, color=None, mode="gradient", opacity=0.8, angle=0, stops=None, corners=None):
        """Overlay that keeps text legible on a photo. mode="gradient": colour fades from `opacity` at the start
        (angle 0 = left edge; 90 = top edge) to clear, so the text sits on the dark end and the subject stays
        visible at the clear end. mode="panel": a flat semi-transparent panel. Place it after the picture and
        before the text. corners: match the photo underneath when it has rounded corners ("all", or "bottom" /
        "top" for a panel flush with that edge of the photo); None = square."""
        col = str(color or self.T["ink"])
        if corners:
            shp = self.box(slide, x, y, w, h, None, corners=corners)
        else:
            shp = self.shape(slide, MSO_SHAPE.RECTANGLE, x, y, w, h, None)
        spPr = shp._element.spPr
        for el in spPr.findall(f"{{{A}}}noFill"):
            spPr.remove(el)
        ln = spPr.find(f"{{{A}}}ln")
        if mode == "panel":
            fill = etree.fromstring(f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="{col}"><a:alpha val="{int(opacity * 100000)}"/></a:srgbClr></a:solidFill>')
        else:
            stops = stops or [(0, opacity), (0.55, opacity * 0.55), (1, 0)]
            gs = "".join(f'<a:gs pos="{int(p_ * 100000)}"><a:srgbClr val="{col}"><a:alpha val="{int(a_ * 100000)}"/></a:srgbClr></a:gs>' for p_, a_ in stops)
            fill = etree.fromstring(f'<a:gradFill xmlns:a="{A}" rotWithShape="1"><a:gsLst>{gs}</a:gsLst><a:lin ang="{int(angle * 60000)}" scaled="0"/></a:gradFill>')
        (ln.addprevious(fill) if ln is not None else spPr.append(fill))
        shp.name = f"Overlay {shp.shape_id}"
        return shp

    def photo(self, slide, x, y, w, h, path=None, label="Photo", crop="rect", focus=(0.5, 0.5), link=None, label_at="bottom"):
        """A photo cropped to fill the frame (like 'fill' in PowerPoint), or, with no file, a placeholder frame
        of exactly the same size that says what photo belongs there, for the user to replace with their picture.
        crop="circle" crops to a circle. focus: which part of the image to keep when cropping (0-1, 0-1).
        link: the page of a suggested free photo (e.g. an Unsplash photo URL). The placeholder shows it as a short
        clickable link under the label and keeps the full URL in its alt text (visual-identity §4c).
        label_at: "bottom" (default) or "top" for large frames; use "top" when an overlay panel or a motif covers the bottom."""
        T = self.T
        bleeds = x <= 0.01 or y <= 0.01 or x + w >= T["W"] - 0.01 or y + h >= T["H"] - 0.01
        rounded = crop == "rect" and T.get("corners") == "rounded" and not bleeds   # photos share the deck's corners
        if path:
            from PIL import Image
            iw, ih = Image.open(path).size
            pic = slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
            set_alt(pic, label if label != "Photo" else "Photo")
            fr, ir = w / h, iw / ih
            if ir > fr:      # image wider than frame: crop the sides
                keep = fr / ir; spare = 1 - keep
                pic.crop_left = spare * focus[0]; pic.crop_right = spare * (1 - focus[0])
            else:
                keep = ir / fr; spare = 1 - keep
                pic.crop_top = spare * focus[1]; pic.crop_bottom = spare * (1 - focus[1])
            geom = pic._element.spPr.find(f"{{{A}}}prstGeom")
            if crop == "circle":
                geom.set("prst", "ellipse")
            elif rounded:
                geom.set("prst", "roundRect")
                av = geom.find(f"{{{A}}}avLst")
                if av is None:
                    av = etree.SubElement(geom, f"{{{A}}}avLst")
                gd = etree.SubElement(av, f"{{{A}}}gd"); gd.set("name", "adj")
                gd.set("fmla", f"val {int(min(0.5, T.get('radius', 0.12) / min(w, h)) * 100000)}")
            return pic
        if rounded:
            ph = self.box(slide, x, y, w, h, self.tint(self.T["grey"], 0.82))
        else:
            kind = MSO_SHAPE.OVAL if crop == "circle" else MSO_SHAPE.RECTANGLE
            ph = self.shape(slide, kind, x, y, w, h, self.tint(self.T["grey"], 0.82))
        ph.name = f"Photo placeholder {ph.shape_id}"
        set_alt(ph, f"Placeholder for a photo: {label}")
        tf = ph.text_frame; tf.word_wrap = True
        big = w > 4 and h > 3
        # on large frames the label sits bottom-right, where it won't sit under the slide's own text or overlay
        tf.vertical_anchor = (MSO_ANCHOR.TOP if label_at == "top" else MSO_ANCHOR.BOTTOM) if big else MSO_ANCHOR.MIDDLE
        if big:
            tf.margin_right = tf.margin_bottom = tf.margin_top = Inches(0.25)
            tf.margin_left = Inches(max(0.1, w - (4.2 if link else 3.6)))
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT if big else PP_ALIGN.CENTER
        if min(w, h) < 1.0:
            return ph          # too small for a readable label (e.g. a presenter chip): the frame alone says "photo here"
        r = p.add_run(); r.text = f"[{label}]"
        r.font.size, r.font.name, r.font.italic = Pt(12), self.T["body"], False
        r.font.color.rgb = self.T["grey"]
        if link:
            ph._element.nvSpPr.cNvPr.set("descr", f"Suggested photo: {link}")
            short = link.split("://", 1)[-1].replace("www.", "")
            if "unsplash.com/photos/" in short:          # Unsplash pages resolve from the id alone: show the short form
                short = "unsplash.com/photos/" + short.rstrip("/").split("/")[-1][-11:]   # ids are the last 11 characters
            p2 = tf.add_paragraph(); p2.alignment = p.alignment
            r2 = p2.add_run(); r2.text = f"Suggested: {short}"
            r2.font.size, r2.font.name, r2.font.italic = Pt(11), self.T["body"], False
            r2.font.color.rgb = self.T["grey"]
            r2.hyperlink.address = link
        return ph

    def icon(self, slide, name, x, y, size, color=None, circle=None, pad=0.22):
        """An icon from the bundled open-licence set (assets/icons.zip, names in assets/icons.txt; Lucide, ISC licence), recoloured and placed
        as a crisp PNG. circle: a fill colour to set the icon (in white) inside a solid circle. One icon set, one
        style, one colour per deck (visual-identity §4b)."""
        if name not in icon_names():
            raise FileNotFoundError(f"No icon '{name}'. The names are listed in assets/icons.txt (or call icon_names())")
        col = "FFFFFF" if circle is not None else str(color or self.T["main"])
        out = _icon_png(name, col)
        if circle is not None:
            self.shape(slide, MSO_SHAPE.OVAL, x, y, size, size, circle)
            inner = size * (1 - 2 * pad)
            pic = slide.shapes.add_picture(out, Inches(x + size * pad), Inches(y + size * pad), Inches(inner), Inches(inner))
        else:
            pic = slide.shapes.add_picture(out, Inches(x), Inches(y), Inches(size), Inches(size))
        pic.name = f"Icon {name}"
        mark_decorative(pic)   # icons support a label that carries the meaning, so screen readers skip them
        return pic

    def transition(self, slide, kind="fade"):
        """One subtle transition for the whole deck (visual-identity §6): fade by default; push only for a
        deck that moves strictly forward through stages."""
        el = etree.fromstring(f'<p:transition xmlns:p="{P}" spd="med"><p:{kind}/></p:transition>')
        sld = slide._element
        old = sld.find(f"{{{P}}}transition")
        if old is not None:
            sld.remove(old)
        timing = sld.find(f"{{{P}}}timing")
        if timing is not None:
            timing.addprevious(el)
        else:
            ext = sld.find(f"{{{P}}}extLst")
            (ext.addprevious(el) if ext is not None else sld.append(el))

    def corner_xy(self, w, h, corner="tr", inset=0.3):
        """Position for a logo or motif in a corner: the SAME gap to both slide edges."""
        x = inset if corner[1] == "l" else self.T["W"] - inset - w
        y = inset if corner[0] == "t" else self.T["H"] - inset - h
        return x, y

    def markers_between(self, slide, cx, gaps, kind="triangle", color=None, w=0.22, h=0.16):
        """Place a downward marker exactly halfway in each (top, bottom) gap, measured from visible edges."""
        for top, bottom in gaps:
            y = (top + bottom) / 2 - h / 2
            m = self.shape(slide, MSO_SHAPE.ISOSCELES_TRIANGLE, cx - w / 2, y, w, h, color or self.T["light"])
            m.rotation = 180

    def bar(self, slide, y, s, pt=18, h=0.85, fill=None):
        """Full-width decision/discussion bar with orphan-free wrapping."""
        b = self.shape(slide, MSO_SHAPE.RECTANGLE, self.T["M"], y, self.CW, h, fill or self.T["main"])
        tf = b.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Inches(0.3)
        self._runs(tf.paragraphs[0], [(self.wrapped(s, pt, self.CW - 0.6, True), True, self.T["white"])], pt, None)
        return b

    # ------------------------------------------------------------------ recurring elements
    def header_tag(self, slide, label, fill=None):
        """Fixed-size header box on every slide, flush with the top-left edge, label left-aligned at the margin."""
        self.shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, self.tag_w, self.T["tag_h"], fill or self.T["main"])
        self.text(slide, self.T["M"], 0, self.tag_w - self.T["M"], self.T["tag_h"], label.upper(), 11,
                  self.T["white"], bold=True, anchor=MSO_ANCHOR.MIDDLE, wrap=False, smart=False)

    def frame(self, slide, n, tag, headline, source=None, head_color=None):
        """Headline, header tag, source note and page number, in the same place on every slide.
        In a template, the headline goes into the layout's title placeholder (keeping the template's font, size
        and colour), the template's own slide-number field is used, and the tag is drawn only if the theme sets
        template_tags=True."""
        T = self.T
        title_ph = self.placeholder(slide, "title") if self.template else None
        if tag and (not self.template or T.get("template_tags")):
            self.header_tag(slide, tag)
        hw = T.get("head_w", self.CW)
        hb = T.get("head_bold", True)
        if title_ph is not None:
            self.fill(slide, "title", self.wrapped(headline, T["head_pt"], hw * 0.98, hb, font=T["display"], headline=True))
            self._head_bottom = T["head_top"] + self.text_height(self.wrapped(headline, T["head_pt"], hw * 0.98, hb, font=T["display"], headline=True),
                                                                 T["head_pt"], hw, hb, font=T["display"], spacing=1.15)
            if source:
                self.text(slide, T["M"], T["foot_top"], self.CW - 1.0, 0.3, source, T["small_pt"], T["grey"], anchor=MSO_ANCHOR.BOTTOM)
            self._src_lines = self.lines(source, T["small_pt"], self.CW - 1.0) if source else 0
            if n and not any(ph.placeholder_format.type == 13 for ph in slide.placeholders):
                self.text(slide, T["W"] - T["M"] - 0.6, T["foot_top"], 0.6, 0.3, str(n), T["small_pt"], T["grey"],
                          align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.BOTTOM, smart=False)
            return
        self.text(slide, T["M"], T["head_top"], self.CW, 0.95, headline, T["head_pt"], head_color or T["main"], bold=True, headline=True)
        self._src_lines = self.lines(source, T["small_pt"], self.CW - 1.0) if source else 0
        if source:
            self.text(slide, T["M"], T["foot_top"], self.CW - 1.0, 0.3, source, T["small_pt"], T["grey"], anchor=MSO_ANCHOR.BOTTOM)
        if n:
            # the page number stays low in the corner; it shares the note's baseline only when the note is the
            # lowest element (page_num_top moves it down when a bottom tracker pushes the notes up)
            self.text(slide, T["W"] - T["M"] - 0.6, T.get("page_num_top") or T["foot_top"], 0.6, 0.3, str(n), T["small_pt"], T["grey"],
                      align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.BOTTOM, smart=False)
        self._head_bottom = T["head_top"] + self.text_height(self.wrapped(headline, T["head_pt"], self.CW, True, headline=True), T["head_pt"], self.CW, True)

    def content_zone(self, slide=None, source_lines=None):
        """(top, bottom) of the content area with equal breathing room below the headline and above the notes.
        With no notes on the slide, the bottom reference is the top of the footer (a bottom tracker, or the
        page number). source_lines=None uses what frame() was given on this slide."""
        T = self.T
        if source_lines is None:
            source_lines = getattr(self, "_src_lines", 1)
        if source_lines:
            note_top = T["foot_top"] + 0.3 - source_lines * T["small_pt"] * 1.2 / 72
        else:
            pn_top = (T.get("page_num_top") or T["foot_top"]) + 0.3 - T["small_pt"] * 1.2 / 72
            note_top = min(T.get("footer_top") or pn_top, pn_top)
        g = T["gap"]
        return self._head_bottom + g, note_top - g

    # ------------------------------------------------------------------ fixed anchors (same spot on every slide)
    def anchors(self, source_lines=1):
        """Deck-wide fixed positions, so recurring elements don't jump between slides:
        content_top: where chart titles, panel headers and table headers start on every content slide. It is set
          by a TWO-line headline plus breathing room, so it is the same whether a headline takes one line or two.
        content_bottom: just above the source note.
        bar_y: top of the decision/discussion bar, which always sits at the bottom of the content zone."""
        T = self.T
        head_2 = T["head_top"] + 2 * T["head_pt"] * 1.2 / 72
        top = T.get("content_top") or round(head_2 + T.get("head_gap", 0.32), 2)
        note_top = T["foot_top"] + 0.3 - source_lines * T["small_pt"] * 1.2 / 72
        bottom = note_top - T.get("foot_gap", 0.12)
        # a filled bar needs more visible clearance above the note than a chart (whose axis labels have padding)
        return top, bottom, bottom - T.get("bar_h", 0.85) - T.get("bar_gap", 0.18)

    def chart_title(self, slide, chart_x, y, w, s, left_axis_title=True, pt=18):
        """Chart title aligned with the axis (the tick labels' left edge), not with the rotated axis title."""
        dx = (14 * 1.2 / 72 + 0.12) if left_axis_title else 0
        return self.chart_header(slide, chart_x + dx, y, w - dx, s, pt)

    def side_points(self, slide, x, w, plot_top, plot_bottom, items, pt=18, lead_color=None):
        """Takeaways beside a chart: the first point starts level with the top of the plotted data; gaps are
        equal and grow to spread the points down the plot, between 1 and 2.5 lines each. If even 2.5-line gaps
        can't fill the plot, the block stays top-aligned with the data rather than spreading further."""
        line = pt * 1.2 / 72
        return self.lead_points(slide, x, plot_top + 0.04, w, items, pt=pt, lead_color=lead_color,
                                gap=line, target_h=plot_bottom - plot_top - 0.04, max_gap=2.5 * line)

    def action_bar(self, slide, s, pt=18, fill=None, source_lines=1):
        """Decision/discussion bar in the same position on every slide that has one."""
        return self.bar(slide, self.anchors(source_lines)[2], s, pt=pt, fill=fill)

    def place_balanced(self, block_h, slide=None, source_lines=None):
        """Top y for a content block of height block_h, centred in the content zone (equal space above/below)."""
        top, bottom = self.content_zone(slide, source_lines)
        gap = (bottom - top - block_h) / 2 + self.T["gap"]          # visible space above (and below) the block
        if gap < self.T.get("min_gap", 0.7) - 0.01:
            print(f"slidekit: content block {block_h:.2f} in leaves only {gap:.2f} in above and below it "
                  f"(want >= {self.T.get("min_gap", 0.7)} in): make the block smaller", file=sys.stderr)
        return top + max(0, (bottom - top - block_h) / 2)

    # ------------------------------------------------------------------ charts
    def style_chart(self, ch, pt=14):
        ch.has_title = False; ch.has_legend = False
        ch.font.name, ch.font.size = self.T["body"], Pt(pt); ch.font.color.rgb = self.T["ink"]

    def axis(self, ax, title=None, pt=14):
        ax.format.line.color.rgb = self.T["grey"]; ax.format.line.width = Pt(1)
        ax.tick_labels.font.size = Pt(pt)
        if title:
            ax.has_title = True
            tf = ax.axis_title.text_frame; tf.text = title
            r = tf.paragraphs[0].runs[0]
            r.font.size, r.font.bold, r.font.name = Pt(pt), False, self.T["body"]; r.font.color.rgb = self.T["grey"]

    @staticmethod
    def plot_layout(chart, x, y, w, h):
        pa = chart._chartSpace.find(f".//{{{C}}}plotArea")
        lay = pa.find(f"{{{C}}}layout")
        if lay is None:
            lay = etree.Element(f"{{{C}}}layout"); pa.insert(0, lay)
        for ch in list(lay):
            lay.remove(ch)
        ml = etree.SubElement(lay, f"{{{C}}}manualLayout")
        for tag, val in (("layoutTarget", "inner"), ("xMode", "edge"), ("yMode", "edge"), ("x", x), ("y", y), ("w", w), ("h", h)):
            etree.SubElement(ml, f"{{{C}}}{tag}").set("val", str(val))

    @staticmethod
    def transparent_chart(chart):
        cs = chart._chartSpace
        for parent in (cs, cs.find(f".//{{{C}}}plotArea")):
            spPr = parent.find(f"{{{C}}}spPr")
            if spPr is None:
                spPr = etree.SubElement(parent, f"{{{C}}}spPr")
                if parent is cs:
                    cs.remove(spPr); cs.insert(cs.index(cs.find(f"{{{C}}}chart")) + 1, spPr)
            for c_ in list(spPr):
                spPr.remove(c_)
            etree.SubElement(spPr, f"{{{A}}}noFill")
            ln = etree.SubElement(spPr, f"{{{A}}}ln"); etree.SubElement(ln, f"{{{A}}}noFill")

    def line_label(self, slide, x, y_line, s, pt=16, color=None, bold=True, w=2.0):
        """Label at the end of a line: box vertically centred on the line, text vertically centred in the box."""
        h = pt * 1.2 / 72 + 0.06
        return self.text(slide, x, y_line - h / 2, w, h, s, pt, color, bold=bold, anchor=MSO_ANCHOR.MIDDLE, smart=False)

    def chart_header(self, slide, x, y, w, s, pt=18, align=PP_ALIGN.LEFT):
        """Chart titles and panel labels: highlight text in the main brand colour."""
        return self.text(slide, x, y, w, 0.45, s, pt, self.T["main"], bold=True, align=align)

    # ------------------------------------------------------------------ finish
    def alt(self, shape, text):
        """Set alt text on any shape; a chart's alt text should say what the chart shows and its main finding."""
        return set_alt(shape, text)

    def save(self, path, author="", title=None):
        """Save the deck. Clears python-pptx's default document properties (its template is 'last modified by'
        the library's author), so a client deck carries only the author you pass, if any."""
        cp = self.prs.core_properties
        cp.author = author; cp.last_modified_by = author
        cp.title = title if title is not None else (cp.title if self.template else "")
        cp.revision = 1
        for f in ("subject", "keywords", "comments", "category"):
            if not self.template:
                setattr(cp, f, "")
        for sl in self.prs.slides:
            for ph in list(sl.placeholders):     # unused template placeholders show "Click to add text"
                if ph.placeholder_format.type not in (13, 15, 16) and ph.has_text_frame and not ph.text_frame.text.strip():
                    ph._element.getparent().remove(ph._element)
            for shp in sl.shapes:
                self.no_effects(shp)
                # every chart gets alt text: an explicit kit.alt() wins; otherwise describe it from its data
                if getattr(shp, "has_chart", False) and shp.has_chart and not _cnvpr(shp).get("descr"):
                    set_alt(shp, describe_chart(shp))
        self.prs.save(path)
        return path
