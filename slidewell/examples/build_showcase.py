"""Showcase: five standalone slides, each in its own (fictional) client's brand, showing the range of
Slidewell's slidecraft: bold-number pillars, a pictogram, a scenario chart, market-sizing bubbles, and
organisation-design options. Every organisation, person and figure is invented.

Run: python build_showcase.py  ->  showcase.pptx next to this script.
"""
import math
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[0] / "scripts"))
from lxml import etree  # noqa: E402
from slidekit import Kit, CategoryChartData, rgb  # noqa: E402
from pptx.enum.chart import XL_CHART_TYPE  # noqa: E402
from pptx.enum.dml import MSO_LINE_DASH_STYLE  # noqa: E402
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE  # noqa: E402
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
kit = Kit(dict(body="Arial", main="1F2A44", accent="E63946", ink="22262B", grey="6B7280", light="D1D5DB"))
T, W, H = kit.T, kit.T["W"], kit.T["H"]
WHITE = rgb("FFFFFF")
C, L, R = PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.RIGHT
MID, TOP_, BOT = MSO_ANCHOR.MIDDLE, MSO_ANCHOR.TOP, MSO_ANCHOR.BOTTOM


def txt(s, x, y, w, h, t, pt, color, bold=False, align=L, anchor=TOP_, smart=True):
    return kit.text(s, x, y, w, h, t, pt, color, bold=bold, align=align, anchor=anchor, smart=smart)


def freeform(s, pts, fill=None, line=None, lw=1.0, closed=True):
    fb = s.shapes.build_freeform(Inches(pts[0][0]), Inches(pts[0][1]), scale=1.0)
    fb.add_line_segments([(Inches(x), Inches(y)) for x, y in pts[1:]], close=closed)
    shp = fb.convert_to_shape()
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line; shp.line.width = Pt(lw)
    kit.no_effects(shp)
    return shp


def curved_arrow(s, x1, y1, x2, y2, color, w=2.5):
    a = s.shapes.add_connector(MSO_CONNECTOR.CURVE, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    a.line.color.rgb = color; a.line.width = Pt(w)
    kit.no_effects(a)
    te = etree.SubElement(a.line._get_or_add_ln(), f"{{{A}}}tailEnd")
    te.set("type", "arrow"); te.set("w", "med"); te.set("len", "med")
    return a


def page_no(s, n, color):
    txt(s, W - 0.6 - 0.6, 6.95, 0.6, 0.3, str(n), 10, color, align=R, anchor=BOT, smart=False)


def source(s, t, color, y=6.95):
    txt(s, 0.6, y, 10.5, 0.3, t, 10, color, anchor=BOT)


# =============================================================================================== 1. pillars
# Brand: Kestrel, an (invented) energy-drink company. Deep blue with a signal red; a slanted brand band.
NAVY, RED, PALE = rgb("14286B"), rgb("E4002B"), rgb("F3F5FA")
s = kit.new_slide()
kit.shape(s, MSO_SHAPE.RECTANGLE, 0, 0, W, H, PALE)
freeform(s, [(0, 0), (1.05, 0), (0.609, 6.62), (0, 6.62)], fill=NAVY)              # slanted brand band, stops at the footer
txt(s, 0.12, 0.42, 0.75, 0.3, "KESTREL", 11, WHITE, bold=True, smart=False)
kit.shape(s, MSO_SHAPE.RECTANGLE, 1.35, 0.5, 0.08, 0.95, RED)
txt(s, 1.6, 0.45, 11.1, 1.05,
    "KESTREL SHOULD BUILD ITS OWN MEDIA STUDIO, TURNING THE EVENTS IT ALREADY RUNS INTO A NEW REVENUE STREAM",
    24, NAVY, bold=True, anchor=MID)
cols = [("trophy", "Can afford to…", "Hold its lead", "41%", "market share in the UK¹",
         ["Category growing 7% a year", "Customers pay 20% more", "New rivals keep failing"]),
        ("film", "Is able to…", "Make and sell content", "£380M", "sales uplift after its 2025 stunt",
         ["In-house film crews", "Proven global audience", "Ready-made sponsor network"]),
        ("megaphone", "Would add value by…", "Diversifying revenue", "1,100+", "events filmed every year²",
         ["Licensing to streamers", "Year-round engagement", "Less reliance on can sales"])]
cw = 3.55; x0 = 1.55; gap = (W - 0.6 - x0 - 3 * cw) / 2
for k, (ic, small, title, big, cap, bullets) in enumerate(cols):
    x = x0 + k * (cw + gap); cx = x + cw / 2
    kit.shape(s, MSO_SHAPE.OVAL, cx - 0.55, 1.9, 1.1, 1.1, WHITE, line=NAVY, lw=2)
    kit.icon(s, ic, cx - 0.3, 2.15, 0.6, color=NAVY)
    kit.shape(s, MSO_SHAPE.RECTANGLE, cx - 0.85, 1.9, 0.38, 0.38, NAVY)
    txt(s, cx - 0.85, 1.9, 0.38, 0.38, str(k + 1), 16, WHITE, bold=True, align=C, anchor=MID, smart=False)
    txt(s, x, 3.12, cw, 0.3, small, 14, RED, bold=True, align=C, smart=False)
    txt(s, x, 3.42, cw, 0.4, title, 20, NAVY, bold=True, align=C, smart=False)
    txt(s, x, 3.9, cw, 0.8, big, 54, RED, bold=True, align=C, anchor=MID, smart=False)
    txt(s, x, 4.72, cw, 0.3, cap, 14, NAVY, align=C, smart=False)
    kit.bullet_list(s, x + 0.35, 5.2, cw - 0.5, 1.0, bullets, pt=14, color=kit.T["ink"], bullet_color=RED, gap_pt=3)
freeform(s, [(0.609, 6.62), (W, 6.62), (W, H), (0.55, H)], fill=NAVY)              # footer continues the slant, corner left open
txt(s, 1.2, 6.78, 9.5, 0.55, ["Note: ¹ UK energy-drink sales by value, 2025. ² Kestrel-owned sports and culture events.",
                              "Source: company accounts; market panel data (all fictional)."], 10, rgb("C9D2EA"), anchor=MID)
txt(s, W - 1.2, 6.78, 0.6, 0.55, "1", 10, rgb("C9D2EA"), align=R, anchor=MID, smart=False)

# =============================================================================================== 2. pictogram
# Brand: Loop, an (invented) reusable-cup scheme. Fresh green and mint on charcoal.
GREEN, MINT, CHAR, LIGHT = rgb("2E7D5B"), rgb("6FD3A8"), rgb("3B4146"), rgb("B9C0C6")
s = kit.new_slide()
freeform(s, [(0, 0.38), (1.75, 0.38), (1.55, 0.78), (0, 0.78)], fill=GREEN)
txt(s, 0.6, 0.38, 1.0, 0.4, "PROBLEM", 11, WHITE, bold=True, anchor=MID, smart=False)
txt(s, 1.75, 0.62, W - 3.5, 1.0,
    [[("Every year, ", False, GREEN), ("9 million takeaway cups", True, GREEN), (" are thrown away", False, GREEN)],
     [("in Harbourton, enough to fill 40 bin lorries a week", False, GREEN)]], 26, GREEN, align=C, anchor=MID)
band_y, band_h = 1.95, 3.75
kit.shape(s, MSO_SHAPE.RECTANGLE, 0, band_y, W, band_h, CHAR)
rows, ncol = 5, 10
iw = 0.5; gx = (W - 1.2 - iw) / (ncol - 1); gy = (band_h - 0.5 - iw) / (rows - 1)
for r_ in range(rows):
    for c_ in range(ncol):
        last = r_ == rows - 1 and c_ == ncol - 1
        kit.icon(s, "coffee", 0.6 + c_ * gx, band_y + 0.25 + r_ * gy, iw, color=MINT if last else LIGHT)
hx, hy = 0.6 + (ncol - 1) * gx + iw / 2, band_y + 0.25 + (rows - 1) * gy + iw
txt(s, 0.6, 6.0, 9.8, 0.6, [[("But only ", False, GREEN), ("1 in 50", True, GREEN), (" is ever recycled", False, GREEN)]],
    30, GREEN, align=R, anchor=MID)
curved_arrow(s, 10.55, 6.3, hx, hy + 0.12, MINT, 3)
source(s, "Source: Harbourton waste audit 2026; each cup = 180,000 cups a year (fictional data)", kit.T["grey"])
page_no(s, 2, kit.T["grey"])

# =============================================================================================== 3. scenario chart
# Brand: Brightwell Energy, an (invented) home-battery maker. Forest green with an ember orange.
FOREST, EMBER, SAGE, INK3 = rgb("1B4332"), rgb("E76F51"), rgb("A3B1A8"), rgb("1F2A24")
s = kit.new_slide()
kit.shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, W - 0.6 - 1.55, 0.42, 1.55, 0.32, FOREST).adjustments[0] = 0.5
txt(s, W - 0.6 - 1.55, 0.42, 1.55, 0.32, "Market outlook", 11, WHITE, bold=True, align=C, anchor=MID, smart=False)
txt(s, 0.6, 0.62, 10.3, 1.0, "Falling costs will grow the home-battery market at least one and a half times by 2030, "
    "whichever scenario plays out", 26, FOREST, bold=True, anchor=TOP_)
years = [str(y) for y in range(2016, 2031)]
hist = [20, 28, 40, 55, 75, 100, 135, 175, 225, 290]
base_end = 290
mult = {"Bull": 3.6, "Base": 2.5, "Bear": 1.6}
def path(m):   # smooth path from 2025 to 2030 reaching base_end * m
    return [base_end * (m ** (k / 5)) for k in range(0, 6)]
cd = CategoryChartData(); cd.categories = years
cd.add_series("Installed", hist + [None] * 5)
for nm in ("Bull", "Base", "Bear"):
    cd.add_series(nm, [None] * 9 + [round(v) for v in path(mult[nm])])
cx_, cy_, cw_, ch_ = 0.6, 2.25, 8.1, 4.35
txt(s, cx_ + 0.62, 1.85, 6, 0.35, "Home batteries installed in the UK, thousands", 16, EMBER, bold=True, smart=False)
gf = s.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(cx_), Inches(cy_), Inches(cw_), Inches(ch_), cd)
ch = gf.chart; kit.style_chart(ch, pt=12); kit.transparent_chart(ch)
PX, PY, PW, PH = 0.075, 0.03, 0.9, 1 - 0.03 - 0.45 / ch_
kit.plot_layout(ch, PX, PY, PW, PH)
cols3 = [INK3, FOREST, SAGE, EMBER]
for ser, col, dash in zip(ch.plots[0].series, cols3, [None, MSO_LINE_DASH_STYLE.DASH, MSO_LINE_DASH_STYLE.DASH, MSO_LINE_DASH_STYLE.DASH]):
    ser.format.line.color.rgb = col; ser.format.line.width = Pt(3); ser.smooth = True
    if dash:
        ser.format.line.dash_style = dash
    m_ = ser._element.find("{http://schemas.openxmlformats.org/drawingml/2006/chart}marker")
    if m_ is None:
        m_ = etree.SubElement(ser._element, "{http://schemas.openxmlformats.org/drawingml/2006/chart}marker")
    etree.SubElement(m_, "{http://schemas.openxmlformats.org/drawingml/2006/chart}symbol").set("val", "none")
va = ch.value_axis; va.minimum_scale = 0; va.maximum_scale = 1100; va.major_unit = 250; va.has_major_gridlines = False
kit.axis(va, None, pt=12); kit.axis(ch.category_axis, pt=12)
ch.category_axis.tick_labels.font.size = Pt(12)
from pptx.oxml.ns import qn  # noqa: E402
ca = ch.category_axis._element
skip = ca.find(qn("c:tickLblSkip"))
if skip is None:
    skip = etree.SubElement(ca, qn("c:tickLblSkip"))
skip.set("val", "2")
X = lambda i: cx_ + (PX + PW * (i + 0.5) / 15) * cw_
Y = lambda v: cy_ + (PY + PH) * ch_ - v / 1100 * PH * ch_
kit.line(s, X(9), Y(0), X(9), cy_ + PY * ch_, T["light"], 1, dash=MSO_LINE_DASH_STYLE.DASH)
txt(s, X(9) + 0.1, cy_ + PY * ch_, 1.6, 0.3, "Forecast →", 12, T["grey"], smart=False)
kit.alt(gf, "Line chart of UK home batteries installed, 20 thousand in 2016 rising to 290 thousand in 2025, with three "
            "scenarios to 2030: bull 1,044 thousand (3.6x), base 725 thousand (2.5x), bear 464 thousand (1.6x).")
# bubbles beside the line ends, joined by dotted leaders
bx = 9.1; d = 0.95
labels = {"Bull": ("Bull case", "Prices fall 40%\vGrants continue", FOREST),
          "Base": ("Base case", "Prices fall 25%\vGrants end 2028", rgb("7F9187")),
          "Bear": ("Bear case", "Prices flat\vGrants end 2027", EMBER)}
ys = {"Bull": 2.35, "Base": 3.85, "Bear": 5.35}
for nm in ("Bull", "Base", "Bear"):
    t_, cap, col = labels[nm]; by = ys[nm]
    ey = Y(base_end * mult[nm])
    kit.line(s, X(14) + 0.05, ey, bx - 0.05, by + d / 2, col, 1.5, dash=MSO_LINE_DASH_STYLE.ROUND_DOT)
    kit.shape(s, MSO_SHAPE.OVAL, bx, by, d, d, col)
    txt(s, bx, by, d, d, f"{mult[nm]}x", 22, WHITE, bold=True, align=C, anchor=MID, smart=False)
    txt(s, bx + d + 0.2, by + 0.03, 2.5, 0.4, t_, 18, col, bold=True, smart=False)
    txt(s, bx + d + 0.2, by + 0.45, 2.5, 0.55, cap, 14, T["grey"], smart=False)
source(s, "Note: multiples are 2030 vs 2025. Source: industry installer survey; Brightwell scenarios (fictional data)", T["grey"])
page_no(s, 3, T["grey"])

# =============================================================================================== 4. market sizing
# Brand: Tessera Health, an (invented) digital-therapy start-up. Ocean teals from light to deep.
STAGEB, T1, T2, T3, BAND, INK4 = rgb("1C4E80"), rgb("5DB7B2"), rgb("3A8F8B"), rgb("23595A"), rgb("F1F4F4"), rgb("1D2B33")
s = kit.new_slide()
txt(s, 0.6, 0.5, 12.1, 1.0, "The market grows almost threefold as it matures, and our realistic share rises from "
    "£3–7M today to £24–35M", 26, STAGEB, bold=True)
heads = [("Market stage", STAGEB, 0.6, 4.1), ("Total market", T1, 5.45, 2.1), ("Serviceable", T2, 7.95, 2.1),
         ("Our share", T3, 10.45, 2.1)]
hy = 1.75
for t_, col, x, w in heads:
    freeform(s, [(x, hy), (x + w - 0.15, hy), (x + w, hy + 0.15), (x + w, hy + 0.45), (x, hy + 0.45)], fill=col)
    txt(s, x, hy, w, 0.45, t_, 16, WHITE, bold=True, align=C, anchor=MID, smart=False)
data = [("Early adopters", "Today", 420, 171, "£3–7M", 5), ("Growing", "2028", 758, 308, "£15–23M", 19),
        ("Mainstream", "2031", 1160, 472, "£24–35M", 29.5)]
row_y, row_h, rg = 2.35, 1.38, 0.08
dmax = 1.25
colmax = {0: 1160, 1: 472, 2: 29.5}
def dia(v, col):   # area proportional to value within each column; the smallest circles are enlarged to fit labels
    return max(0.9, dmax * math.sqrt(v / colmax[col]))
for k, (stage, when, tam, sam, som, som_mid) in enumerate(data):
    y = row_y + k * (row_h + rg); cyr = y + row_h / 2
    kit.shape(s, MSO_SHAPE.RECTANGLE, 0.6, y, W - 1.2, row_h, BAND)
    # mini adoption curve: the whole S-curve faint, this stage's stretch bold, with a dot at its end
    gx0, gw, gy0, gh = 0.95, 2.4, y + 0.22, row_h - 0.5
    pts = [(gx0 + gw * t / 40, gy0 + gh - gh / (1 + math.exp(-9 * (t / 40 - 0.62)))) for t in range(41)]
    freeform(s, pts, line=rgb("C3CCD3"), lw=2, closed=False)
    seg = {0: (0, 16), 1: (16, 28), 2: (28, 40)}[k]
    freeform(s, pts[seg[0]:seg[1] + 1], line=STAGEB, lw=3.5, closed=False)
    ex, ey = pts[seg[1]]
    kit.shape(s, MSO_SHAPE.OVAL, ex - 0.08, ey - 0.08, 0.16, 0.16, WHITE, line=STAGEB, lw=2.5)
    for j, xx in enumerate((16, 28)):
        kit.line(s, gx0 + gw * xx / 40, gy0 - 0.05, gx0 + gw * xx / 40, gy0 + gh + 0.05, rgb("C3CCD3"), 0.75, dash=MSO_LINE_DASH_STYLE.DASH)
    txt(s, 3.6, cyr - 0.42, 1.75, 0.4, stage, 18, STAGEB, bold=True, smart=False)
    txt(s, 3.6, cyr + 0.02, 1.75, 0.35, when, 14, T["grey"], smart=False)
    for ci, ((val, lab), (_, col, x, w)) in enumerate(zip([(tam, f"£{tam}M" if tam < 1000 else f"£{tam / 1000:.2f}bn"),
                                                        (sam, f"£{sam}M"), (som_mid, som)], heads[1:])):
        dd = dia(val, ci)
        ccx = x + w / 2
        kit.shape(s, MSO_SHAPE.OVAL, ccx - dd / 2, cyr - dd / 2, dd, dd, col)
        txt(s, ccx - 0.8, cyr - 0.25, 1.6, 0.5, lab, 15 if dd >= 1.05 else 13, WHITE, bold=True, align=C, anchor=MID, smart=False)
source(s, "Note: circle area ∝ value within each column; smallest circles enlarged to fit labels. "
          "Source: payer data; adoption model (fictional)", T["grey"])
page_no(s, 4, T["grey"])

# =============================================================================================== 5. org options
# Brand: Haldane Trust, an (invented) health charity. Midnight blue, sky blue, and a coral highlight.
MID_B, SKY, CORAL, LINEC = rgb("1D3557"), rgb("A8DADC"), rgb("D1495B"), rgb("8899AA")
s = kit.new_slide()
txt(s, 0.6, 0.38, 6, 0.3, "TEAM OPTIONS", 13, T["grey"], bold=True, smart=False)
txt(s, 0.6, 0.7, 12.1, 1.0, "With a team of five, there are three ways to set up the new partnerships unit; "
    "we recommend a ringfenced team", 26, MID_B, bold=True)
opts = [("Ringfenced team", "One team led by a partnerships director, reporting to the chief executive", "ring"),
        ("Distributed team", "Five people with protected time, spread across existing teams", "dist"),
        ("Integrated team", "Partnership work shared among current staff, equal to five new posts", "integ")]
ow = 3.85; ogap = (W - 1.2 - 3 * ow) / 2; oy = 2.1; oh = 4.5
def org(s, kind, cx, top):
    bw, bh = 0.34, 0.22
    def bx(x, y, new=False, half=False):
        kit.shape(s, MSO_SHAPE.RECTANGLE, x, y, bw, bh, SKY if new else WHITE, line=MID_B, lw=1.75)
        if half:
            kit.shape(s, MSO_SHAPE.RECTANGLE, x + bw / 2, y + 0.02, bw / 2 - 0.02, bh - 0.04, SKY)
    kit.line(s, cx, top + bh, cx, top + 0.45, LINEC, 1)
    bx(cx - bw / 2, top)
    txt(s, cx + bw / 2 + 0.08, top - 0.02, 1.4, 0.26, "Chief executive", 11, MID_B, bold=True, smart=False)
    if kind == "ring":
        heads_ = [cx - 0.75, cx, cx + 0.75]; newcol = 2
    else:
        heads_ = [cx - 0.55, cx + 0.55]; newcol = None
    kit.line(s, heads_[0] + bw / 2, top + 0.45, heads_[-1] + bw / 2, top + 0.45, LINEC, 1)
    for i, hx_ in enumerate(heads_):
        hx_ -= bw / 2
        kit.line(s, hx_ + bw / 2, top + 0.45, hx_ + bw / 2, top + 0.55, LINEC, 1)
        bx(hx_, top + 0.55, new=(newcol == i))
        n = (4 if newcol == i else 5) if kind == "ring" else 3
        for j in range(n):
            yy = top + 0.95 + j * 0.33
            if kind == "ring":
                kit.line(s, hx_ + 0.06, top + 0.77, hx_ + 0.06, yy + bh / 2, LINEC, 1)
                kit.line(s, hx_ + 0.06, yy + bh / 2, hx_ + 0.14, yy + bh / 2, LINEC, 1)
                bx(hx_ + 0.14, yy, new=(newcol == i))
            else:
                for side in (-0.21, 0.21):
                    xx = hx_ + side
                    new = kind == "dist" and ((i, j, side) in {(0, 0, 0.21), (0, 2, -0.21), (1, 0, -0.21), (1, 1, 0.21), (1, 2, 0.21)})
                    bx(xx, yy, new=new, half=(kind == "integ"))
for k, (title, desc, kind) in enumerate(opts):
    x = 0.6 + k * (ow + ogap)
    pref = k == 0
    if pref:
        kit.shape(s, MSO_SHAPE.RECTANGLE, x, oy, ow, oh, None, line=CORAL, lw=3)
        kit.shape(s, MSO_SHAPE.RECTANGLE, x + ow / 2 - 1.05, oy - 0.2, 2.1, 0.4, WHITE)
        txt(s, x + ow / 2 - 1.05, oy - 0.2, 2.1, 0.4, "Recommended", 15, CORAL, bold=True, align=C, anchor=MID, smart=False)
    elif k == 1:
        kit.line(s, x - ogap / 2, oy + 0.1, x - ogap / 2, oy + oh - 0.1, T["light"], 1)
    if k == 2:
        kit.line(s, x - ogap / 2, oy + 0.1, x - ogap / 2, oy + oh - 0.1, T["light"], 1)
    txt(s, x + 0.2, oy + 0.3, ow - 0.4, 0.4, title.upper(), 18, MID_B, bold=True, align=C, smart=False)
    txt(s, x + 0.3, oy + 0.75, ow - 0.6, 0.75, desc, 14, T["ink"], align=C)
    org(s, kind, x + ow / 2, oy + 1.65)
# legend
lx = W - 0.6 - 4.1
kit.shape(s, MSO_SHAPE.RECTANGLE, lx, 6.73, 0.3, 0.2, SKY, line=MID_B, lw=1.5)
txt(s, lx + 0.4, 6.7, 1.6, 0.26, "New partnership roles", 11, T["grey"], smart=False)
kit.shape(s, MSO_SHAPE.RECTANGLE, lx + 2.3, 6.73, 0.3, 0.2, WHITE, line=MID_B, lw=1.5)
txt(s, lx + 2.7, 6.7, 1.4, 0.26, "Existing roles", 11, T["grey"], smart=False)
source(s, "Note: illustrative structures (fictional)", T["grey"], y=6.95)
page_no(s, 5, T["grey"])

# =============================================================================================== 6. agenda
# Brand: Corran Foundation, an (invented) grant-maker. A bold crimson field with white callouts.
CRIMSON = rgb("B3132B")
s = kit.new_slide()
kit.shape(s, MSO_SHAPE.RECTANGLE, 0, 0, W, H, CRIMSON)
txt(s, 0.6, 0.55, 6, 0.35, "AGENDA", 16, rgb("F6C9CF"), bold=True, smart=False)
txt(s, 0.6, 0.9, 9, 0.7, "What we'll cover today", 36, WHITE, bold=True, smart=False)
ly = 3.7
kit.line(s, 0, ly, W, ly, WHITE, 2, dash=MSO_LINE_DASH_STYLE.DASH)
stops = [(2.55, "1", ["Where are we now,", "and why does it", "matter?"], False),
         (6.67, "2", None, True),
         (10.78, "3", ["What happens next,", "and who does", "what?"], False)]
for cx, n, lines, cur in stops:
    f = 0.62 if cur else 0.42
    kit.shape(s, MSO_SHAPE.RECTANGLE, cx - f / 2, ly - 0.3 - f - 0.12, f, f, WHITE)
    txt(s, cx - f / 2, ly - 0.3 - f - 0.12, f, f, n, 30 if cur else 20, CRIMSON, bold=True, align=C, anchor=MID, smart=False)
    freeform(s, [(cx - 0.1, ly - 0.42), (cx + 0.1, ly - 0.42), (cx, ly - 0.3)], fill=WHITE)
    kit.shape(s, MSO_SHAPE.OVAL, cx - 0.17, ly - 0.17, 0.34, 0.34, CRIMSON, line=WHITE, lw=3)
    bw, bh = (3.1, 1.75) if cur else (2.3, 1.45)
    by = ly + 0.45
    freeform(s, [(cx - 0.16, by), (cx, by - 0.16), (cx + 0.16, by)], fill=WHITE)
    kit.shape(s, MSO_SHAPE.RECTANGLE, cx - bw / 2, by, bw, bh, WHITE)
    if cur:
        txt(s, cx - bw / 2 + 0.1, by, bw - 0.2, bh, [[("How", True, INK4), (" do we choose\vwhere to grow, and\vwhat will it take?", False, INK4)]],
            20, INK4, align=C, anchor=MID)
    else:
        txt(s, cx - bw / 2 + 0.15, by, bw - 0.3, bh, ["\v".join(lines)], 16, INK4, align=C, anchor=MID, smart=False)
txt(s, W - 1.2, 6.95, 0.6, 0.3, "6", 10, rgb("F6C9CF"), align=R, anchor=BOT, smart=False)

# =============================================================================================== 7. annotated chart
# Brand: The Orpheum, an (invented) independent cinema. Deep slate with a cool teal for the opportunities.
BURG, TEAL7, GREY7, SHADE = rgb("1D3A4A"), rgb("1F8A8A"), rgb("D9D9D9"), rgb("E6F4F3")
s = kit.new_slide()
nav = [("Overview", False), ("Programme", True), ("Audiences", False), ("Value", False)]
nx = 4.3
for k, (t_, on) in enumerate(nav):
    txt(s, nx + k * 1.3, 0.3, 1.25, 0.28, t_, 11, BURG if on else rgb("A3A3A3"), bold=on, align=C, smart=False)
txt(s, 0.6, 0.72, 12.1, 1.0, "Big releases carry us through summer and winter, but spring and autumn need films "
    "we choose ourselves", 26, BURG, bold=True)
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
lo = [9, 7, 6, 5, 5, 7, 9, 9, 4, 5, 7, 10]; hi = [14, 11, 12, 8, 8.5, 12, 15, 14, 8, 9, 12, 16]
med = [12, 8.5, 9, 6.5, 6.8, 9.5, 12.5, 12, 5.5, 6.5, 9, 13.5]
avg = sum(med) / 12
cx7, cy7, cw7, ch7 = 0.6, 2.3, 9.0, 4.3
PX7, PY7, PW7, PH7 = 0.07, 0.03, 0.91, 1 - 0.03 - 0.4 / ch7
X7 = lambda i: cx7 + (PX7 + PW7 * (i + 0.5) / 12) * cw7
Y7 = lambda v: cy7 + (PY7 + PH7) * ch7 - v / 18 * PH7 * ch7
colw = PW7 * cw7 / 12 * 0.62
for i0, i1, lab, sub in [(3, 4, "Late-spring dip", "Families wait for summer"), (8, 9, "Autumn dip", "Students arrive in town")]:
    x0 = X7(i0) - PW7 * cw7 / 24; x1 = X7(i1) + PW7 * cw7 / 24
    kit.shape(s, MSO_SHAPE.RECTANGLE, x0, Y7(18), x1 - x0, Y7(0) - Y7(18), SHADE, line=TEAL7, lw=1.5)
    txt(s, x0 - 0.4, 1.78, x1 - x0 + 0.8, 0.3, lab, 14, TEAL7, bold=True, align=C, smart=False)
    txt(s, x0 - 0.5, 2.03, x1 - x0 + 1.0, 0.25, sub, 12, TEAL7, align=C, smart=False)
for i in range(12):
    kit.shape(s, MSO_SHAPE.RECTANGLE, X7(i) - colw / 2, Y7(hi[i]), colw, Y7(lo[i]) - Y7(hi[i]), GREY7)
cd = CategoryChartData(); cd.categories = months; cd.add_series("Median", med)
gf = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(cx7), Inches(cy7), Inches(cw7), Inches(ch7), cd)
ch = gf.chart; kit.style_chart(ch, pt=12); kit.transparent_chart(ch); kit.plot_layout(ch, PX7, PY7, PW7, PH7)
ser = ch.plots[0].series[0]; ser.format.line.color.rgb = BURG; ser.format.line.width = Pt(3); ser.smooth = False
from pptx.enum.chart import XL_MARKER_STYLE  # noqa: E402
ser.marker.style = XL_MARKER_STYLE.CIRCLE; ser.marker.size = 7
ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = BURG; ser.marker.format.line.color.rgb = WHITE
va = ch.value_axis; va.minimum_scale = 0; va.maximum_scale = 18; va.major_unit = 3; va.has_major_gridlines = False
kit.axis(va, "Admissions (thousands)", pt=12); kit.axis(ch.category_axis, pt=12)
kit.alt(gf, "Monthly cinema admissions: median line with the range of each month behind it. Admissions dip in April-May "
            "(about 6.5 thousand) and September-October (about 6 thousand), against an average of about 9.3 thousand.")
kit.line(s, X7(0) - 0.3, Y7(avg), X7(11) + 0.3, Y7(avg), rgb("8C8C8C"), 1, dash=MSO_LINE_DASH_STYLE.ROUND_DOT)
txt(s, X7(11) + 0.4, Y7(avg) - 0.13, 0.9, 0.25, "Average", 11, rgb("707070"), smart=False)
lx7 = 9.95
kit.shape(s, MSO_SHAPE.RECTANGLE, lx7, 2.55, 0.32, 0.5, GREY7)
txt(s, lx7 + 0.45, 2.52, 2.4, 0.6, "Lowest to highest\vmonth, 2016–2025", 14, rgb("595959"), smart=False)
kit.line(s, lx7, 3.55, lx7 + 0.32, 3.55, BURG, 3)
txt(s, lx7 + 0.45, 3.4, 2.4, 0.6, "Median month", 14, BURG, bold=True, smart=False)
kit.shape(s, MSO_SHAPE.RECTANGLE, lx7, 4.65, 2.7, 1.1, WHITE, line=rgb("8C8C8C"), lw=1)
txt(s, lx7 + 0.15, 4.65, 2.4, 1.1, "Some dip is unavoidable: the aim is to lift the floor, not flatten the year", 14,
    rgb("404040"), anchor=MID)
curved_arrow(s, lx7, 5.2, X7(9) + 0.35, Y7(7), rgb("8C8C8C"), 1.5)
source(s, "Source: box-office records, 2016–2025 (fictional data)", T["grey"])
page_no(s, 7, T["grey"])

# =============================================================================================== 8. heatmap
# Brand: Arden Group, an (invented) professional-services firm. Blue intensity scale with orange for problems.
BLUE8, ORANGE8 = rgb("1F5FA8"), rgb("E07A2E")
s = kit.new_slide()
freeform(s, [(W - 2.1, 0), (W, 0), (W, 2.1)], fill=rgb("8E9BD8"))
txt(s, W - 0.72 - 0.72, 0.72 - 0.16, 1.44, 0.32, "Illustrative", 12, WHITE, bold=True, align=C, anchor=MID, smart=False).rotation = 45
grid = [[1, 0, 0, 0, 0, 0, 0, 0], [2, 5, 4, 0, 0, 0, 3, 0], [0, 1, 6, 28, 9, 0, 5, 2],
        [0, 0, 2, 14, 61, 38, 4, 0], [0, 0, 0, 3, 22, 95, 57, 9], [0, 0, 0, 0, 1, 18, 74, 41]]
grades = ["G8", "G7", "G6", "G5", "G4", "G3", "G2", "G1"]
total8 = sum(map(sum, grid))
wrong8 = grid[1][6] + grid[1][7] + grid[2][6] + grid[2][7] + grid[3][2] + grid[3][3] + grid[4][2] + grid[4][3]
assert (total8, wrong8) == (505, 29)
txt(s, 0.6, 0.55, 10.8, 1.0, f"Nearly {wrong8 // 10 * 10} people sit at the wrong grade for their layer, adding cost and "
    "blurring who is accountable", 26, BLUE8, bold=True)
gx0, gy0, cs = 4.7, 2.3, 0.54
txt(s, gx0, gy0 - 0.62, 8 * cs, 0.3, "Grade (most senior on the left)", 12, BLUE8, bold=True, align=C, smart=False)
for j, gname in enumerate(grades):
    txt(s, gx0 + j * cs, gy0 - 0.32, cs, 0.28, gname, 12, rgb("5A6B80"), align=C, smart=False)
kit.line(s, gx0 + 0.1, gy0 + 0.1, gx0 + 8 * cs - 0.1, gy0 + 6 * cs - 0.1, rgb("9AA5B1"), 1.25, dash=MSO_LINE_DASH_STYLE.DASH)
for i, row in enumerate(grid):
    txt(s, gx0 - 0.85, gy0 + i * cs, 0.75, cs, f"Layer {i + 1}", 12, rgb("5A6B80"), align=R, anchor=MID, smart=False)
    for j, v in enumerate(row):
        x, y = gx0 + j * cs, gy0 + i * cs
        if v:
            a = min(1, 0.12 + v / 100)
            fill = kit._tint(BLUE8, 1 - a)
            kit.shape(s, MSO_SHAPE.RECTANGLE, x + 0.02, y + 0.02, cs - 0.04, cs - 0.04, fill)
            txt(s, x, y, cs, cs, str(v), 12, WHITE if a > 0.5 else BLUE8, bold=a > 0.5, align=C, anchor=MID, smart=False)
        else:
            kit.shape(s, MSO_SHAPE.RECTANGLE, x + 0.02, y + 0.02, cs - 0.04, cs - 0.04, rgb("F3F5F8"))
# clusters: junior grades in senior layers (top right) and senior grades deep in the organisation (bottom left)
kit.shape(s, MSO_SHAPE.RECTANGLE, gx0 + 6 * cs, gy0 + 1 * cs, 2 * cs, 2 * cs, None, line=ORANGE8, lw=2.5)
kit.shape(s, MSO_SHAPE.RECTANGLE, gx0 + 2 * cs, gy0 + 3 * cs, 2 * cs, 2 * cs, None, line=ORANGE8, lw=2.5)
txt(s, gx0, gy0 + 6 * cs + 0.12, 8 * cs, 0.3, f"Total: {total8} people", 12, rgb("5A6B80"), align=C, smart=False)
lx8 = 0.6
txt(s, lx8, 3.55, 2.75, 0.3, "Below the line", 14, ORANGE8, bold=True, smart=False)
txt(s, lx8, 3.85, 2.75, 0.6, "Senior grades deep in the organisation", 16, BLUE8, bold=True)
txt(s, lx8, 4.55, 2.75, 1.2, [[("Why: ", True, rgb("404040")), ("long tenure, specialist roles", False, rgb("404040"))],
    [("Impact: ", True, rgb("404040")), ("higher cost, slower decisions", False, rgb("404040"))]], 14, rgb("404040"))
kit.arrow(s, 3.35, gy0 + 4 * cs, gx0 + 2 * cs - 0.05, gy0 + 4 * cs, ORANGE8, 1.75)
rx8 = gx0 + 8 * cs + 0.6
txt(s, rx8, 2.35, 3.2, 0.3, "Above the line", 14, ORANGE8, bold=True, smart=False)
txt(s, rx8, 2.65, 3.2, 0.6, "Junior grades in senior layers", 16, BLUE8, bold=True)
txt(s, rx8, 3.3, 3.2, 1.2, [[("Why: ", True, rgb("404040")), ("assistants and small teams", False, rgb("404040"))],
    [("Impact: ", True, rgb("404040")), ("senior work falls to junior staff", False, rgb("404040"))]], 14, rgb("404040"))
curved_arrow(s, rx8 - 0.05, 2.6, gx0 + 8 * cs + 0.05, gy0 + 1.4 * cs, ORANGE8, 1.75)
txt(s, rx8, 5.0, 3.2, 0.9, "Should everyone at G2 sit in the same layers across the business?", 16, BLUE8, bold=True)
source(s, "Note: counts of people by organisational layer and grade (fictional data)", T["grey"])
page_no(s, 8, T["grey"])

# =============================================================================================== 9. range of estimates
# Brand: neutral research house style. A single strong blue, one annotation colour.
ROYAL = rgb("2459B8")
s = kit.new_slide()
txt(s, 0.6, 0.55, 12.1, 1.0, "Estimates of the home-care technology market vary more than tenfold, from £3.1bn "
    "to £38bn", 26, rgb("16213E"), bold=True)
txt(s, 8.8, 1.75, 3.9, 0.35, "Non-exhaustive", 16, rgb("6B7280"), align=R, smart=False).text_frame.paragraphs[0].runs[0].font.italic = True
src9 = [("Investment bank", "All care spending¹"), ("Consultancy", "Care plus prevention"), ("Industry body", "Paid care services"),
        ("Policy review", "Funded care only"), ("Venture investor", "Software and devices")]
vals9 = [38, 21.5, 6, 4.2, 3.1]
cd = CategoryChartData(); cd.categories = [a for a, _ in src9]; cd.add_series("Estimate", vals9)
cx9, cy9, cw9, ch9 = 0.6, 2.2, 8.4, 4.25
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(cx9), Inches(cy9), Inches(cw9), Inches(ch9), cd)
ch = gf.chart; kit.style_chart(ch, pt=12); kit.transparent_chart(ch)
PX9, PY9, PW9, PH9 = 0.08, 0.1, 0.92, 1 - 0.1 - 0.75 / ch9
kit.plot_layout(ch, PX9, PY9, PW9, PH9)
pl = ch.plots[0]; pl.gap_width = 55
ser = pl.series[0]; ser.format.fill.solid(); ser.format.fill.fore_color.rgb = ROYAL
dl = ser.data_labels; dl.show_value = True; dl.number_format = '"£"General"bn"'; dl.number_format_is_linked = False
from pptx.enum.chart import XL_LABEL_POSITION  # noqa: E402
dl.position = XL_LABEL_POSITION.OUTSIDE_END; dl.font.size = Pt(16); dl.font.bold = True; dl.font.color.rgb = rgb("16213E")
va = ch.value_axis; va.minimum_scale = 0; va.maximum_scale = 40; va.major_unit = 10; va.has_major_gridlines = False
kit.axis(va, "£bn", pt=12)
ca9 = ch.category_axis; ca9.tick_labels.font.size = Pt(13); ca9.tick_labels.font.bold = True
ca9.format.line.color.rgb = rgb("9CA3AF")
kit.alt(gf, "Column chart of five estimates of the home-care technology market: £38bn (investment bank), £21.5bn "
            "(consultancy), £6bn (industry body), £4.2bn (policy review) and £3.1bn (venture investor), a more than tenfold range.")
X9 = lambda i: cx9 + (PX9 + PW9 * (i + 0.5) / 5) * cw9
Y9 = lambda v: cy9 + (PY9 + PH9) * ch9 - v / 40 * PH9 * ch9
for i, (_, scope) in enumerate(src9):
    txt(s, X9(i) - 0.85, cy9 + ch9 - 0.3, 1.7, 0.3, scope, 12, rgb("4B5563"), align=C, smart=False)
bxx = 9.25
kit.line(s, X9(0) + 0.6, Y9(38), bxx, Y9(38), ROYAL, 1, dash=MSO_LINE_DASH_STYLE.DASH)
kit.line(s, X9(4) + 0.6, Y9(3.1), bxx, Y9(3.1), ROYAL, 1, dash=MSO_LINE_DASH_STYLE.DASH)
kit.line(s, bxx, Y9(38), bxx, Y9(3.1), ROYAL, 1.75)
for yy in (Y9(38), Y9(3.1)):
    kit.line(s, bxx - 0.08, yy, bxx + 0.08, yy, ROYAL, 1.75)
kit.shape(s, MSO_SHAPE.RECTANGLE, 9.6, 3.05, 3.1, 1.9, rgb("EEF3FC"))
txt(s, 9.85, 3.15, 2.7, 0.8, "~12x", 48, ROYAL, bold=True, smart=False)
txt(s, 9.85, 3.95, 2.7, 0.9, "gap between the highest and lowest estimates", 15, rgb("374151"))
source(s, "Note: ¹ includes informal care, so wider than the technology market. Sources: published market reports "
          "(all fictional)", T["grey"])
page_no(s, 9, T["grey"])

# =============================================================================================== 10. barriers and solutions
# Brand: Warmly, an (invented) home-heating programme. Violet and lavender on a dark split.
VIOLET, LAV, DARK10 = rgb("5B3FA0"), rgb("B9A5F0"), rgb("2B2733")
s = kit.new_slide()
split = 6.2
freeform(s, [(split, 1.95), (W, 1.95), (W, H), (split, H)], fill=DARK10)
freeform(s, [(0, 0.38), (1.85, 0.38), (1.65, 0.78), (0, 0.78)], fill=VIOLET)
txt(s, 0.6, 0.38, 1.2, 0.4, "APPROACH", 11, WHITE, bold=True, anchor=MID, smart=False)
txt(s, 2.0, 0.55, 10.7, 1.1, [[("Three ", False, VIOLET), ("barriers", True, VIOLET), (" stop people switching to heat pumps; ", False, VIOLET)],
    [("our programme ", False, VIOLET), ("removes each one", True, VIOLET)]], 26, VIOLET, align=C, anchor=MID)
txt(s, 0.6, 2.05, 2.5, 0.35, "BARRIER", 14, rgb("6B6B6B"), bold=True, smart=False)
txt(s, split + 0.75, 2.05, 3.5, 0.35, "WHAT WE DO", 14, LAV, bold=True, smart=False)
rows10 = [("wallet", [("High ", False), ("upfront cost", True), (" of", False)], "a new system", [("£2,500 grants", True), (" paid at", False)], "the point of sale"),
          ("wrench", [("Too few ", False), ("trained installers", True)], "in the area", [("120 installers", True), (" trained", False)], "in 18 months"),
          ("lightbulb", [("Low ", False), ("awareness", True), (" of grants", False)], "and savings", [("Open-home events", True), (" on", False)], "every street")]
for k, (ic, l1, l2, r1, r2) in enumerate(rows10):
    cyk = 2.95 + k * 1.3
    txt(s, 0.6, cyk - 0.42, split - 1.35, 0.85, [[(t_, b, VIOLET if b else rgb("2B2B2B")) for t_, b in l1],
        [(l2, False, rgb("2B2B2B"))]], 20, rgb("2B2B2B"), align=R, anchor=MID)
    kit.shape(s, MSO_SHAPE.OVAL, split - 0.48, cyk - 0.48, 0.96, 0.96, WHITE, line=LAV, lw=2)
    kit.icon(s, ic, split - 0.27, cyk - 0.27, 0.54, color=VIOLET)
    txt(s, split + 0.75, cyk - 0.42, W - split - 1.35, 0.85, [[(t_, b, LAV if b else WHITE) for t_, b in r1],
        [(r2, False, WHITE)]], 20, WHITE, anchor=MID)
txt(s, split + 0.75, 6.95, 5, 0.3, "Source: programme plan, 2027 (fictional)", 10, rgb("A39FB0"), anchor=BOT)
txt(s, W - 1.2, 6.95, 0.6, 0.3, "10", 10, rgb("A39FB0"), align=R, anchor=BOT, smart=False)


# ---------------------------------------------------------------------------------------------- shared helpers
from pptx.enum.chart import XL_LABEL_POSITION, XL_LEGEND_POSITION  # noqa: E402


def no_markers(ser):
    m_ = ser._element.find("{http://schemas.openxmlformats.org/drawingml/2006/chart}marker")
    if m_ is None:
        m_ = etree.SubElement(ser._element, "{http://schemas.openxmlformats.org/drawingml/2006/chart}marker")
    etree.SubElement(m_, "{http://schemas.openxmlformats.org/drawingml/2006/chart}symbol").set("val", "none")


def pill(s, x, y, w, h, fill, t, pt, color, bold=True):
    b = kit.shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill); b.adjustments[0] = 0.5
    txt(s, x, y, w, h, t, pt, color, bold=bold, align=C, anchor=MID, smart=False)
    return b


# =============================================================================================== 11. PESTLE
# Brand: Calder Homes, an (invented) housebuilder. Deep emerald field, white tiles, amber for the top risks.
EMER, MINT11, AMBER = rgb("0E6B4F"), rgb("A7E3CC"), rgb("F2B33D")
s = kit.new_slide()
kit.shape(s, MSO_SHAPE.RECTANGLE, 0, 0, W, H, EMER)
txt(s, 0.6, 0.45, 9, 0.3, "QUESTION 3: WHAT COULD SLOW US DOWN?", 14, MINT11, bold=True, smart=False)
txt(s, 0.6, 0.8, 11.5, 1.0, "A PESTLE scan shows that the biggest risks to expansion are economic and legal, "
    "not demand", 28, WHITE, bold=True)
pest = [("P", "Political", ["Planning reform delays", "Local election cycle"], False),
        ("E", "Economic", ["Interest rates stay high", "Build costs up 9%"], True),
        ("S", "Social", ["Ageing population", "Demand for flexible living"], False),
        ("T", "Technological", ["Modular build matures", "Smart-home expectations"], False),
        ("L", "Legal", ["New building-safety rules", "Leasehold reform"], True),
        ("E", "Environmental", ["Net-zero standards from 2028", "Flood-risk remapping"], False)]
pg = 0.2; pw = (W - 1.2 - 5 * pg) / 6
for k, (letter, name, items, hot) in enumerate(pest):
    x = 0.6 + k * (pw + pg); cx = x + pw / 2
    kit.shape(s, MSO_SHAPE.RECTANGLE, cx - 0.45, 2.55, 0.9, 0.9, WHITE)
    txt(s, cx - 0.45, 2.55, 0.9, 0.9, letter, 40, EMER, bold=True, align=C, anchor=MID, smart=False)
    txt(s, x - 0.1, 3.58, pw + 0.2, 0.4, name, 18, WHITE, bold=True, align=C, smart=False)
    kit.line(s, cx, 4.05, cx, 4.32, WHITE, 1.5)
    kit.shape(s, MSO_SHAPE.OVAL, cx - 0.06, 4.26, 0.12, 0.12, WHITE)
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, 4.38, pw, 1.75, WHITE)
    if hot:
        kit.shape(s, MSO_SHAPE.RECTANGLE, x, 4.38, pw, 0.09, AMBER)
        pill(s, cx - 0.72, 6.28, 1.44, 0.34, AMBER, "Highest risk", 12, rgb("3A2A00"))
    kit.bullet_list(s, x + 0.14, 4.6, pw - 0.24, 1.4, items, pt=14, color=rgb("22262B"), bullet_color=EMER, gap_pt=6)
txt(s, W - 1.2, 6.95, 0.6, 0.3, "11", 10, MINT11, align=R, anchor=BOT, smart=False)

# =============================================================================================== 12. revenue pools
# Brand: neutral strategy house style. Indigo for digital, charcoal for physical; sizes added to each pool.
IND, PALE12, GREY12, DARK12 = rgb("3A3FB3"), rgb("E9EAF8"), rgb("DCDDE3"), rgb("2A2A31")
s = kit.new_slide()
txt(s, 0.6, 0.5, 12.1, 1.0, "The digital-care market is settling on one definition, with three revenue pools worth "
    "£9bn together", 26, rgb("1E1F3B"), bold=True)
t12 = txt(s, 7.7, 1.68, 5.0, 0.3, "Excludes hardware and infrastructure", 15, rgb("6B6B7B"), align=R, smart=False)
t12.text_frame.paragraphs[0].runs[0].font.italic = True
r1y, r1h, r2y, r2h = 2.15, 2.95, 5.27, 1.25
lw12 = 2.1
kit.shape(s, MSO_SHAPE.RECTANGLE, 0.6, r1y, lw12, r1h, IND)
txt(s, 0.8, r1y, lw12 - 0.3, r1h, "Digital care", 22, WHITE, bold=True, anchor=MID, smart=False)
kit.shape(s, MSO_SHAPE.RECTANGLE, 0.6, r2y, lw12, r2h, DARK12)
txt(s, 0.8, r2y, lw12 - 0.3, r2h, "Physical care tech", 20, WHITE, bold=True, anchor=MID)
px0 = 0.6 + lw12 + 0.2
pools = [("Consumer apps", "Paid subscriptions to wellbeing and self-care apps", "£1.4bn", 2.55),
         ("Provider tools", "Scheduling, records and care planning for providers", "£2.9bn", 2.55)]
for k, (t_, d_, v_, w_) in enumerate(pools):
    x = px0 + k * (w_ + 0.18)
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, r1y, w_, r1h, PALE12)
    txt(s, x + 0.22, r1y + 0.22, w_ - 0.4, 0.4, t_, 20, rgb("1E1F3B"), bold=True, smart=False)
    txt(s, x + 0.22, r1y + 0.7, w_ - 0.4, 1.2, d_, 15, rgb("3A3A48"))
    txt(s, x + 0.22, r1y + r1h - 0.75, w_ - 0.4, 0.55, v_, 28, IND, bold=True, smart=False)
x3 = px0 + 2 * (2.55 + 0.18); w3 = W - 0.6 - x3
kit.shape(s, MSO_SHAPE.RECTANGLE, x3, r1y, w3, r1h, PALE12)
txt(s, x3 + 0.22, r1y + 0.22, w3 - 0.4, 0.4, "Payer and employer services", 20, rgb("1E1F3B"), bold=True, smart=False)
txt(s, x3 + 0.22, r1y + 0.7, w3 - 0.4, 0.6, "Services that councils, insurers and employers buy for their people", 15, rgb("3A3A48"))
cw12 = (w3 - 0.44 - 0.18) / 2
for j, (a_, b_, v_) in enumerate([("Remote monitoring", " (first to scale)", "£3.1bn"), ("Other services", "", "£1.6bn")]):
    xx = x3 + 0.22 + j * (cw12 + 0.18)
    kit.shape(s, MSO_SHAPE.RECTANGLE, xx, r1y + 1.45, cw12, 1.3, WHITE)
    txt(s, xx + 0.18, r1y + 1.55, cw12 - 0.3, 0.6, [[(a_, True, IND)], [(b_.strip(), False, IND)]], 15, IND)
    txt(s, xx + 0.18, r1y + 2.25, cw12 - 0.3, 0.42, v_, 22, IND, bold=True, smart=False)
kit.shape(s, MSO_SHAPE.RECTANGLE, px0, r2y, W - 0.6 - px0, r2h, GREY12)
txt(s, px0 + 0.22, r2y + 0.2, 6.5, 0.4, "Devices and robotics", 20, rgb("1E1F3B"), bold=True, smart=False)
txt(s, px0 + 0.22, r2y + 0.65, 6.8, 0.4, "Sensors, lifting aids and companion robots, outside this review", 15, rgb("3A3A48"))
txt(s, W - 0.6 - 2.3, r2y, 2.1, r2h, "Not sized", 18, rgb("6B6B7B"), bold=True, align=R, anchor=MID, smart=False)
assert round(1.4 + 2.9 + 3.1 + 1.6, 1) == 9.0
source(s, "Note: 2026 UK market sizes. Source: market reports; team analysis (fictional)", T["grey"])
page_no(s, 12, T["grey"])

# =============================================================================================== 13. benchmark panels
# Brand: neutral research style in two blues: estimates dark, benchmarks light.
B13, BL13, PANEL13 = rgb("2F5BD3"), rgb("A9C4F5"), rgb("EEF3FD")
s = kit.new_slide()
txt(s, 0.6, 0.5, 12.1, 1.0, "Both the subscription and advertising estimates fall within plausible ranges when "
    "tested against benchmarks", 26, rgb("16213E"), bold=True)
for k, (col, lab) in enumerate([(B13, "Market estimate"), (BL13, "Benchmark")]):
    lx = 8.9 + k * 2.1
    kit.shape(s, MSO_SHAPE.RECTANGLE, lx, 1.78, 0.24, 0.24, col)
    txt(s, lx + 0.34, 1.72, 1.7, 0.35, lab, 14, rgb("374151"), smart=False)
panels = [("Pet-care subscriptions", ["Analyst A", "Trade body", "Upper bound¹"], [0.76, 0.30, 0.90], [0, 0, 1],
           [("Analyst A's £0.76bn", True), (" implies 3.2M owners paying £20 a month, close to the upper bound", False)]),
          ("Pet-care advertising", ["Analyst A", "All pet advertising"], [0.60, 1.06], [0, 1],
           [("Analyst A's £0.6bn", True), (" is 57% of all pet advertising, most of it already digital", False)])]
pw13 = (W - 1.2 - 0.3) / 2
for k, (title, cats, vals, bench, take) in enumerate(panels):
    x = 0.6 + k * (pw13 + 0.3); y = 2.2; ph = 4.45
    b = kit.shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, pw13, ph, PANEL13); b.adjustments[0] = 0.04
    txt(s, x + 0.3, y + 0.25, pw13 - 0.6, 0.4, title, 20, B13, bold=True, smart=False)
    cd = CategoryChartData(); cd.categories = cats; cd.add_series("£bn", vals)
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x + 0.3), Inches(y + 0.75), Inches(pw13 - 0.6), Inches(2.3), cd)
    ch = gf.chart; kit.style_chart(ch, pt=12); kit.transparent_chart(ch)
    kit.plot_layout(ch, 0.1, 0.14, 0.9, 0.72)
    pl = ch.plots[0]; pl.gap_width = 70; ser = pl.series[0]
    for i, bm in enumerate(bench):
        pt_ = ser.points[i]; pt_.format.fill.solid(); pt_.format.fill.fore_color.rgb = BL13 if bm else B13
    dl = ser.data_labels; dl.show_value = True; dl.number_format = '"£"0.00"bn"'; dl.number_format_is_linked = False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END; dl.font.size = Pt(15); dl.font.bold = True; dl.font.color.rgb = rgb("16213E")
    for bp in dl._element.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr"):
        bp.set("wrap", "none")
    va = ch.value_axis; va.minimum_scale = 0; va.maximum_scale = 1.2; va.major_unit = 0.4; va.has_major_gridlines = False
    kit.axis(va, "£bn", pt=12); va.tick_labels.number_format = '0.0'; va.tick_labels.number_format_is_linked = False
    ch.category_axis.tick_labels.font.size = Pt(13); ch.category_axis.format.line.color.rgb = rgb("B8C2D6")
    kit.alt(gf, f"{title}: " + ", ".join(f"{c} £{v:.2f}bn" for c, v in zip(cats, vals)) + ".")
    kit.shape(s, MSO_SHAPE.RECTANGLE, x + 0.3, y + 3.25, pw13 - 0.6, 0.95, WHITE)
    txt(s, x + 0.5, y + 3.25, pw13 - 1.0, 0.95, [[(t_, b_, B13 if b_ else rgb("1F2937")) for t_, b_ in take]], 15, rgb("1F2937"), anchor=MID)
assert round(3.75e6 * 20 * 12 / 1e9, 2) == 0.90 and round(0.76e9 / 240 / 1e6, 1) == 3.2 and round(0.60 / 1.06 * 100) == 57
source(s, "Note: ¹ 3.75M owners at £20 a month. Sources: analyst and trade reports (all fictional)", T["grey"])
page_no(s, 13, T["grey"])

# =============================================================================================== 14. value bridge
# Brand: Fenwick Foods, an (invented) food manufacturer. Deep green, fresh green gains, coral for the cost.
DG14, G14, CORAL14 = rgb("0B3D2E"), rgb("3BA55D"), rgb("E4572E")
s = kit.new_slide()
txt(s, 0.6, 0.5, 12.1, 1.0, "We can grow profit from £18M to £31M by 2028, with pricing alone worth £5.5M", 26, DG14, bold=True)
steps = [("Today", 18, "total"), ("Pricing", 5.5, "up"), ("Procurement", 3.5, "up"), ("Product mix", 2.5, "up"),
         ("Automation", 2.5, "up"), ("Investment", -1.0, "down"), ("2028", 31, "total")]
run_ = 0; base, val = [], []
for nm, v, kind in steps:
    if kind == "total":
        base.append(0); val.append(v); run_ = v
    elif v >= 0:
        base.append(run_); val.append(v); run_ += v
    else:
        run_ += v; base.append(run_); val.append(-v)
assert run_ == 31 or steps[-1][1] == 31
assert 18 + 5.5 + 3.5 + 2.5 + 2.5 - 1.0 == 31
cd = CategoryChartData(); cd.categories = [n for n, _, _ in steps]
cd.add_series("Base", base); cd.add_series("Value", val)
cx14, cy14, cw14, ch14 = 0.6, 2.35, 9.2, 4.1
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(cx14), Inches(cy14), Inches(cw14), Inches(ch14), cd)
ch = gf.chart; kit.style_chart(ch, pt=12); kit.transparent_chart(ch)
PX14, PY14, PW14, PH14 = 0.09, 0.08, 0.91, 1 - 0.08 - 0.45 / ch14
kit.plot_layout(ch, PX14, PY14, PW14, PH14)
pl = ch.plots[0]; pl.gap_width = 45; pl.overlap = 100
pl.series[0].format.fill.background(); pl.series[0].format.line.fill.background()
for i, (_, v, kind) in enumerate(steps):
    pt_ = pl.series[1].points[i]; pt_.format.fill.solid()
    pt_.format.fill.fore_color.rgb = DG14 if kind == "total" else (G14 if kind == "up" else CORAL14)
va = ch.value_axis; va.minimum_scale = 0; va.maximum_scale = 36; va.major_unit = 6; va.has_major_gridlines = False
kit.axis(va, "Operating profit (£M)", pt=12)
ch.category_axis.tick_labels.font.size = Pt(14); ch.category_axis.format.line.color.rgb = rgb("B5BDB8")
kit.alt(gf, "Profit bridge: £18M today, plus pricing £5.5M, procurement £3.5M, product mix £2.5M and automation £2.5M, "
            "less £1M of investment, gives £31M in 2028.")
X14 = lambda i: cx14 + (PX14 + PW14 * (i + 0.5) / 7) * cw14
Y14 = lambda v: cy14 + (PY14 + PH14) * ch14 - v / 36 * PH14 * ch14
for i, (nm, v, kind) in enumerate(steps):
    top = base[i] + val[i]
    lab = f"£{v:g}M" if kind == "total" else (f"+{v:g}" if v > 0 else f"−{-v:g}")
    txt(s, X14(i) - 0.7, Y14(top) - 0.42, 1.4, 0.36, lab, 18, DG14 if kind == "total" else (G14 if kind == "up" else CORAL14),
        bold=True, align=C, anchor=BOT, smart=False)
for i in range(6):
    lvl = (base[i] + val[i]) if steps[i][2] != "down" else base[i]
    kit.line(s, X14(i) + 0.42, Y14(lvl), X14(i + 1) - 0.42, Y14(lvl), rgb("9AA59F"), 1, dash=MSO_LINE_DASH_STYLE.DASH)
# bracket over the levers
bx0, bx1, by = X14(1) - 0.45, X14(4) + 0.45, Y14(31) - 0.55
kit.line(s, bx0, by, bx1, by, G14, 1.5); kit.line(s, bx0, by, bx0, by + 0.15, G14, 1.5); kit.line(s, bx1, by, bx1, by + 0.15, G14, 1.5)
txt(s, bx0, by - 0.42, bx1 - bx0, 0.36, "+£14M from four levers", 16, G14, bold=True, align=C, smart=False)
rx = 10.25
txt(s, rx, 2.4, 2.5, 0.35, "Confidence", 16, DG14, bold=True, smart=False)
conf = [("Pricing", 3), ("Procurement", 3), ("Product mix", 2), ("Automation", 1)]
for k, (nm, n) in enumerate(conf):
    yy = 2.95 + k * 0.62
    txt(s, rx, yy, 1.6, 0.35, nm, 15, rgb("2F3A35"), smart=False)
    for d_ in range(3):
        kit.shape(s, MSO_SHAPE.OVAL, rx + 1.75 + d_ * 0.27, yy + 0.08, 0.18, 0.18, G14 if d_ < n else rgb("D5DDD8"))
kit.shape(s, MSO_SHAPE.RECTANGLE, rx, 5.55, 2.5, 0.9, rgb("EAF5EE"))
txt(s, rx + 0.15, 5.55, 2.25, 0.9, "Margin rises from 9% to 13%", 15, DG14, bold=True, anchor=MID)
source(s, "Note: £M, operating profit before tax. Source: management accounts; team estimates (fictional)", T["grey"])
page_no(s, 14, T["grey"])

# =============================================================================================== 15. 2x2 prioritisation
# Brand: Westway Housing, an (invented) housing association. Petrol axes, coral quick wins.
PET, COR15, GR15 = rgb("0E4D64"), rgb("E4572E"), rgb("B8C4CA")
s = kit.new_slide()
txt(s, 0.6, 0.5, 12.1, 1.0, "Five of the fourteen ideas are quick wins: high value for tenants and easy to "
    "deliver this year", 26, PET, bold=True)
mx, my, mw, mh = 1.25, 1.95, 6.5, 4.55
kit.shape(s, MSO_SHAPE.RECTANGLE, mx + mw / 2, my, mw / 2, mh / 2, rgb("FCEBE5"))
kit.line(s, mx, my + mh, mx + mw, my + mh, PET, 1.5, head=True)
kit.line(s, mx, my + mh, mx, my, PET, 1.5, head=True)
kit.line(s, mx + mw / 2, my, mx + mw / 2, my + mh, rgb("C9D3D8"), 1, dash=MSO_LINE_DASH_STYLE.DASH)
kit.line(s, mx, my + mh / 2, mx + mw, my + mh / 2, rgb("C9D3D8"), 1, dash=MSO_LINE_DASH_STYLE.DASH)
txt(s, mx, my + mh + 0.08, mw, 0.3, "Easier to deliver →", 14, PET, bold=True, align=R, smart=False)
vl = txt(s, mx - 0.3 - 1.3, my + 1.3 - 0.15, 2.6, 0.3, "Higher value →", 14, PET, bold=True, align=C, smart=False); vl.rotation = 270
for (qx, qy, t_, col) in [(mx + mw / 2 + 0.15, my + 0.12, "QUICK WINS", COR15), (mx + 0.15, my + 0.12, "BIG BETS", rgb("6B7C85")),
                          (mx + mw - 1.9, my + mh - 0.42, "FILL-INS", rgb("6B7C85")), (mx + 0.15, my + mh - 0.42, "PARK", rgb("6B7C85"))]:
    txt(s, qx, qy, 1.75, 0.3, t_, 13, col, bold=True, align=R if qx > mx + mw / 2 + 0.5 else L, smart=False)
ideas = [(8.4, 8.2, 3, "Online repairs booking", True), (7.4, 8.1, 2, "Text-message updates", True), (9.0, 6.6, 1, "Damp checks on visits", True),
         (6.6, 7.0, 2, "Tenant advice line", True), (8.0, 5.9, 1, "Key-safe scheme", True),
         (2.2, 9.0, 4, "Retrofit insulation"), (3.4, 7.6, 3, "New repairs contractor"), (1.6, 6.4, 4, "Community hubs"),
         (6.5, 3.8, 1, "Welcome packs"), (8.4, 2.6, 1, "Newsletter redesign"), (2.6, 3.1, 2, "New CRM system"),
         (4.2, 2.0, 2, "Rebrand"), (1.4, 1.6, 3, "Merge two depots"), (5.6, 3.0, 1, "Survey refresh")]
qn = 0
for (ex, vy, cost, name, *qw) in ideas:
    win = bool(qw)
    d = 0.32 + 0.13 * cost
    cx, cy = mx + ex / 10 * mw, my + mh - vy / 10 * mh
    kit.shape(s, MSO_SHAPE.OVAL, cx - d / 2, cy - d / 2, d, d, COR15 if win else GR15, line=WHITE, lw=1.5)
    if win:
        qn += 1
        txt(s, cx - d / 2, cy - d / 2, d, d, str(qn), 14, WHITE, bold=True, align=C, anchor=MID, smart=False)
lx15 = 8.55
txt(s, lx15, 2.0, 4.2, 0.35, "The five quick wins", 18, COR15, bold=True, smart=False)
for k, (_, _, _, name, *_q) in enumerate([i for i in ideas if len(i) == 5]):
    yy = 2.55 + k * 0.6
    kit.shape(s, MSO_SHAPE.OVAL, lx15, yy, 0.38, 0.38, COR15)
    txt(s, lx15, yy, 0.38, 0.38, str(k + 1), 14, WHITE, bold=True, align=C, anchor=MID, smart=False)
    txt(s, lx15 + 0.55, yy - 0.01, 3.6, 0.4, name, 16, rgb("1F2A30"), anchor=MID, smart=False)
txt(s, lx15, 5.55, 3, 0.3, "Bubble size: cost", 12, rgb("5B6970"), bold=True, smart=False)
for k, (d_, lab) in enumerate([(0.45, "Under £50k"), (0.58, "£50–250k"), (0.84, "Over £250k")]):
    ccx = lx15 + 0.42 + k * 1.35
    kit.shape(s, MSO_SHAPE.OVAL, ccx - d_ / 2, 6.3 - d_, d_, d_, GR15)
    txt(s, ccx - 0.65, 6.36, 1.3, 0.28, lab, 11, rgb("5B6970"), align=C, smart=False)
source(s, "Note: scores from the tenant panel and delivery team, 1–10 (fictional)", T["grey"])
page_no(s, 15, T["grey"])

# =============================================================================================== 16. journey funnel
# Brand: Kindred Credit, an (invented) community lender. Plum with a magenta alarm on the leak.
PLUM, MAG, PL16 = rgb("4B1D52"), rgb("D6246E"), rgb("D9C8DE")
s = kit.new_slide()
txt(s, 0.6, 0.5, 12.1, 1.0, "Half of all applicants drop out at the identity check, the biggest leak in the "
    "loan journey", 26, PLUM, bold=True)
stages = [("Visited the site", 100000), ("Started an application", 42000), ("Passed the identity check", 21000),
          ("Approved", 15500), ("Loan paid out", 12400)]
fx0, fw, fy0, fh, fg = 3.6, 6.2, 2.05, 0.72, 0.18
for k, (nm, n) in enumerate(stages):
    w_ = fw * (n / 100000) ** 0.55
    x = fx0 + (fw - w_) / 2; y = fy0 + k * (fh + fg)
    leak = k == 2
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, y, w_, fh, MAG if leak else (PLUM if k in (0, 4) else rgb("7A4B82")))
    txt(s, x, y, w_, fh, f"{n:,}", 20, WHITE, bold=True, align=C, anchor=MID, smart=False)
    txt(s, 0.6, y, fx0 - 0.85, fh, nm, 16, MAG if leak else PLUM, bold=leak, align=R, anchor=MID)
    if k:
        conv = n / stages[k - 1][1]
        xr = fx0 + fw + 0.25
        txt(s, xr, y - fg / 2 - 0.25, 1.2, 0.5, f"{conv:.0%}", 18 if leak else 15, MAG if leak else rgb("7A4B82"),
            bold=True, anchor=MID, smart=False)
assert stages[2][1] / stages[1][1] == 0.5
txt(s, fx0 + fw + 0.25, 1.62, 1.6, 0.3, "Step conversion", 12, rgb("7A4B82"), bold=True, smart=False)
cy16 = fy0 + 2 * (fh + fg) + fh / 2
kit.shape(s, MSO_SHAPE.RECTANGLE, 10.85, cy16 - 0.95, 1.9, 1.9, rgb("FBE3EC"))
txt(s, 11.0, cy16 - 0.9, 1.6, 0.55, "−21,000", 22, MAG, bold=True, smart=False)
txt(s, 11.0, cy16 - 0.35, 1.65, 1.2, "applicants lost when asked for two documents", 13, rgb("5A2340"))
source(s, "Note: applications January–June 2026. Source: website analytics (fictional data)", T["grey"])
page_no(s, 16, T["grey"])

# =============================================================================================== 17. roadmap
# Brand: Harbourline, an (invented) ferry operator. Slate structure, teal phases, amber decision gates.
SL, TL17, AMB17, LT17 = rgb("334155"), rgb("0F9D8A"), rgb("F59E0B"), rgb("E2E8F0")
s = kit.new_slide()
txt(s, 0.6, 0.5, 12.1, 1.0, "A three-phase plan gets the first two routes running in nine months, with a decision "
    "gate before each step up", 26, SL, bold=True)
qs = ["Q1 2027", "Q2", "Q3", "Q4", "Q1 2028", "Q2"]
gx, gw = 3.1, W - 0.6 - 3.1; qw = gw / 6; ty = 2.3
phases = [("PREPARE", 0, 1, rgb("94A3B8")), ("PILOT", 1, 3, TL17), ("SCALE", 3, 6, SL)]
for nm, a_, b_, col in phases:
    x0 = gx + a_ * qw; x1 = gx + b_ * qw
    freeform(s, [(x0, ty - 0.55), (x1 - 0.12, ty - 0.55), (x1 + 0.05, ty - 0.33), (x1 - 0.12, ty - 0.11), (x0, ty - 0.11),
                 (x0 + (0.17 if a_ else 0), ty - 0.33)], fill=col)
    txt(s, x0 + 0.15, ty - 0.55, x1 - x0 - 0.3, 0.44, nm, 14, WHITE, bold=True, align=C, anchor=MID, smart=False)
for i, q in enumerate(qs):
    txt(s, gx + i * qw, ty, qw, 0.3, q, 13, rgb("64748B"), bold=True, align=C, smart=False)
rows17 = [("Vessels and berths", [(0, 1.0, rgb("94A3B8")), (1.0, 3, TL17), (3, 6, SL)], [(0.95, "Leases signed", "on")]),
          ("Crew and training", [(0.3, 1.3, rgb("94A3B8")), (1.3, 3, TL17), (3.2, 5.5, SL)], [(1.3, "First crews ready", "on")]),
          ("Booking and ticketing", [(0, 1.5, rgb("94A3B8")), (1.5, 2.6, TL17), (3.4, 5, SL)], [(2.6, "App live", "left")]),
          ("Marketing", [(1.4, 3, TL17), (3.1, 6, SL)], [(1.5, "Launch", "on")])]
ry0, rh = ty + 0.45, 0.85
for k, (nm, bars, ms) in enumerate(rows17):
    y = ry0 + k * rh
    kit.shape(s, MSO_SHAPE.RECTANGLE, 0.6, y, W - 1.2, rh - 0.08, LT17 if k % 2 == 0 else rgb("F8FAFC"))
    txt(s, 0.75, y, gx - 0.9, rh - 0.08, nm, 16, SL, bold=True, anchor=MID)
    for a_, b_, col in bars:
        bb = kit.shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, gx + a_ * qw + 0.04, y + 0.2, (b_ - a_) * qw - 0.08, rh - 0.48, col)
        bb.adjustments[0] = 0.5
    for m_, lab, where in ms:
        cx = gx + m_ * qw
        d_ = s.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(cx - 0.15), Inches(y + 0.14), Inches(0.3), Inches(0.42))
        d_.fill.solid(); d_.fill.fore_color.rgb = WHITE; d_.line.color.rgb = SL; d_.line.width = Pt(2); kit.no_effects(d_)
        if where == "on":
            txt(s, cx + 0.22, y + 0.18, 1.8, 0.35, lab, 12, WHITE, bold=True, anchor=MID, smart=False)
        else:
            txt(s, cx - 0.22 - 1.6, y + 0.18, 1.6, 0.35, lab, 12, WHITE, bold=True, align=R, anchor=MID, smart=False)
gy_end = ry0 + 4 * rh
for gq, lab in [(1, "Gate 1: board go-ahead"), (3, "Gate 2: scale or stop")]:
    x = gx + gq * qw
    kit.line(s, x, ty + 0.35, x, gy_end + 0.1, AMB17, 2.25, dash=MSO_LINE_DASH_STYLE.DASH)
    pill(s, x - 1.15, gy_end + 0.15, 2.3, 0.36, AMB17, lab, 12, rgb("3A2A00"))
source(s, "Note: phases overlap where work can start early. Source: programme plan (fictional)", T["grey"])
page_no(s, 17, T["grey"])

# =============================================================================================== 18. KPI tiles
# Brand: Greenbank Leisure, an (invented) gym chain. Deep pine field, white tiles, green and red movement.
PINE, UP, DOWN = rgb("123A33"), rgb("1E9E6A"), rgb("D64545")
s = kit.new_slide()
kit.shape(s, MSO_SHAPE.RECTANGLE, 0, 0, W, H, PINE)
txt(s, 0.6, 0.45, 8, 0.3, "QUARTERLY DASHBOARD  •  Q3 2026", 13, rgb("9CD8C0"), bold=True, smart=False)
txt(s, 0.6, 0.8, 12.1, 1.0, "Membership is growing and churn has halved, but income per member is slipping", 28, WHITE, bold=True)
kpis = [("Members", "48,200", "▲ 12% on last year", True, [36, 38, 39, 41, 42, 44, 46, 48.2]),
        ("Monthly churn", "4.1%", "▼ from 8.2% last year", True, [8.2, 7.6, 7.1, 6.3, 5.6, 5.0, 4.5, 4.1]),
        ("Income per member", "£86", "▼ 6% on last year", False, [91.5, 92, 91, 90, 89, 88, 87, 86]),
        ("Net promoter score", "+41", "▲ 9 points", True, [32, 33, 35, 34, 37, 38, 40, 41])]
tw, tg = (W - 1.2 - 3 * 0.25) / 4, 0.25
for k, (lab, big, delta, good, series) in enumerate(kpis):
    x = 0.6 + k * (tw + tg); y = 2.25; th = 4.25
    b = kit.shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, tw, th, WHITE); b.adjustments[0] = 0.05
    txt(s, x + 0.3, y + 0.3, tw - 0.6, 0.35, lab, 16, rgb("4B5D58"), bold=True, smart=False)
    txt(s, x + 0.3, y + 0.75, tw - 0.6, 0.85, big, 44, PINE, bold=True, smart=False)
    pill(s, x + 0.3, y + 1.72, tw - 0.6, 0.38, rgb("E3F5EC") if good else rgb("FBE4E4"), delta, 12, UP if good else DOWN)
    cd = CategoryChartData(); cd.categories = [f"Q{(i % 4) + 1}" for i in range(8)]; cd.add_series(lab, series)
    gf = s.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(x + 0.25), Inches(y + 2.35), Inches(tw - 0.5), Inches(1.5), cd)
    ch = gf.chart; kit.transparent_chart(ch); ch.has_legend = False; ch.has_title = False
    ch.value_axis.visible = False; ch.category_axis.visible = False
    ch.value_axis.has_major_gridlines = False
    lo_, hi_ = min(series), max(series); pad = (hi_ - lo_) * 0.15
    ch.value_axis.minimum_scale = lo_ - pad; ch.value_axis.maximum_scale = hi_ + pad
    kit.plot_layout(ch, 0.02, 0.05, 0.96, 0.9)
    ser = ch.plots[0].series[0]; ser.format.line.color.rgb = UP if good else DOWN; ser.format.line.width = Pt(3); ser.smooth = True
    no_markers(ser)
    kit.alt(gf, f"{lab} over the last eight quarters, from {series[0]:g} to {series[-1]:g}.")
    txt(s, x + 0.3, y + th - 0.4, tw - 0.6, 0.28, "Last eight quarters", 11, rgb("8A9A95"), smart=False)
txt(s, 0.6, 6.95, 9, 0.3, "Source: membership system and quarterly survey (fictional data)", 10, rgb("9CD8C0"), anchor=BOT)
txt(s, W - 1.2, 6.95, 0.6, 0.3, "18", 10, rgb("9CD8C0"), align=R, anchor=BOT, smart=False)


# =============================================================================================== 19-20. executive summary
# Brand: Merrow Transit, an (invented) regional bus operator. Petrol ink, mint and amber accents.
# The same answer-first summary twice: on white (19) and on a deep colour field (20).
PETROL, MINT19, AMB19, INK19 = rgb("0F3D4C"), rgb("4FC1A6"), rgb("F2A93B"), rgb("1E2A30")
summary = [("1", "Two depots do the work of one", "The north-side depots are 3 miles apart and each runs at 55% capacity", "55%", "depot capacity used"),
           ("2", "Merging them frees £4M", "Selling the Hallam site and sharing one workshop saves £1.1M a year", "£4M", "one-off sale value"),
           ("3", "That pays for 12 electric buses", "Enough to electrify the two busiest routes, cutting fuel costs by a third", "12", "buses, two routes"),
           ("4", "Service and jobs are protected", "No routes cut; all 140 staff move to the merged depot", "0", "routes or jobs lost")]


def exec_summary(s, dark, n):
    bg = PETROL if dark else WHITE
    head = WHITE if dark else PETROL
    body = rgb("D5E6EA") if dark else INK19
    if dark:
        kit.shape(s, MSO_SHAPE.RECTANGLE, 0, 0, W, H, PETROL)
    txt(s, 0.6, 0.42, 6, 0.3, "EXECUTIVE SUMMARY", 14, MINT19, bold=True, smart=False)
    txt(s, 0.6, 0.75, 12.1, 1.0, "We recommend merging the two north-side depots and using the £4M it frees to buy "
        "12 electric buses", 26, head, bold=True)
    y0, rh, rg = 2.2, 0.86, 0.13
    for k, (num, lead, rest, big, cap) in enumerate(summary):
        y = y0 + k * (rh + rg)
        if dark:
            kit.shape(s, MSO_SHAPE.RECTANGLE, 0.6, y, W - 1.2, rh, rgb("16505F"))
        else:
            kit.shape(s, MSO_SHAPE.RECTANGLE, 0.6, y, W - 1.2, rh, rgb("F1F6F7"))
        kit.shape(s, MSO_SHAPE.RECTANGLE, 0.6, y, 0.09, rh, MINT19)
        kit.shape(s, MSO_SHAPE.OVAL, 0.95, y + (rh - 0.52) / 2, 0.52, 0.52, rgb("1F8A70"))
        txt(s, 0.95, y + (rh - 0.52) / 2, 0.52, 0.52, num, 18, WHITE, bold=True, align=C, anchor=MID, smart=False)
        txt(s, 1.75, y + 0.1, 7.6, 0.4, lead, 20, head, bold=True, smart=False)
        txt(s, 1.75, y + 0.48, 7.6, 0.38, rest, 16, body, smart=False)
        kit.line(s, 9.65, y + 0.16, 9.65, y + rh - 0.16, rgb("3E7383") if dark else rgb("C9D8DC"), 1)
        txt(s, 9.9, y + 0.04, 2.7, 0.55, big, 30, MINT19 if dark else rgb("1F8A70"), bold=True, smart=False)
        txt(s, 9.9, y + 0.56, 2.7, 0.3, cap, 13, body, smart=False)
    by = y0 + 4 * (rh + rg) + 0.04
    kit.shape(s, MSO_SHAPE.RECTANGLE, 0.6, by, W - 1.2, 0.62, AMB19)
    txt(s, 0.85, by, W - 1.7, 0.62, [[("Decision needed: ", True, rgb("3A2600")),
        ("approve the depot merger and the £4M bus order before the March budget", False, rgb("3A2600"))]], 17, rgb("3A2600"), anchor=MID)
    txt(s, W - 1.2, 6.95, 0.6, 0.3, str(n), 10, rgb("9FC3CC") if dark else T["grey"], align=R, anchor=BOT, smart=False)
    txt(s, 0.6, 6.95, 9, 0.3, "Source: depot utilisation 2026; property valuation; fleet plan (fictional)", 10,
        rgb("9FC3CC") if dark else T["grey"], anchor=BOT)


s = kit.new_slide(); exec_summary(s, False, 19)
s = kit.new_slide(); exec_summary(s, True, 20)


out = HERE / "showcase.pptx"
kit.save(str(out), title="Slidewell showcase")
print(out)
