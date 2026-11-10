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
freeform(s, [(0, 0), (1.05, 0), (0.55, H), (0, H)], fill=NAVY)                    # slanted brand band
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
kit.shape(s, MSO_SHAPE.RECTANGLE, 0.55, 6.62, W - 0.55, 0.88, NAVY)
freeform(s, [(0.55, 6.62), (0.7, 6.62), (0.55, H)], fill=NAVY)
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

out = HERE / "showcase.pptx"
kit.save(str(out), title="Slidewell showcase")
print(out)
