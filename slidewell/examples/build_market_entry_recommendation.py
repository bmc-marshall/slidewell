"""Draft deck B: analytics-led board recommendation (fictional company and data).
Palette: near-black brand brown for text and headlines, crema gold accent (bullets, focal data), roast brown for
highlight text, cream tints. Motif: the slanted band from the (fictional) logo's label. Navigation: header tag."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from slidekit import Kit, CategoryChartData, XyChartData, rgb
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION

ESP, ROAST, GOLD, CREAM = rgb("3B2A20"), rgb("8A5A3B"), rgb("C99A3E"), rgb("F5EFE6")
kit = Kit(dict(body="Arial", main="3B2A20", main_pale="F5EFE6", accent="C99A3E", ink="3B2A20", grey="7D746E",
               light="D9CFC4", muted_ink="B3A99F",
               tag_labels=("Executive summary", "Market growth", "Locations", "Store financials", "Strategic options",
                           "Risks", "Recommendation", "Methodology")))
T, CW, M, W, H = kit.T, kit.CW, kit.T["M"], kit.T["W"], kit.T["H"]
RAG = {"Good": (rgb("DCEBD9"), rgb("2F6B2F")), "Fair": (rgb("F6E7C8"), rgb("8A6212")), "Poor": (rgb("F3D9D4"), rgb("9A3324"))}


def slant(s, x, y, w, h, fill, cut=0.3):
    """Motif: a band with a straight slanted right end (like the label on the cup)."""
    fb = s.shapes.build_freeform(Inches(x), Inches(y), scale=1.0)
    fb.add_line_segments([(Inches(x + w), Inches(y)), (Inches(x + w - cut), Inches(y + h)), (Inches(x), Inches(y + h))], close=True)
    shp = fb.convert_to_shape(); shp.fill.solid(); shp.fill.fore_color.rgb = fill; shp.line.fill.background(); kit.no_effects(shp)
    return shp


def slide(n, tag, headline, source=None):
    s = kit.new_slide()
    kit.frame(s, n, None, headline, source)
    slant(s, 0, 0, kit.tag_w + 0.3, T["tag_h"], ESP)                      # motif role 2: slanted header tag, fixed size
    kit.text(s, M, 0, kit.tag_w - M, T["tag_h"], tag.upper(), 11, rgb("FFFFFF"), bold=True, anchor=MSO_ANCHOR.MIDDLE, wrap=False, smart=False)
    return s


def bookend(s, title, sub, small):
    kit.shape(s, MSO_SHAPE.RECTANGLE, 0, 0, W, H, ESP)
    slant(s, 0, 4.9, 9.2, 0.5, GOLD, cut=0.5)                                # motif role 1: large band on title/closing
    th = kit.text_height(kit.wrapped(title, 40, 9.5, True), 40, 9.5, True)
    y = 4.9 - 0.35 - 0.5 - th
    kit.text(s, M, y, 9.5, th + 0.05, title, 40, rgb("FFFFFF"), bold=True)
    kit.text(s, M, y + th + 0.12, 9.5, 0.45, sub, 20, CREAM)
    if small:
        kit.text(s, M, 5.6, 9.5, 0.4, small, 14, kit.tint(GOLD, 0.35))


CH_H = 3.5                      # the same chart group on every chart slide, so charts line up from slide to slide
TITLE_GAP = 0.7                 # chart title to chart frame: clear space between a header and what it heads
CHART_TOP = None                # set from the first chart slide; slides with the same headline length reuse it
AX = 14 * 1.2 / 72 + 0.12       # chart titles start at the axis labels, not at the rotated axis title


def panel(s, rx, top, pw, title, points, plot_top, plot_bottom):
    """Takeaways beside a chart: header level with the chart title, points spread down the plotted data."""
    kit.chart_header(s, rx, top, pw, title)
    kit.side_points(s, rx, pw, plot_top, plot_bottom, points, lead_color=ROAST)


def italic(tb):
    for p in tb.text_frame.paragraphs:
        for r in p.runs:
            r.font.italic = True


kit.chart_header = lambda s, x, y, w, t_, pt=18, align=PP_ALIGN.LEFT: kit.text(s, x, y, w, 0.45, t_, pt, ROAST, bold=True, align=align)

# ---------------------------------------------------------------- accuracy checks for every derived figure
cagr = lambda a, b, n: (b / a) ** (1 / n) - 1
assert round(cagr(120, 270, 10), 2) == 0.08 and round(cagr(210, 280, 10), 2) == 0.03   # "about 8%" vs "about 3% a year"
assert cagr(120, 270, 10) > 2 * cagr(210, 280, 10)            # "more than twice as fast"
assert 1000 - 310 - 280 - 170 - 100 == 140 and 140 / 1000 == 0.14   # waterfall and "14% margin"
assert round(170 / 1000, 2) == 0.17                            # "rent 17% of revenue"
assert 0.8 * 140 == 112 and 350 - 112 - 140 == 98 and round(98 / 140, 1) == 0.7   # payback build-up, 2.7 years
assert round(0.7 * 12) == 8                                    # "Year 3 profit (8 months)", "2 years, 8 months"
assert 6 * 350 + 300 == 2400                                   # "£2.4M = six fit-outs at £350K + £0.3M"

# ---------------------------------------------------------------- 1 title
s = kit.new_slide()
bookend(s, "Should Harbor Roast open in Northfield?", "Recommendation to the board",
        "February 2027  |  Illustrative example with fictional data")

# ---------------------------------------------------------------- 2 executive summary
s = slide(2, "Executive summary", "We recommend a six-store pilot in Northfield's two university districts from March 2027")
rows = [("Fast-growing market", "Northfield's specialty coffee market is growing more than twice as fast as our home region"),
        ("Profitable stores", "A Northfield store would earn a 14% margin and pay back its fit-out in under three years"),
        ("Pilot is the best route", "A six-store pilot balances value, risk and brand control better than a rollout or franchising"),
        ("£2.4M to approve", "Approve £2.4M to open six stores between March and September 2027")]
rh, gap = 0.95, 0.18
top = kit.place_balanced(4 * rh + 3 * gap)
for i, (lab, txt) in enumerate(rows):
    y = top + i * (rh + gap)
    slant(s, M, y, 3.7, rh, ESP if i < 3 else GOLD, cut=0.25)              # motif role 3: label bands; the ask in gold
    kit.text(s, M + 0.3, y, 3.1, rh, lab, 18, rgb("FFFFFF"), bold=True, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    kit.text(s, M + 4.0, y, CW - 4.4, rh, txt, 20, T["ink"], bold=(i == 3), anchor=MSO_ANCHOR.MIDDLE)

# ---------------------------------------------------------------- 3 market: line with forecast
s = slide(3, "Market growth", "Northfield's specialty coffee market is growing more than twice as fast as our home region",
          "Note: 2026–2030 are forecasts (E). Specialty coffee = independent and premium chain coffee shop spend. Source: Regional Consumer Spend Survey 2025; team forecast (fictional data)")
CHART_TOP = kit.place_balanced(TITLE_GAP + CH_H)    # set on the most constrained chart slide (two-line notes)...
top = CHART_TOP
cw = 8.4
kit.chart_header(s, M + AX, top, cw - AX, "Specialty coffee spend by region, 2020–2030E")
years = [str(y) for y in range(2020, 2026)] + [f"{y}E" for y in range(2026, 2031)]
north = [120, 131, 143, 156, 170, 185, 200, 216, 233, 251, 270]
home = [210, 218, 226, 233, 240, 247, 254, 260, 267, 273, 280]
cd = CategoryChartData(); cd.categories = years
cd.add_series("Northfield", north[:6] + [None] * 5); cd.add_series("Northfield forecast", [None] * 5 + north[5:])
cd.add_series("Home", home[:6] + [None] * 5); cd.add_series("Home forecast", [None] * 5 + home[5:])
cy, chh = top + TITLE_GAP, CH_H
ch = s.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(M), Inches(cy), Inches(cw), Inches(chh), cd).chart
kit.style_chart(ch)
PX, PY, PW, PH = 0.1, 0.04, 0.78, 1 - 0.04 - 0.55 / chh
kit.plot_layout(ch, PX, PY, PW, PH)
for ser, col, w, dash in zip(ch.plots[0].series, [GOLD, GOLD, ROAST, ROAST], [3, 3, 2.25, 2.25], [None, 4, None, 4]):
    ser.format.line.color.rgb = col; ser.format.line.width = Pt(w); ser.smooth = False
    if dash:
        ser.format.line.dash_style = dash
    ser.marker.style = None
    from lxml import etree
    m_ = ser._element.find("{http://schemas.openxmlformats.org/drawingml/2006/chart}marker")
    if m_ is None:
        m_ = etree.SubElement(ser._element, "{http://schemas.openxmlformats.org/drawingml/2006/chart}marker")
    etree.SubElement(m_, "{http://schemas.openxmlformats.org/drawingml/2006/chart}symbol").set("val", "none")
va = ch.value_axis; va.has_major_gridlines = False; va.minimum_scale = 0; va.maximum_scale = 300; va.major_unit = 100
kit.axis(va, "Spend (£M)"); kit.axis(ch.category_axis, pt=12)
n = len(years)
X = lambda i: M + (PX + PW * (i + 0.5) / n) * cw
Y = lambda v: cy + (PY + PH) * chh - v / 300 * PH * chh
dx = X(5)                                  # divider on the last actual year, exactly where the dashed forecasts begin
kit.line(s, dx, cy + PY * chh, dx, cy + (PY + PH) * chh, T["grey"], 0.75, dash=4)
kit.text(s, dx + 0.08, cy + PY * chh, 1.4, 0.3, "Forecast →", 12, T["grey"], smart=False)
kit.line_label(s, X(10) + 0.12, Y(280) - 0.12, "Home", 16, ROAST)
kit.line_label(s, X(10) + 0.12, Y(270) + 0.14, "Northfield", 16, GOLD)
rx = M + cw + 0.5
panel(s, rx, top, W - M - rx, "Growth, 2020–2030E", [
    ("Northfield:", "about 8% a year, to £270M by 2030"),
    ("Home region:", "about 3% a year"),
    ("By 2030:", "Northfield would be nearly the size of our home market")], cy + PY * chh, cy + (PY + PH) * chh)

# ---------------------------------------------------------------- 4 districts: scatter with target zone
s = slide(4, "Locations", "Two university districts offer high footfall and the least competition",
          "Note: footfall is the average weekday pedestrian count on main streets. Source: council footfall counters 2025; café listings, January 2027 (fictional data)")
top = CHART_TOP                                     # ...and reused, so chart and panel headers sit in the same place
cw = 8.4
kit.chart_header(s, M + AX, top, cw - AX, "Northfield districts: footfall vs specialty cafés")
pts = [("University Quarter", 48, 0.8, True), ("Riverside Campus", 41, 1.1, True), ("Old Town", 52, 3.8, False),
       ("Station", 45, 3.1, False), ("Retail Park", 30, 1.3, False), ("Eastgate", 22, 1.8, False),
       ("Westbrook", 18, 1.4, False), ("Market Hill", 35, 2.6, False)]
cy, chh = top + TITLE_GAP, CH_H
PX, PY, PW, PH = 0.11, 0.04, 0.84, 1 - 0.04 - 0.75 / chh
X = lambda v: M + (PX + PW * v / 60) * cw
Y = lambda v: cy + (PY + PH) * chh - v / 4.5 * PH * chh
kit.shape(s, MSO_SHAPE.RECTANGLE, X(38), Y(2.0), X(60) - X(38), Y(0) - Y(2.0), kit.tint(GOLD, 0.82))     # target zone behind the chart
kit.text(s, X(60) - 2.1, Y(2.0) + 0.08, 2.0, 0.28, "Target zone", 12, ROAST, bold=True, align=PP_ALIGN.RIGHT, smart=False)   # inside, top-right corner
cd = XyChartData()
for flag in (False, True):
    se = cd.add_series("Focus" if flag else "Other")
    for nm, x, y, f in pts:
        if f == flag:
            se.add_data_point(x, y)
ch = s.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER, Inches(M), Inches(cy), Inches(cw), Inches(chh), cd).chart
kit.style_chart(ch); kit.plot_layout(ch, PX, PY, PW, PH); kit.transparent_chart(ch)
from pptx.enum.chart import XL_MARKER_STYLE
for ser, col, size in zip(ch.plots[0].series, [T["light"], GOLD], [12, 16]):
    ser.marker.style = XL_MARKER_STYLE.CIRCLE; ser.marker.size = size
    ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = col; ser.marker.format.line.fill.background()
    ser.format.line.fill.background()
va = ch.value_axis; va.has_major_gridlines = False; va.minimum_scale = 0; va.maximum_scale = 4.5; va.major_unit = 1
kit.axis(va, "Cafés per 10,000 residents")
xa = ch.category_axis; xa.minimum_scale = 0; xa.maximum_scale = 60; xa.major_unit = 10; xa.has_major_gridlines = False
kit.axis(xa, "Weekday footfall (thousands a day)")
# label side per point: r = right of the point, l = left, b = below (centred). No-wrap boxes are avoided: renderers re-centre them.
side = {"University Quarter": "b", "Riverside Campus": "a", "Old Town": "l", "Retail Park": "b", "Westbrook": "l"}
for nm, x, y, f in pts:
    pt_, lw = (16, 2.3) if f else (14, 1.6)
    sd = side.get(nm, "r")
    if sd == "r":
        bx, by, al = X(x) + 0.16, Y(y) - 0.15, PP_ALIGN.LEFT
    elif sd == "l":
        bx, by, al = X(x) - 0.16 - lw, Y(y) - 0.15, PP_ALIGN.RIGHT
    elif sd == "a":
        bx, by, al = X(x) - lw / 2, Y(y) - 0.45, PP_ALIGN.CENTER
    else:
        bx, by, al = X(x) - lw / 2, Y(y) + 0.13, PP_ALIGN.CENTER
    kit.text(s, bx, by, lw, 0.3, nm, pt_, ESP if f else T["grey"], bold=f, align=al, smart=False)
rx = M + cw + 0.5
panel(s, rx, top, W - M - rx, "Why start there", [
    ("Footfall:", "41–48 thousand people a day, among the highest in Northfield"),
    ("Competition:", "the fewest specialty cafés per resident in Northfield"),
    ("Customers:", "students match our home-region core customer")], cy + PY * chh, cy + (PY + PH) * chh)

# ---------------------------------------------------------------- 5 economics: waterfall
s = slide(5, "Store financials", "A Northfield store would earn a 14% margin, above our 11% average",
          "Note: steady-state year (year 2 onwards), per store, £ thousands. Source: home-region store P&Ls 2026; Northfield rent quotes (fictional data)")
top = CHART_TOP
cw = 8.4
kit.chart_header(s, M + AX, top, cw - AX, "Annual profit bridge for one store (£K)")
steps = [("Revenue", 1000, "total"), ("Coffee & food", -310, "neg"), ("Staff", -280, "neg"), ("Rent", -170, "neg"),
         ("Other", -100, "neg"), ("Operating profit", 140, "result")]
base, vis, run = [], [], 0
for nm, v, kind in steps:
    if kind == "total":
        base.append(0); vis.append(v); run = v
    elif kind == "neg":
        run += v; base.append(run); vis.append(-v)
    else:
        base.append(0); vis.append(v)
cd = CategoryChartData(); cd.categories = [n_ for n_, _, _ in steps]
cd.add_series("base", base); cd.add_series("value", vis)
cy, chh = top + TITLE_GAP, CH_H
ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(M), Inches(cy), Inches(cw), Inches(chh), cd).chart
kit.style_chart(ch)
PY5, PH5 = 0.04, 1 - 0.04 - 0.75 / chh
kit.plot_layout(ch, 0.1, PY5, 0.88, PH5)
plot = ch.plots[0]; plot.gap_width = 55; plot.overlap = 100
b_ser, v_ser = plot.series[0], plot.series[1]
b_ser.format.fill.background(); b_ser.format.line.fill.background()
for i, (nm, v, kind) in enumerate(steps):
    p_ = v_ser.points[i]; p_.format.fill.solid()
    p_.format.fill.fore_color.rgb = ESP if kind == "total" else (GOLD if kind == "result" else kit.tint(ROAST, 0.35))
dl = v_ser.data_labels; dl.show_value = True; dl.position = XL_LABEL_POSITION.CENTER;   # centred in the bar dl.font.size = Pt(14); dl.font.bold = True
dl.font.color.rgb = rgb("FFFFFF")
va = ch.value_axis; va.has_major_gridlines = False; va.minimum_scale = 0; va.maximum_scale = 1100; va.major_unit = 250
kit.axis(va, "£ thousands per year"); kit.axis(ch.category_axis)
rx = M + cw + 0.5
panel(s, rx, top, W - M - rx, "Versus our home stores", [
    ("Margin:", "14% vs an 11% home average"),
    ("Rent:", "17% of revenue, three points below home"),
    ("Staff:", "in line with home, at local wage rates")], cy + PY5 * chh, cy + (PY5 + PH5) * chh)

# ---------------------------------------------------------------- 6 payback: build-up in the slide 5 style
s = slide(6, "Store financials", "Each pilot store would pay back its £350K fit-out in under three years",
          "Note: year 1 assumes 80% of steady-state profit while the store builds custom; no discounting. Source: team model (fictional data)")
top = kit.place_balanced(TITLE_GAP + CH_H - 0.25)   # one-line headline: sits higher (accepted); the bar notes end above the frame bottom
cw = CW
kit.chart_header(s, M + AX, top, cw - AX, "Fit-out cost vs cumulative profit for one store (£K)")
steps6 = [("Fit-out cost", 0, 350, ESP, "Build and equipment"),
          ("Year 1 profit", 0, 112, GOLD, "80% of steady state"),
          ("Year 2 profit", 112, 140, GOLD, "Steady state (slide 5)"),
          ("Year 3 profit (8 months)", 252, 98, GOLD, "0.7 of a £140K year")]
cd = CategoryChartData(); cd.categories = [c[0] for c in steps6]
cd.add_series("base", [c[1] for c in steps6]); cd.add_series("value", [c[2] for c in steps6])
cy, chh = top + TITLE_GAP, CH_H
ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(M), Inches(cy), Inches(cw), Inches(chh), cd).chart
kit.style_chart(ch)
PX, PY, PW, PH = 0.07, 0.04, 0.76, 1 - 0.04 - 0.95 / chh      # room under the category labels for the notes
kit.plot_layout(ch, PX, PY, PW, PH)
plot = ch.plots[0]; plot.gap_width = 70; plot.overlap = 100
b_ser, v_ser = plot.series[0], plot.series[1]
b_ser.format.fill.background(); b_ser.format.line.fill.background()
for i, c in enumerate(steps6):
    p_ = v_ser.points[i]; p_.format.fill.solid(); p_.format.fill.fore_color.rgb = c[3]
dl = v_ser.data_labels; dl.show_value = True; dl.position = XL_LABEL_POSITION.CENTER
dl.font.size = Pt(16); dl.font.bold = True; dl.font.color.rgb = rgb("FFFFFF")   # white on gold, as on slide 5
f0 = v_ser.points[0].data_label.font; f0.color.rgb = rgb("FFFFFF"); f0.bold = True; f0.size = Pt(16)
va = ch.value_axis; va.has_major_gridlines = False; va.minimum_scale = 0; va.maximum_scale = 400; va.major_unit = 100
kit.axis(va, "£ thousands"); kit.axis(ch.category_axis)
n = len(steps6)
X = lambda i: M + (PX + PW * (i + 0.5) / n) * cw
Y = lambda v: cy + (PY + PH) * chh - v / 400 * PH * chh
colw6 = PW * cw / n
for i, c in enumerate(steps6):        # notes centred under the bar they explain
    kit.text(s, X(i) - colw6 / 2 + 0.1, cy + (PY + PH) * chh + 0.42, colw6 - 0.2, 0.3, c[4], 14, T["grey"], align=PP_ALIGN.CENTER, smart=False)
kit.line(s, M + PX * cw, Y(350), M + (PX + PW) * cw, Y(350), T["grey"], 1.0, dash=4)
lab = kit.text(s, M + (PX + PW) * cw + 0.12, Y(350) - 0.3, 2.1, 0.6, "Paid back after\v2 years, 8 months", 16, ESP, bold=True,
               anchor=MSO_ANCHOR.MIDDLE, smart=False)      # centred on the line it labels

# ---------------------------------------------------------------- 7 options: RAG matrix
s = slide(7, "Strategic options", "A six-store pilot best balances value, risk and brand control",
          "Note: ratings are the team's assessment against the board's criteria agreed in November 2026 (fictional).")
opts = [("Six-store pilot", ["Fair", "Good", "Good", "Good"], "Good", True),
        ("20-store rollout", ["Good", "Poor", "Fair", "Good"], "Fair", False),
        ("Franchise partner", ["Fair", "Good", "Good", "Poor"], "Fair", False),
        ("Do nothing", ["Poor", "Good", "Good", "Good"], "Poor", False)]
crit = ["Value", "Risk", "Speed", "Brand control", "Overall"]
x0 = M + 0.5                                   # leaves room for the rotated label left of the outline
tw = W - M - 0.12 - x0
c0w, cw_ = 3.0, (tw - 3.0) / 5
rh = 0.95
top = kit.place_balanced(0.5 + 4 * rh)
for j, c_ in enumerate(crit):
    kit.text(s, x0 + c0w + j * cw_, top, cw_, 0.45, c_, 18, ROAST, bold=True, align=PP_ALIGN.CENTER, smart=False)
for i, (nm, ratings, overall, rec) in enumerate(opts):
    y = top + 0.5 + i * rh
    kit.text(s, x0, y, c0w - 0.2, rh, nm, 18, ESP, bold=rec, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    for j, r_ in enumerate(ratings + [overall]):
        fill, tc = RAG[r_]
        kit.shape(s, MSO_SHAPE.RECTANGLE, x0 + c0w + j * cw_ + 0.06, y + 0.06, cw_ - 0.12, rh - 0.12, fill)
        kit.text(s, x0 + c0w + j * cw_, y, cw_, rh, r_, 16, tc, bold=(j == 4), align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    if rec:   # the label says what the outline means, right where it is: no key to decode
        kit.shape(s, MSO_SHAPE.RECTANGLE, x0 - 0.12, y, W - M - (x0 - 0.12), rh, None, line=GOLD, lw=2.25)
        kit.vertical_label(s, x0 - 0.12 - 0.08, y + rh / 2, "RECOMMENDED", 10, ROAST, side="left")

# ---------------------------------------------------------------- 8 risks
s = slide(8, "Risks", "Main risks are manageable if we sign leases early and hire locally",
          "Note: likelihood is the team's judgement after mitigation (fictional).")
risks = [("Leases in the two districts are taken by competitors", 2, "Sign heads of terms for four sites by March; two backups identified"),
         ("We can't hire experienced managers locally", 2, "Transfer two home-region managers for year 1; train local deputies"),
         ("Student demand drops in summer holidays", 3, "Summer menu and opening hours; target staff and visitors on campus"),
         ("Fit-out costs overrun", 1, "Fixed-price contracts; 10% contingency inside the £2.4M")]
cols_ = [("Risk", 4.6), ("Likelihood", 1.9), ("Mitigation", CW - 6.5)]
rh = 0.85
top = kit.place_balanced(0.4 + 0.55 + 4 * rh - 0.25) + 0.4      # legend sits just above the table
leg = kit.text(s, M, top - 0.38, CW, 0.3, "Likelihood: ● low   ●● medium   ●●● high", 12, T["grey"], align=PP_ALIGN.RIGHT, smart=False)
italic(leg)
x_ = M
for nm, w_ in cols_:
    kit.text(s, x_, top, w_, 0.45, nm, 18, ROAST, bold=True, align=PP_ALIGN.CENTER if nm == "Likelihood" else PP_ALIGN.LEFT, smart=False)
    x_ += w_
for i, (risk, lik, mit) in enumerate(risks):
    y = top + 0.55 + i * rh
    if i % 2 == 0:
        kit.shape(s, MSO_SHAPE.RECTANGLE, M - 0.12, y, CW + 0.24, rh, CREAM)
    kit.text(s, M, y, 4.4, rh, risk, 16, ESP, anchor=MSO_ANCHOR.MIDDLE)
    for k in range(3):
        cx_ = M + 4.6 + 1.9 / 2 + (k - 1) * 0.32
        kit.shape(s, MSO_SHAPE.OVAL, cx_ - 0.1, y + rh / 2 - 0.1, 0.2, 0.2, GOLD if k < lik else T["light"])
    kit.text(s, M + 6.5, y, CW - 6.5, rh, mit, 16, ESP, anchor=MSO_ANCHOR.MIDDLE)

# ---------------------------------------------------------------- 9 recommendation: timeline + decision
s = slide(9, "Recommendation", "Approve £2.4M to open six stores between Mar and Sep 2027",
          "Note: £2.4M = six fit-outs at £350K (£2.1M) + £0.3M launch marketing; 10% contingency held within fit-out budgets (fictional).")
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
top = kit.place_balanced(4.15)
kit.chart_header(s, M, top, CW, "Pilot plan, 2027")
tx = M + 0.75; step = (CW - 1.5) / 11
ty = top + 1.6
trk = kit.shape(s, MSO_SHAPE.PENTAGON, M, ty - 0.08, CW, 0.16, kit.tint(GOLD, 0.7)); trk.adjustments[0] = 0.5   # pointed end: the pilot runs on into 2028
for i, m in enumerate(months):
    kit.text(s, tx + i * step - 0.4, ty + 0.2, 0.8, 0.32, m, 14, ESP, align=PP_ALIGN.CENTER, smart=False)   # secondary: not bold
events = {0: "Sign leases", 2: "Stores 1–2 open", 5: "Stores 3–4 open", 8: "Stores 5–6 open", 11: "Board review"}
for i, lab in events.items():
    x = tx + i * step
    kit.shape(s, MSO_SHAPE.OVAL, x - 0.1, ty - 0.1, 0.2, 0.2, GOLD if i in (2, 5, 8) else ESP, line=rgb("FFFFFF"), lw=1.5)
    kit.text(s, x - 0.75, ty - 1.15, 1.5, 0.62, lab, 16, ESP, bold=i in (2, 5, 8), align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM)
    kit.arrow(s, x, ty - 0.48, x, ty - 0.16)
# phases as motif bands (slanted ends), one level each
for k, (a_, b_, lab, fill, tc) in enumerate([(0, 7, "Leases and fit-out", CREAM, ESP), (2, 11, "Stores trading; monthly sales tracked against plan", GOLD, rgb("FFFFFF"))]):
    y = ty + 0.62 + k * 0.46
    x0 = tx + a_ * step - step / 2; x1 = tx + b_ * step + step / 2
    slant(s, x0, y, x1 - x0, 0.38, fill, cut=0.18)
    kit.text(s, x0 + 0.15, y, x1 - x0 - 0.4, 0.38, lab, 14, tc, bold=True, anchor=MSO_ANCHOR.MIDDLE, smart=False)
kit.bar(s, ty + 1.7, "Decision: Does the board approve £2.4M for a six-store Northfield pilot, opening March to September 2027?", fill=ESP)

# ---------------------------------------------------------------- 10 appendix
s = slide(10, "Methodology", "We built the case from regional spend data, district counts, and our own store accounts",
          "Source: Regional Consumer Spend Survey 2025; council footfall counters 2025; café listings Jan 2027; home-region store P&Ls 2026 (all fictional)")
steps_ = [("Market size", "Regional spend survey 2020–2025; forecast extends each region's recent growth, tapering by 0.2 points a year"),
          ("District screen", "Footfall from council counters; competition from café listings per 10,000 residents"),
          ("Store economics", "Home-region P&Ls adjusted for Northfield rents and wage rates"),
          ("Option appraisal", "Each option scored against the board's four criteria; overall rating weights value and risk most")]
row = 1.0
top = kit.place_balanced(row * 4 - 0.35)
for i, (t_, d) in enumerate(steps_):
    y = top + i * row
    kit.badge(s, M + 0.3, y + 0.3, 0.6, str(i + 1), pt=16, ring_color=ESP)
    kit.text(s, M + 0.9, y, 2.8, 0.6, t_, 18, ESP, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    kit.text(s, M + 3.8, y, CW - 4.2, 0.6, d, 18, ESP, anchor=MSO_ANCHOR.MIDDLE)

print(kit.save(str(Path(__file__).with_name("market-entry-recommendation.pptx"))))
