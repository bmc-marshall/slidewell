"""Example: a steering-group update built INSIDE a corporate template (fictional company and data).

Shows template mode: Kit(template=...) reads the template's colours, fonts, headline position and footer, then
the deck uses the template's own layouts ("Cover", "Title only", "Title and content", "Section divider"), its
title and body placeholders (so headlines and bullets keep the template's formatting), and its native
slide-number field. Slidewell adds only the content: charts, tables, status chips, options and a decision bar.

The template itself is built by templates/make_brackwater_template.py, the way a real company's would be.
"""
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[0] / "scripts"))
from slidekit import Kit, CategoryChartData, rgb  # noqa: E402
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE  # noqa: E402
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402

TEMPLATE = HERE / "templates" / "brackwater-template.pptx"
if not TEMPLATE.exists():
    import runpy
    runpy.run_path(str(HERE / "templates" / "make_brackwater_template.py"), run_name="__main__")

kit = Kit(dict(grey="5E6A70", light="C9D1D4"), template=str(TEMPLATE))
T, M, W, CW = kit.T, kit.T["M"], kit.T["W"], kit.CW
MAIN, ACCENT, INK, WHITE = T["main"], T["accent"], T["ink"], rgb("FFFFFF")
RAG = {"On track": (rgb("DCEBD9"), rgb("2F6B2F")), "At risk": (rgb("F6E7C8"), rgb("8A6212")),
       "Off track": (rgb("F3D9D4"), rgb("9A3324"))}
TOP, BOTTOM, BAR_Y = kit.anchors()

# ---- figures (fictional), checked below so the slides can't disagree with each other
plan_km, actual_km = [50, 60, 50], [52, 58, 32]
budget = {"Mains renewal": (18.0, 11.8, 9.4), "Customer connections": (4.0, 3.4, 3.1),
          "Programme management": (2.0, 1.4, 1.2), "Contingency": (2.0, 0.2, 0.0)}   # (budget, plan to date, spend)
tot_b = sum(v[0] for v in budget.values()); tot_p = sum(v[1] for v in budget.values()); tot_s = sum(v[2] for v in budget.values())
assert sum(plan_km) == 160 and sum(actual_km) == 142 and sum(plan_km) - sum(actual_km) == 18
assert round(tot_b, 1) == 26.0 and round(tot_p, 1) == 16.8 and round(tot_s, 1) == 13.7 and round(tot_p - tot_s, 1) == 3.1
assert round(budget["Mains renewal"][1] - budget["Mains renewal"][2], 1) == 2.4


def content(n, headline, source=None, layout="Title only"):
    s = kit.new_slide(layout)
    kit.frame(s, n, None, headline, source)
    return s


def chip(s, x, y, status, w=1.45, h=0.42):
    fill, ink = RAG[status]
    b = kit.box(s, x, y, w, h, fill)
    kit.text(s, x, y, w, h, status, 14, ink, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    return b


# 1 ---- cover: the template's Cover layout, its title and subtitle placeholders
s = kit.new_slide("Cover")
kit.fill(s, "title", "Network renewal programme")
kit.fill(s, "subtitle", ["Steering group update, Q3 2026–27", "Illustrative example with fictional data"])

# 2 ---- executive summary: status by dimension, then the decision
s = content(2, "On budget but four weeks late: we need a decision on how to recover",
            "Note: status at 30 November 2026 against the baseline approved in April (fictional data).")
rows = [("Schedule", "Off track", "142 km renewed against 160 km planned, mainly because the Eastfield contractor started four weeks late"),
        ("Budget", "On track", "£3.1M under plan to date; £2.4M of that is timing, so the full-year forecast is unchanged"),
        ("Customers", "On track", "Supply interruptions are 12% below target, and complaints are flat"),
        ("Safety", "On track", "No lost-time incidents this quarter")]
rh, gap = 0.62, 0.16
block = 4 * rh + 3 * gap
y0 = TOP + (BAR_Y - TOP - block) / 2 - 0.05     # centred between the headline zone and the decision bar
for k, (dim, st, txt) in enumerate(rows):
    y = y0 + k * (rh + gap)
    kit.text(s, M, y, 2.2, rh, dim, 18, MAIN, bold=True, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    chip(s, M + 2.3, y + (rh - 0.42) / 2, st)
    kit.text(s, M + 4.1, y, CW - 4.1, rh, txt, 18, INK, anchor=MSO_ANCHOR.MIDDLE)
kit.action_bar(s, "Decision needed: approve a second Eastfield crew, funded from contingency (slide 6)", fill=MAIN)

# 3 ---- progress chart
s = content(3, "We've renewed 142 km of mains, 18 km behind plan, all of it in Q3",
            "Source: programme progress reports, April–November 2026 (fictional data).")
cd = CategoryChartData()
cd.categories = ["Q1", "Q2", "Q3"]
cd.add_series("Planned", plan_km)
cd.add_series("Renewed", actual_km)
cw, ch_h = 7.9, BOTTOM - TOP - 0.85
kit.chart_title(s, M, TOP, cw, "Kilometres of mains renewed per quarter, 2026–27")
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(M), Inches(TOP + 0.55), Inches(cw), Inches(ch_h), cd)
ch = gf.chart
kit.style_chart(ch); kit.transparent_chart(ch)
PX, PY, PW, PH = 0.1, 0.12, 0.88, 1 - 0.12 - 0.5 / ch_h
kit.plot_layout(ch, PX, PY, PW, PH)
plot = ch.plots[0]; plot.gap_width = 70; plot.overlap = -10
for ser, col in zip(plot.series, [kit.tint(MAIN, 0.7), MAIN]):
    ser.format.fill.solid(); ser.format.fill.fore_color.rgb = col; ser.format.line.fill.background()
    dl = ser.data_labels; dl.show_value = True; dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(14); dl.font.color.rgb = INK
q3 = plot.series[1].points[2]; q3.format.fill.solid(); q3.format.fill.fore_color.rgb = ACCENT
va = ch.value_axis; va.has_major_gridlines = False; va.minimum_scale = 0; va.maximum_scale = 70; va.major_unit = 10
kit.axis(va, "km renewed"); kit.axis(ch.category_axis, pt=14)
ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.TOP; ch.legend.include_in_layout = False
ch.legend.font.size = Pt(14)
kit.alt(gf, "Column chart of planned versus renewed kilometres of mains by quarter: Q1 50 planned and 52 renewed, "
            "Q2 60 and 58, Q3 50 and 32. The programme is 18 km behind plan, all of it in Q3.")
sx = M + cw + 0.5
kit.side_points(s, sx, W - M - sx, TOP + 0.55 + PY * ch_h, TOP + 0.55 + (PY + PH) * ch_h, [
    ("Q1–Q2: ", "ahead of plan in Q1 and 2 km behind in Q2"),
    ("Q3: ", "18 km short, because Eastfield's contractor started four weeks late"),
    ("Elsewhere: ", "the other three areas are on plan")], pt=16, lead_color=MAIN)

# 4 ---- budget table
s = content(4, "Spend is £3.1M below plan, but £2.4M of that is timing, not savings",
            "Source: finance system, spend to 30 November 2026; £ millions (fictional data).")
cols = ["Workstream", "Annual budget", "Plan to date", "Spend to date", "Variance to date"]
data = [[k, f"{b:.1f}", f"{p:.1f}", f"{sp:.1f}", f"{sp - p:+.1f}"] for k, (b, p, sp) in budget.items()]
data.append(["Total", f"{tot_b:.1f}", f"{tot_p:.1f}", f"{tot_s:.1f}", f"{tot_s - tot_p:+.1f}"])
tw, row_h = CW, 0.62
tbl = s.shapes.add_table(len(data) + 1, len(cols), Inches(M), Inches(TOP + 0.1), Inches(tw), Inches(row_h * (len(data) + 1))).table
widths = [4.4, 1.95, 1.95, 1.95, 1.883]
for j, wd in enumerate(widths):
    tbl.columns[j].width = Inches(wd)
tbl.first_row = True; tbl.horz_banding = False
for i in range(len(data) + 1):
    tbl.rows[i].height = Inches(row_h)
    for j in range(len(cols)):
        c = tbl.cell(i, j)
        txt = cols[j] if i == 0 else data[i - 1][j]
        c.text = txt
        c.margin_left = c.margin_right = Inches(0.12)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = c.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT
        r = p.runs[0]; r.font.size = Pt(16); r.font.name = T["body"]
        last = i == len(data)
        r.font.bold = i == 0 or last
        c.fill.solid()
        if i == 0:
            c.fill.fore_color.rgb = MAIN; r.font.color.rgb = WHITE
        else:
            c.fill.fore_color.rgb = T["main_pale"] if last else WHITE
            r.font.color.rgb = INK
        if i == 1 and j == 4:
            r.font.color.rgb = ACCENT; r.font.bold = True
kit.text(s, M, TOP + 0.1 + row_h * (len(data) + 1) + 0.3, CW, 0.8,
         "Mains renewal is £2.4M behind because Eastfield started late; that work moves into Q4, so the full-year forecast stays at £26.0M.", 16, INK)

# 5 ---- section divider: the template's own layout
s = kit.new_slide("Section divider")
kit.fill(s, "title", "Recovering the schedule")
kit.fill(s, "body", "Three ways to recover the Eastfield delay")

# 6 ---- options
s = content(6, "A second crew recovers the whole delay for £0.6M, within contingency",
            "Note: costs and recovery are the delivery team's estimates; ratings against the criteria the board agreed in April (fictional).")
crit = ["Extra cost", "Weeks recovered", "Delivery risk"]
opts = [("A. Second crew", ["£0.6M", "All 4", "Low"], ["At risk", "On track", "On track"], True),
        ("B. Overtime only", ["£0.3M", "2 of 4", "Medium"], ["On track", "At risk", "At risk"], False),
        ("C. Accept the delay", ["None", "0 of 4", "High"], ["On track", "Off track", "Off track"], False)]
lab_w, cell_w, cgap, rh2 = 3.3, 2.55, 0.15, 0.82
xs = [M + lab_w + k * (cell_w + cgap) for k in range(3)]
for k, cname in enumerate(crit):
    kit.text(s, xs[k], TOP, cell_w, 0.4, cname, 16, MAIN, bold=True, align=PP_ALIGN.CENTER, smart=False)
for r_, (name, vals, rags, rec) in enumerate(opts):
    y = TOP + 0.6 + r_ * (rh2 + 0.22)
    if rec:
        kit.box(s, M - 0.12, y - 0.1, lab_w + 3 * cell_w + 2 * cgap + 0.24, rh2 + 0.2, None, line=ACCENT, lw=2.0)
    kit.text(s, M, y, lab_w - 0.2, rh2, name, 18, INK, bold=rec, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    for k in range(3):
        fill, ink = RAG[rags[k]]
        kit.box(s, xs[k], y, cell_w, rh2, fill)
        kit.text(s, xs[k], y, cell_w, rh2, vals[k], 18, ink, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, smart=False)
kit.text(s, M, TOP + 0.6 + 3 * (rh2 + 0.22) + 0.15, CW, 0.7,
         "Option C keeps spend down, but the regulator's delivery target would be missed, with a likely penalty of about £0.4M.", 16, INK)

# 7 ---- risks: the template's body placeholder and its own bullet styles
s = content(7, "Three risks need the board's attention this quarter", layout="Title and content")
kit.fill(s, "body", None, bullets=[
    ("Winter weather could cost another two weeks in January", 0),
    ("Mitigation: front-load sheltered urban sites; the second crew gives us slack", 1),
    ("Eastfield's contractor is short of supervisors", 0),
    ("Mitigation: two supervisors seconded from the Northbank contract until March", 1),
    ("Pipe prices rise 6% from April under the supply contract's indexation", 0),
    ("Mitigation: bring forward the Q1 order; already allowed for in the forecast", 1)])
body = kit.placeholder(s, "body")
body_h = 6 * 0.5 + 0.35                       # 3 risks of 18pt + 3 mitigations of 16pt, with the template's spacing
kit.move(body, y=kit.place_balanced(body_h, s, source_lines=0), h=body_h)   # keeps the template's left edge and width

# 8 ---- decision
s = content(8, "We ask the board to approve a second crew, funded from contingency",
            "Note: contingency remaining after this decision: £1.4M of £2.0M (fictional).")
steps = [("Today", "Board decision"), ("By 15 Dec", "Second crew mobilised"), ("Jan–Feb", "Both crews on site"),
         ("End of March", "Back on plan: 210 km")]
ty = TOP + (BAR_Y - TOP) / 2 + 0.2
kit.line(s, M + 0.2, ty, W - M - 0.2, ty, T["light"], 3)
step_w = CW / len(steps)
for k, (when, what) in enumerate(steps):
    cx = M + step_w * (k + 0.5)
    kit.shape(s, MSO_SHAPE.OVAL, cx - 0.13, ty - 0.13, 0.26, 0.26, ACCENT if k == 0 else MAIN)
    kit.text(s, cx - step_w / 2, ty - 0.8, step_w, 0.4, when, 18, MAIN, bold=True, align=PP_ALIGN.CENTER, smart=False)
    kit.text(s, cx - step_w / 2 + 0.1, ty + 0.35, step_w - 0.2, 0.7, what, 18, INK, align=PP_ALIGN.CENTER)
kit.action_bar(s, "Decision: does the board approve a second Eastfield crew, funded with £0.6M from contingency?", fill=MAIN)

out = HERE / "brackwater-update.pptx"
kit.save(str(out))
print(out)
