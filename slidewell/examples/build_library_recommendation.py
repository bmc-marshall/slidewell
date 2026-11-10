"""Worked example v5: Riverside Library (fictional), built with slidekit."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from slidekit import Kit, CategoryChartData, XyChartData, rgb
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION

THEME = dict(body="Arial", display="Cambria", main="0F4C5C", main_pale="E3EFF1", mid="3E7C8A", accent="E36414",
             ink="222222", grey="7A8088", light="C6CBD1", muted_ink="A9AEB4",
             tag_labels=("Executive summary", "Contents", "Visits by time", "Visitor needs", "Staff capacity", "Cost and impact", "Pilot plan", "Success measures", "Methodology"))
kit = Kit(THEME)
T, CW, M, W, H = kit.T, kit.CW, kit.T["M"], kit.T["W"], kit.T["H"]
MID = rgb("3E7C8A"); SHADE = rgb("FBDDC9")


def ribbon(s, x, top, w, h, color=None, notch_color=None):
    """Motif: bookmark ribbon hanging from a top edge, with a notched (triangle cut-out) end."""
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, top, w, h, color or T["accent"])
    n = kit.shape(s, MSO_SHAPE.ISOSCELES_TRIANGLE, x, top + h - w * 0.5, w, w * 0.5, notch_color or T["white"])
    n.rotation = 180


def content_slide(n, tag, headline, source):
    s = kit.new_slide()
    kit.frame(s, n, tag, headline, source)
    ribbon(s, W - M - 0.3, 0, 0.3, 0.62)          # motif role 2: small bookmark on every content slide
    return s


# ---------------------------------------------------------------- storyline
HEAD_EXEC = "A demand-led timetable lets our staff support ~20% more visitors"
SECTIONS = ["Demand has moved to evenings", "A new timetable fits staff to visitors", "Two decisions for the board"]
ROWS = [["62% of visits now happen after 4pm, up from 35% in 2019", "Most evening visitors are students who come to study"],
        ["Moving two daytime shifts into the evening would cover the peak",
         "The change is cost-neutral and lets staff support ~20% more visitors"],
        ["Approve a six-month pilot of the new timetable from March", "Agree three measures of success before the pilot starts"]]


def exec_rows(s, active=None):
    rh, gap = 1.4, 0.2
    block = 3 * rh + 2 * gap
    top = kit.place_balanced(block)
    px, pw = M + 0.3, 3.9
    for i, pts in enumerate(ROWS):
        y = top + i * (rh + gap)
        on = active is None or active == i
        kit.shape(s, MSO_SHAPE.RECTANGLE, px, y, pw, rh, T["main_pale"] if on else T["muted_fill"])
        kit.badge(s, px, y + rh / 2, 0.56, str(i + 1), pt=16, on=on)
        kit.text(s, px + 0.5, y, pw - 0.7, rh, SECTIONS[i], 18, T["main"] if on else T["muted_main"], bold=True, anchor=MSO_ANCHOR.MIDDLE)
        bx = px + pw + 0.35
        kit.bullet_list(s, bx, y, (W - M - 0.4) - bx, rh, pts, pt=18,
                        color=T["ink"] if on else T["muted_ink"],                 # black fades to grey
                        bullet_color=T["main"] if on else T["muted_main"],        # teal fades to pale teal
                        anchor=MSO_ANCHOR.MIDDLE)
        if active is not None and on:
            ribbon(s, px + pw - 0.5, y, 0.26, 0.5, notch_color=T["main_pale"])   # motif role 3: marks the current section


def tracker(n, active):
    s = content_slide(n, "Contents", HEAD_EXEC, None)
    exec_rows(s, active)


def bookend(s, title, sub, small):
    kit.shape(s, MSO_SHAPE.RECTANGLE, 0, 0, W, H, T["main"])
    ribbon(s, W - 2.2, 0, 0.7, 3.2, notch_color=T["main"])                     # motif role 1: large on title/closing
    lines = [(title, 44, T["white"], True, T["display"]), (sub, 20, T["white"], False, T["body"])]
    if small:
        lines.append((small, 16, T["main_pale"], False, T["body"]))
    hs = [kit.text_height(kit.wrapped(t_, pt, 9.5, b, f), pt, 9.5, b, f) for t_, pt, _, b, f in lines]
    y = (H - (sum(hs) + 0.2 * (len(lines) - 1))) / 2
    for (t_, pt, col, b, f), h in zip(lines, hs):
        kit.text(s, M, y, 9.5, h + 0.05, t_, pt, col, bold=b, font=f)
        y += h + 0.2


def chart_with_panel(s, chart_w, title, panel_title, points, build_chart, panel_pt=18):
    """Chart (header + plot) on the left, takeaways on the right, the whole block balanced vertically."""
    ch_h = 3.6                                      # same chart group on every chart slide, so they line up
    top = kit.place_balanced(0.7 + ch_h)            # group distributed between headline and notes
    ch_top = top + 0.7                              # clear space between the chart title and the chart
    plot_top, plot_bottom, axis_title = build_chart(ch_top, ch_h)
    kit.chart_title(s, M, top, chart_w, title, left_axis_title=True)   # same x on every chart slide, even without a left axis title
    rx = M + chart_w + 0.6
    pw = W - M - rx
    kit.chart_header(s, rx, top, pw, panel_title)  # panel header level with the chart title
    kit.side_points(s, rx, pw, plot_top, plot_bottom, points, pt=panel_pt)


# ---------------------------------------------------------------- 1 title, 2 summary, 3 contents
s = kit.new_slide()
bookend(s, "Opening hours that follow our visitors", "Proposal to the Riverside Library board",
        "March board meeting  |  Illustrative example with fictional data")
s = content_slide(2, "Executive summary", HEAD_EXEC, None)
exec_rows(s)
tracker(3, 0)

# ---------------------------------------------------------------- 4 time bands
s = content_slide(4, "Visits by time", "62% of visits now happen after 4pm, up from 35% in 2019",
                  "Note: share of door-count visits by arrival time, weekdays. Source: Riverside Library door counts, 2019 and 2025 (fictional data)")


def chart4(cy, chh):
    cx, cw = M, 7.4
    cd = CategoryChartData(); cd.categories = ["2019", "2025"]
    bands = [("Before noon", (30, 12)), ("Noon–4pm", (35, 26)), ("After 4pm", (35, 62))]
    for nm, v in bands:
        cd.add_series(nm, v)
    ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED_100, Inches(cx), Inches(cy), Inches(cw), Inches(chh), cd).chart
    kit.style_chart(ch)
    PX, PY, PW, PH = 0.13, 0.03, 0.6, 1 - 0.03 - (0.45 / chh)
    kit.plot_layout(ch, PX, PY, PW, PH)
    plot = ch.plots[0]; plot.gap_width = 70; plot.overlap = 100
    cols = [T["light"], MID, T["accent"]]
    for ser, c in zip(plot.series, cols):
        ser.format.fill.solid(); ser.format.fill.fore_color.rgb = c
        dl = ser.data_labels; dl.show_value = True; dl.position = XL_LABEL_POSITION.CENTER
        dl.number_format = '0"%"'; dl.number_format_is_linked = False
        dl.font.size = Pt(14); dl.font.bold = True; dl.font.color.rgb = T["ink"] if c == T["light"] else T["white"]
    va = ch.value_axis; va.has_major_gridlines = False; va.minimum_scale = 0; va.maximum_scale = 1.0
    va.tick_labels.number_format = '0%'; va.tick_labels.number_format_is_linked = False
    kit.axis(va, "Share of visits (%)"); kit.axis(ch.category_axis)
    cat_w = PW * cw / 2; bar_w = cat_w / 1.7
    leg_x = cx + PX * cw + cat_w + (cat_w - bar_w) / 2 + bar_w + 0.15
    cum = 0
    for (nm, v), c in zip(bands, cols):
        mid = cum + v[1] / 2; cum += v[1]
        ym = cy + (PY + PH) * chh - mid / 100 * PH * chh
        kit.line_label(s, leg_x, ym, nm, 16, T["accent"] if c == T["accent"] else T["ink"], bold=c == T["accent"], w=1.7)
    return cy + PY * chh, cy + (PY + PH) * chh, True


chart_with_panel(s, 7.4, "Visits by time of arrival, 2019 vs 2025", "What changed", [
    ("Evenings lead:", "after-4pm visits nearly doubled their share"),
    ("Mornings emptied:", "before-noon visits fell from 30% to 12%"),
    ("Same total:", "visits are back to 2019 levels, so this is a shift in timing, not growth")], chart4)

# ---------------------------------------------------------------- 5 purpose
s = content_slide(5, "Visitor needs", "Most evening visitors are students who come to study",
                  "Note: after-4pm visitors, n = 412; one main purpose per visit. Source: Riverside Library visitor survey, October 2025 (fictional data)")


def chart5(cy, chh):
    purposes = [("Study / quiet space", 52), ("Borrowing books", 18), ("Events and classes", 12), ("Computers and printing", 10), ("Other", 8)]
    cd = CategoryChartData(); cd.categories = [p for p, _ in reversed(purposes)]
    cd.add_series("Share", [v for _, v in reversed(purposes)])
    ch = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(M), Inches(cy), Inches(7.4), Inches(chh), cd).chart
    kit.style_chart(ch)
    kit.plot_layout(ch, 0.36, 0.02, 0.56, 1 - 0.02 - (0.75 / chh))
    plot = ch.plots[0]; plot.gap_width = 60; plot.vary_by_categories = False
    ser = plot.series[0]; ser.format.fill.solid(); ser.format.fill.fore_color.rgb = T["light"]
    pt_ = ser.points[len(purposes) - 1]; pt_.format.fill.solid(); pt_.format.fill.fore_color.rgb = T["accent"]
    dl = ser.data_labels; dl.show_value = True; dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.number_format = '0"%"'; dl.number_format_is_linked = False; dl.font.size = Pt(14)
    va = ch.value_axis; va.has_major_gridlines = False; va.minimum_scale = 0; va.maximum_scale = 60; va.major_unit = 20
    kit.axis(va, "Share of after-4pm visitors (%)"); kit.axis(ch.category_axis)
    return cy + 0.02 * chh, cy + (1 - 0.75 / chh) * chh, False


chart_with_panel(s, 7.4, "Main purpose of visit, after 4pm", "What it means", [
    ("Space, not services:", "evening staff mainly need to open floors and keep them safe"),
    ("Term-time peaks:", "the timetable should flex in exam season")], chart5)

# ---------------------------------------------------------------- 6 contents
tracker(6, 1)

# ---------------------------------------------------------------- 7 demand vs capacity
s = content_slide(7, "Staff capacity", "Moving two daytime shifts into the evening would cover the peak",
                  "Note: capacity assumes one staff member can support ~55 visitors an hour. Source: door counts by hour, average weekday, 2025; staff rota (fictional data)")


def chart7(cy, chh):
    cx, cw = M, 8.6
    PX, PY, PW, PH = 0.09, 0.04, 0.70, 1 - 0.04 - (0.75 / chh)
    X = lambda v: cx + (PX + PW * (v - 9) / 12) * cw
    Y = lambda v: cy + (PY + PH) * chh - v / 250 * PH * chh
    demand = [40, 50, 60, 70, 70, 75, 95, 140, 165, 178, 150, 110]
    cross = 15.5 + (110 - 95) / (140 - 95)
    poly = [(cross, 110), (16.5, 140), (17.5, 165), (18.5, 178), (19.5, 150), (20.5, 110)]
    fb = s.shapes.build_freeform(Inches(X(poly[0][0])), Inches(Y(poly[0][1])), scale=1.0)
    fb.add_line_segments([(Inches(X(a)), Inches(Y(b))) for a, b in poly[1:]], close=True)
    g = fb.convert_to_shape(); g.fill.solid(); g.fill.fore_color.rgb = SHADE; g.line.fill.background(); kit.no_effects(g)
    cd = XyChartData()
    sv = cd.add_series("Visitors")
    for i, d in enumerate(demand):
        sv.add_data_point(9.5 + i, d)
    for nm, lvl in (("Current capacity", 110), ("New capacity", 220)):
        se = cd.add_series(nm); se.add_data_point(9, lvl); se.add_data_point(21, lvl)
    ch = s.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS, Inches(cx), Inches(cy), Inches(cw), Inches(chh), cd).chart
    kit.style_chart(ch); kit.plot_layout(ch, PX, PY, PW, PH); kit.transparent_chart(ch)
    series = list(ch.plots[0].series)
    series[0].format.line.color.rgb = T["accent"]; series[0].format.line.width = Pt(3); series[0].smooth = False
    for ser, col in zip(series[1:], [T["grey"], T["main"]]):
        ser.format.line.color.rgb = col; ser.format.line.width = Pt(2.25); ser.format.line.dash_style = 3; ser.smooth = False
    va = ch.value_axis; va.has_major_gridlines = False; va.minimum_scale = 0; va.maximum_scale = 250; va.major_unit = 50
    kit.axis(va, "Visitors per hour")
    xa = ch.category_axis; xa.minimum_scale = 9; xa.maximum_scale = 21; xa.major_unit = 2; xa.has_major_gridlines = False
    xa.tick_labels.number_format = '0":00"'; xa.tick_labels.number_format_is_linked = False
    kit.axis(xa, "Time of day")
    kit.line_label(s, X(21) + 0.12, Y(220), "New capacity", 16, T["main"])
    kit.line_label(s, X(21) + 0.12, Y(110), "Current capacity", 16, T["grey"])
    kit.text(s, X(12.3), Y(58), 1.4, 0.35, "Visitors", 16, T["accent"], bold=True, smart=False)
    lab_y = Y(84)
    kit.text(s, X(18.6) - 1.2, lab_y, 2.4, 0.4, "Unmet demand today", 16, T["ink"], bold=True, align=PP_ALIGN.CENTER, smart=False)
    kit.two_tone_arrow(s, X(18.6), lab_y - 0.04, X(18.6), Y(110), X(18.6), Y(140), T["ink"], T["white"])
    return cy + PY * chh, cy + (PY + PH) * chh, True


chart_with_panel(s, 8.6, "Visitors per hour vs evening staff capacity, average weekday", "Today vs new", [
    ("Today:", "from 4pm, demand outruns staff by up to 68 visitors an hour"),
    ("New:", "capacity covers every hour, including the 6pm peak of 178"),
    ("Daytime:", "demand stays below capacity under both rotas, so only evenings are shown")], chart7, panel_pt=16)

# ---------------------------------------------------------------- 8 show the math
s = content_slide(8, "Cost and impact", "That change is cost-neutral and lets staff support ~20% more visitors",
                  "Note: visitors supported = visitors arriving each hour, up to that hour's staff capacity. Source: door counts and staff rota, 2025 (fictional data)")
cols8 = (("Cost", [("36", "staff-hours a day, today", "4 staff 9am–3pm + 2 staff 3pm–9pm"),
                   ("36", "staff-hours a day, new timetable", "2 staff 9am–3pm + 4 staff 3pm–9pm"),
                   ("£0", "change in staff cost", "Two shifts move later; same total hours")]),
         ("Visitors supported", [("1,010", "visitors supported a day, today", "Evening demand above capacity is unmet"),
                                 ("1,203", "visitors supported a day, new", "Every hour's demand is within capacity"),
                                 ("+19%", "more visitors supported", "≈ 190 more visitors a day")]))
NUMW = 1.8
LW = CW / 2 - 0.3 - NUMW - 0.2
LBL_PT = 18 if all(kit.width_in(r[1], 18, True) <= LW * 0.97 for _, rows in cols8 for r in rows) else 16
for _, rows in cols8:   # every pair must be one label line + one note line so rows align across columns
    for r in rows:
        assert kit.lines(r[2], 14, LW) == 1, ("note wraps", r[2])
hdr_h = max(kit.text_height(t_, 18, NUMW, True) for t_, _ in cols8)
ROW_STEP = 1.5
top = kit.place_balanced(hdr_h + 0.3 + 2 * ROW_STEP + 0.6)   # last row: label + note visible height                               # headers on the deck's content line, like the chart titles
bottom_text = 0
for j, (title, rows) in enumerate(cols8):
    x = [M, M + CW / 2 + 0.3][j]
    kit.text(s, x, top, NUMW, hdr_h + 0.05, title, 18, T["main"], bold=True, align=PP_ALIGN.CENTER)
    pairs = []
    for i, (num, label, note) in enumerate(rows):
        y = top + hdr_h + 0.3 + i * ROW_STEP
        lh = kit.text_height(label, LBL_PT, LW, True)
        nh = kit.text_height(note, 14, LW)
        kit.text(s, x + NUMW + 0.2, y, LW, lh + 0.02, label, LBL_PT, T["ink"], bold=True)
        kit.text(s, x + NUMW + 0.2, y + lh + 0.06, LW, nh + 0.02, note, 14, T["grey"])
        pb = y + lh + 0.06 + nh
        mid = (y + pb) / 2
        kit.text(s, x, mid - 0.4, NUMW, 0.8, num, 36, T["accent"] if i == 2 else T["main"], bold=True,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, smart=False)
        pairs.append((mid - 0.22, mid + 0.22))         # visible extent of the 36pt figure
        bottom_text = max(bottom_text, pb)
    kit.markers_between(s, x + NUMW / 2, [(pairs[k][1], pairs[k + 1][0]) for k in range(2)], color=T["muted_main"])
kit.line(s, M + CW / 2, top + 0.04, M + CW / 2, bottom_text, T["main"], 0.5)

# ---------------------------------------------------------------- 9 contents
tracker(9, 2)

# ---------------------------------------------------------------- 10 decision
s = content_slide(10, "Pilot plan", "Approve a six-month pilot of the new timetable from March",
                  "Note: dates assume board approval at this meeting; staff consultation runs in February (fictional plan)")
months = ["Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep"]
G10 = 4.4                                             # same group height on slides 10 and 11, so the bars line up
top = kit.place_balanced(G10)
bar_y = top + G10 - 0.85
kit.chart_header(s, M, top, CW, "Pilot plan")
tx = M + 1.0; step = (CW - 2.6) / (len(months) - 1)
ty = (bar_y - 0.63 + top + 0.45 + 1.35) / 2         # equal space: header→event labels (1.35 above the line) and months (0.63 below)→bar
trk = kit.shape(s, MSO_SHAPE.PENTAGON, tx - 0.5, ty - 0.09, (W - M) - (tx - 0.5), 0.18, T["main_pale"]); trk.adjustments[0] = 0.5
kit.shape(s, MSO_SHAPE.RECTANGLE, tx + step, ty - 0.17, step * 5, 0.34, T["accent"])
kit.text(s, tx + step, ty - 0.17, step * 5, 0.34, "Six-month pilot", 14, T["white"], bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, smart=False)
for i, m in enumerate(months):
    kit.text(s, tx + i * step - 0.6, ty + 0.28, 1.2, 0.35, m, 14, T["ink"], align=PP_ALIGN.CENTER, smart=False)
DOT, LBL_W = 0.16, 1.3
events = {0: "Staff consultation", 1: "Pilot starts", 3: "Interim review", 6: "Pilot ends; final review", 7: "Board decides"}
lab_bottom = ty - 0.75
for i, label in events.items():
    x = tx + i * step
    if i != 3:
        kit.shape(s, MSO_SHAPE.OVAL, x - DOT / 2, ty - DOT / 2, DOT, DOT, T["main"], line=T["white"], lw=1.5)
    kit.text(s, x - LBL_W / 2, lab_bottom - 0.6, LBL_W, 0.6, label, 14, T["ink"], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM)
    kit.arrow(s, x, lab_bottom + 0.05, x, (ty - 0.22) if i == 3 else (ty - DOT / 2 - 0.06))
kit.bar(s, bar_y, "Decision: Does the board approve a six-month pilot of the demand-led timetable, starting in March?")

# ---------------------------------------------------------------- 11 discuss
s = content_slide(11, "Success measures", "Agree three measures of success before the pilot starts",
                  "Note: targets are proposals for discussion; baselines from 2025 data (fictional)")
kpis = [("+15%", "Weekday visitors supported", "Baseline 1,010 a day; target ~1,160"),
        ("≥ 4.0", "Evening satisfaction score", "Out of 5; baseline 3.2 in October 2025"),
        ("£0", "Increase in staff cost", "Rota stays at 36 staff-hours a day")]
GAP = 0.6; x0 = M + 0.3; kw = (CW - 0.3 - 2 * GAP) / 3
top = kit.place_balanced(G10)
bar_y = top + G10 - 0.85
ky = top + 0.28                                      # badges overhang the tiles' top edge
tile_h = bar_y - 0.45 - ky                           # and fill down to a fixed gap above the action bar
nl = max(kit.lines(kit.wrapped(l, 16, kw - 0.7, True), 16, kw - 0.7, True) for _, l, _ in kpis)
for i, (num, label, note) in enumerate(kpis):
    x = x0 + i * (kw + GAP)
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, ky, kw, tile_h, T["main_pale"])
    kit.badge(s, x, ky, 0.56, str(i + 1), pt=16)
    off = (tile_h - 2.6) / 2                         # tile contents centred in the taller tile
    kit.text(s, x + 0.4, ky + off + 0.45, kw - 0.7, 0.8, num, 40, T["main"], bold=True, smart=False)
    kit.text(s, x + 0.4, ky + off + 1.3, kw - 0.7, 0.3 * nl, label, 16, T["ink"], bold=True)
    kit.text(s, x + 0.4, ky + off + 1.38 + 0.3 * nl, kw - 0.7, 0.5, note, 14, T["grey"])
kit.bar(s, bar_y, "For discussion: Are these the right measures and targets? What would make the board stop the pilot early?")

# ---------------------------------------------------------------- 12 closing, 13 appendix
s = kit.new_slide()
bookend(s, "Thank you", "Questions and discussion", None)
s = content_slide(13, "Methodology", "We combined door counts, a visitor survey, and the staff rota to model the new timetable",
                  "Source: Riverside Library door counts 2019 and 2025; visitor survey Oct 2025; staff rota Jan 2026 (all fictional)")
steps = [("Door counts", "Hourly visits for every weekday in 2019 and 2025, averaged by hour"),
         ("Visitor survey", "412 after-4pm visitors gave one main purpose for their visit"),
         ("Capacity model", "Each staff member supports ~55 visitors an hour (from 2025 peak-hour observations)"),
         ("Timetable test", "Compared visitors supported per hour under today's rota and the proposed one")]
row = 1.0
top = kit.place_balanced(row * len(steps) - 0.4)
for i, (t_, d) in enumerate(steps):
    y = top + i * row
    kit.badge(s, M + 0.3, y + 0.3, 0.6, str(i + 1), pt=16)
    kit.text(s, M + 0.9, y, 3.0, 0.6, t_, 18, T["ink"], bold=True, anchor=MSO_ANCHOR.MIDDLE)
    kit.text(s, M + 4.0, y, CW - 4.4, 0.6, d, 18, T["ink"], anchor=MSO_ANCHOR.MIDDLE)

print(kit.save(str(Path(__file__).with_name("library-recommendation.pptx"))))
