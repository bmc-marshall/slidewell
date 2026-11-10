"""Draft deck A (v2): facilitated workshop (fictional organisation, fictional data).
Palette: single colour (plum + greys). Motif: open arc from the logo. Corners: subtly rounded (circular brand).
Navigation: grey topic label above the headline + the agenda tracked small at bottom centre (introduced on slide 3)."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from slidekit import Kit, CategoryChartData, rgb
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION

kit = Kit(dict(body="Arial", main="5B2A6E", main_pale="EFE7F2", accent="5B2A6E", ink="222222", grey="6E7177",
               light="CFC8D3", muted_ink="A9A9AF", tag_labels=("x",), corners="rounded", radius=0.12, foot_top=6.5,
               page_num_top=6.94, footer_top=6.98))   # notes sit above the bottom tracker; the page number stays low, level with the tracker
T, CW, M, W, H = kit.T, kit.CW, kit.T["M"], kit.T["W"], kit.T["H"]
PLUM, PALE = T["main"], T["main_pale"]
SESSION = rgb("1F6FB2")          # share-version colour for input from the room (outside the brand palette)
STAGES = ["Attract", "Onboard", "Engage", "Retain"]
AGENDA = [("10:00", "Welcome", "Introductions and warm-up"), ("10:15", "Breakpoints", "Where does the journey break?"),
          ("10:45", "Ideas", "How could we fix it?"), ("11:10", "Pick three", "Vote on what to test"),
          ("11:25", "Next steps", "Who does what, by when")]
ARC_D = 0.38


def arc(s, x, y, d, thick=0.08, color=None, start=300, end=200):
    a = kit.shape(s, MSO_SHAPE.BLOCK_ARC, x, y, d, d, color or PLUM)
    a.adjustments[0] = start; a.adjustments[1] = end; a.adjustments[2] = thick
    return a


def tracker(s, current, on_dark=False):
    """The agenda, small, bottom centre: current item bold plum, the rest muted. Introduced on the agenda slide."""
    pt, gap = 12, 0.3
    names = [a[1] for a in AGENDA]
    slot = max(kit.width_in(n, pt, True) for n in names) + 0.1    # equal slots: nothing shifts as the bold item moves
    x = (W - (slot * len(names) + gap * (len(names) - 1))) / 2
    for i, n in enumerate(names):
        on = i == current
        col = (rgb("FFFFFF") if on else kit.tint(PLUM, 0.45)) if on_dark else (PLUM if on else T["muted_ink"])
        kit.text(s, x, 6.98, slot, 0.26, n, pt, col, bold=on, align=PP_ALIGN.CENTER, smart=False)
        if i < len(names) - 1:
            dot = (kit.tint(PLUM, 0.45) if on_dark else T["light"])
            kit.shape(s, MSO_SHAPE.OVAL, x + slot + gap / 2 - 0.03, 6.98 + 0.1, 0.06, 0.06, dot)
        x += slot + gap


def slide(n, label, headline, source=None, step=None):
    s = kit.new_slide()
    kit.frame(s, n, None, headline, source)
    kit.text(s, M, 0.26, 6, 0.32, label.upper(), 14, T["grey"], bold=True, smart=False)   # grey: distinct from the headline
    ax, ay = kit.corner_xy(ARC_D, ARC_D, "tr", 0.3)                                       # equal gap to top and right
    arc(s, ax, ay, ARC_D, thick=0.12)
    if step is not None:
        tracker(s, step)
    return s


def bookend(s, title, sub, small):
    arc(s, 7.9, 0.95, 5.6, thick=0.08)
    hs = [kit.text_height(kit.wrapped(title, 40, 7.3, True), 40, 7.3, True), 0.45, 0.4]
    y = (H - (sum(hs) + 0.4)) / 2
    kit.text(s, M, y, 7.3, hs[0] + 0.05, title, 40, T["ink"], bold=True); y += hs[0] + 0.2
    sh_ = kit.text_height(kit.wrapped(sub, 20, 6.6), 20, 6.6)
    kit.text(s, M, y, 6.6, sh_ + 0.05, sub, 20, PLUM); y += sh_ + 0.2
    if small:
        kit.text(s, M, y, 7.3, 0.4, small, 14, T["grey"])


def question_slide(label, question, prompt, step):
    """Question to the room: full-bleed brand colour so everyone knows it's their turn. No page number."""
    s = kit.new_slide()
    kit.shape(s, MSO_SHAPE.RECTANGLE, 0, 0, W, H, PLUM)
    kit.text(s, M, 0.26, 6, 0.32, label.upper(), 14, kit.tint(PLUM, 0.6), bold=True, smart=False)
    ax, ay = kit.corner_xy(ARC_D, ARC_D, "tr", 0.3)
    arc(s, ax, ay, ARC_D, thick=0.12, color=PALE)
    qh = kit.text_height(kit.wrapped(question, 36, 10.5, True), 36, 10.5, True)
    y = (H - (qh + 0.25 + 0.4)) / 2
    kit.text(s, (W - 10.5) / 2, y, 10.5, qh + 0.05, question, 36, T["white"], bold=True, align=PP_ALIGN.CENTER)
    kit.text(s, (W - 10.5) / 2, y + qh + 0.25, 10.5, 0.4, prompt, 18, PALE, align=PP_ALIGN.CENTER)
    tracker(s, step, on_dark=True)


def column_pair(s, x, y, w, head_h, body_h, title, head_fill=None, body_fill=None):
    """Header box rounded on top + content box rounded at the bottom, meeting on a straight edge."""
    kit.box(s, x, y, w, head_h, head_fill or PLUM, corners="top")
    kit.box(s, x, y + head_h, w, body_h, body_fill or PALE, corners="bottom")
    kit.text(s, x, y, w, head_h, title, 16, T["white"], bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, smart=False)


def corner_cell(s, x, y, w, h, fill, corner):
    """Grid cell with only its outer corner rounded, so the 2x2 grid reads as one rounded shape."""
    shp = kit.shape(s, MSO_SHAPE.ROUND_1_RECTANGLE, x, y, w, h, fill)    # rounds the top-right corner by default
    shp.adjustments[0] = min(0.5, T["radius"] / min(w, h))
    xfrm = shp._element.spPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm")
    if corner in ("tl", "bl"):
        xfrm.set("flipH", "1")
    if corner in ("bl", "br"):
        xfrm.set("flipV", "1")
    return shp


# ---------------------------------------------------------------- accuracy checks for every derived figure
mins = lambda t: int(t[:2]) * 60 + int(t[3:])
assert mins("11:25") - mins("10:45") == 40                     # "Main focus: 10:45–11:25, 40 minutes" (Ideas + Pick three)
assert mins("11:25") - mins("10:45") > mins("10:45") - mins("10:15")   # "the most time": 40 > 30 on breakpoints
assert mins("11:30") - mins("10:00") == 90                     # "90 minutes" (Next steps runs 11:25–11:30)
assert round(1 - 180 / 400, 2) == 0.55                         # "55% of enquirers never start"
assert abs((180 - 120) / 180 - 1 / 3) < 0.01                   # "a third of starters stop within three months"
assert round((120 - 70) / 120, 1) == 0.4                       # "four in ten of those leave within the year"

# ---------------------------------------------------------------- 1 title
s = kit.new_slide()
bookend(s, "Designing our 2027 volunteer programme", "Workshop with the volunteer leads  |  90 minutes",
        "Harbourside Community Health  |  Illustrative example with fictional content")

# ---------------------------------------------------------------- 2 objectives
s = slide(2, "Objectives", "Today we'll agree where our volunteer journey breaks and pick three fixes to test")
objs = [("Understand", "where we lose volunteers, using this year's numbers"),
        ("Agree", "the two biggest breakpoints in each stage of the journey"),
        ("Choose", "three fixes to test before the spring recruitment drive")]
row = 1.0
block = 0.45 + 0.35 + 2 * row + 0.6
top = kit.place_balanced(block)
kit.chart_header(s, M, top, CW, "By the end of this session, we will:")
for i, (verb, rest) in enumerate(objs):
    y = top + 0.8 + i * row
    kit.badge(s, M + 0.3, y + 0.3, 0.6, str(i + 1), pt=16)
    kit.text(s, M + 0.9, y, CW - 1.5, 0.6, [[(verb + " ", True, PLUM), (rest, False, T["ink"])]], 20, anchor=MSO_ANCHOR.MIDDLE)

# ---------------------------------------------------------------- 3 agenda: what the tracker will follow
s = slide(3, "Agenda", "We'll spend the most time on ideas and choosing what to test")
COLW = 2.2
xs = kit.distribute_x([COLW] * len(AGENDA))                 # equal gaps between stops and to the slide edges
block = 0.2 + 0.4 + 0.3 + 0.36 + 0.3 + 0.4 + 0.65 + 0.15 + 0.4 + 0.15
top = kit.place_balanced(block)
by = top                                                     # focus box top
ty_time = top + 0.2
ty = ty_time + 0.4 + 0.3 + 0.18                              # timeline centre
trk = kit.shape(s, MSO_SHAPE.PENTAGON, M, ty - 0.07, CW, 0.14, PALE); trk.adjustments[0] = 0.5
fx0, fx1 = xs[2] - 0.12, xs[3] + COLW + 0.12
kit.box(s, fx0, by, fx1 - fx0, block, None, line=kit.tint(PLUM, 0.45), lw=1.5)   # focus: pale outline, not bigger dots
for i, (t_, name, desc) in enumerate(AGENDA):
    x = xs[i]; cx = x + COLW / 2
    focus = i in (2, 3)
    kit.text(s, x, ty_time, COLW, 0.4, t_, 18, PLUM, bold=focus, align=PP_ALIGN.CENTER, smart=False)
    kit.badge(s, cx, ty, 0.36, "", on=True)
    kit.text(s, x, ty + 0.48, COLW, 0.4, name, 18, T["ink"], bold=True, align=PP_ALIGN.CENTER, smart=False)
    kit.text(s, x + 0.1, ty + 0.9, COLW - 0.2, 0.65, desc, 14, T["grey"], align=PP_ALIGN.CENTER)
kit.text(s, fx0, by + block - 0.55, fx1 - fx0, 0.4, "Main focus: 10:45–11:25, 40 minutes", 16, PLUM, bold=True, align=PP_ALIGN.CENTER, smart=False)

# ---------------------------------------------------------------- 4 norms
s = slide(4, "How we'll work", "A few ways of working will help us get the most from 90 minutes", step=0)
items = [("Be present:", "laptops closed unless you're the note-taker"),
         ("Build on ideas", "(\"yes, and…\") before critiquing them"),
         ("Share", "what you've seen first-hand; numbers come later"),
         ("Park anything off-topic", "on the flipchart, and we'll return to it")]   # key words bold in plum
bh = sum(kit.text_height(a_ + " " + b_, 20, 6.8, True) for a_, b_ in items) + 3 * 0.25
y0 = kit.place_balanced(bh - 0.15)                           # measure visible text: drop the last line's leading
kit.bullet_list(s, M, y0, 7.1, bh + 0.2, items, pt=20, gap_pt=18)
qx, qw = M + 8.0, CW - 8.0
kit.box(s, qx, y0 + bh / 2 - 0.85, qw, 1.7, PALE)
kit.text(s, qx + 0.35, y0 + bh / 2 - 0.85, qw - 0.7, 1.7, "What else will help us make the most of our time together?", 20, PLUM,
         bold=True, anchor=MSO_ANCHOR.MIDDLE)

# ---------------------------------------------------------------- 5 question to the room
question_slide("Warm-up", "Think of a volunteer who stayed for years. What kept them?",
               "Two minutes with your neighbour, then we'll hear three stories", step=0)

# ---------------------------------------------------------------- 6 teach: funnel chart
s = slide(6, "The volunteer journey", "We lose more than half of interested volunteers before they even start",
          "Note: people who enquired between January and December 2025. Source: volunteer database (fictional data)", step=1)
CH_W, CH_H = 7.6, 3.5                                        # same chart group on every chart slide
top = kit.place_balanced(0.6 + CH_H)
stages = [("Enquired (Attract)", 400), ("Started (Onboard)", 180), ("Active at 3 months (Engage)", 120), ("Active at 12 months (Retain)", 70)]
cd = CategoryChartData(); cd.categories = [n for n, _ in reversed(stages)]
cd.add_series("People", [v for _, v in reversed(stages)])
ch_top = top + 0.6
ch = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(M), Inches(ch_top), Inches(CH_W), Inches(CH_H), cd).chart
kit.style_chart(ch)
PY, PH = 0.02, 1 - 0.02 - (0.75 / CH_H)
kit.plot_layout(ch, 0.4, PY, 0.5, PH)
plot = ch.plots[0]; plot.gap_width = 55; plot.vary_by_categories = False
ser = plot.series[0]; ser.format.fill.solid(); ser.format.fill.fore_color.rgb = T["light"]
pt_ = ser.points[2]; pt_.format.fill.solid(); pt_.format.fill.fore_color.rgb = PLUM
dl = ser.data_labels; dl.show_value = True; dl.position = XL_LABEL_POSITION.OUTSIDE_END; dl.font.size = Pt(14)
va = ch.value_axis; va.has_major_gridlines = False; va.minimum_scale = 0; va.maximum_scale = 450; va.major_unit = 150
kit.axis(va, "People", pt=12); kit.axis(ch.category_axis, pt=16)   # metric axis smaller than the labels naming the data
kit.chart_header(s, M, top, CH_W, "Volunteers remaining at each stage, 2025 intake")
rx = M + 8.2; pw = W - M - rx
kit.chart_header(s, rx, top, pw, "Where we lose people")
kit.side_points(s, rx, pw, ch_top + PY * CH_H, ch_top + (PY + PH) * CH_H, [
    ("Onboard:", "55% of enquirers never start, our biggest drop"),
    ("Engage:", "a third of starters stop within three months"),
    ("Retain:", "four in ten of those leave within the year")], lead_color=PLUM)

# ---------------------------------------------------------------- 7 exercise 1
s = slide(7, "Exercise 1", "Where does the journey break for our volunteers?", step=1)
G_H = 4.1
top = kit.place_balanced(G_H)
specs = [("Task", "List the moments where volunteers drift away, one per sticky note"),
         ("Groups", "Four tables, one stage each"),
         ("Time", "20 minutes, then two minutes per table to report back"),
         ("Output", "Your top two breakpoints, starred")]
rh = G_H / 4
for i, (k_, v) in enumerate(specs):
    y = top + i * rh
    kit.text(s, M, y, 1.4, rh, k_.upper(), 14, PLUM, bold=True, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    kit.text(s, M + 1.5, y, 4.6, rh, v, 18, T["ink"], anchor=MSO_ANCHOR.MIDDLE)
    if i:
        kit.line(s, M, y, M + 6.1, y, PALE, 1)
gx = M + 6.6; gw = (W - M - gx - 3 * 0.2) / 4
for i, st in enumerate(STAGES):
    column_pair(s, gx + i * (gw + 0.2), top, gw, 0.5, G_H - 0.5, st)

# ---------------------------------------------------------------- 8 capture (share version)
s = slide(8, "What we heard", "The room's breakpoints cluster around the first few weeks",
          "Note: blue text is input from the session on 12 March (fictional). Bold items were voted top two by each table.", step=1)
heard = {"Attract": [("Roles described vaguely online", False), ("Enquiries unanswered for 2+ weeks", True)],
         "Onboard": [("Background checks take 6–8 weeks", True), ("Induction only offered monthly", True), ("Paperwork asked for twice", False)],
         "Engage": [("First shift with no named buddy", True), ("Rotas change at short notice", False)],
         "Retain": [("No thank-you or recognition", True), ("Little chance to try new roles", False)]}
G_H = 4.1                                                    # same column height as the exercise slide it reports on
top = kit.place_balanced(G_H)
gw = (CW - 3 * 0.25) / 4
for i, st in enumerate(STAGES):
    x = M + i * (gw + 0.25)
    column_pair(s, x, top, gw, 0.5, G_H - 0.5, st)
    kit.bullet_list(s, x + 0.2, top + 0.75, gw - 0.4, G_H - 0.9, [t_ for t_, _ in heard[st]], pt=16, color=SESSION,
                    bullet_color=SESSION, gap_pt=10, bold=[b_ for _, b_ in heard[st]])

# ---------------------------------------------------------------- 9 trade-off
s = slide(9, "A tension to resolve", "Faster onboarding must be balanced against our safeguarding duties", step=2)
cols9 = (("Speed", ["Volunteers lose interest during a 6–8 week wait", "Most first roles involve no unsupervised contact",
                    "Peer charities start people within two weeks"]),
         ("Safeguarding", ["Checks are a legal duty for some roles", "One incident could undermine trust in the whole service",
                           "Staff capacity to supervise new starters is limited"]))
gutter = 1.4
colw = (CW - gutter) / 2
list_h = max(sum(kit.text_height(kit.wrapped(p_, 16, colw - 0.35), 16, colw - 0.35) for p_ in pts) + 2 * 8 / 72 for _, pts in cols9)
col_h = 0.6 + list_h
q_h = 1.35
block = col_h + 0.4 + q_h
top = kit.place_balanced(block)
for j, (title, pts) in enumerate(cols9):
    x = M + j * (colw + gutter)
    kit.text(s, x, top, colw, 0.45, title, 22, PLUM, bold=True, align=PP_ALIGN.CENTER, smart=False)
    kit.bullet_list(s, x, top + 0.6, colw, list_h + 0.1, pts, pt=16, gap_pt=8)
mid_y = top + col_h / 2                                      # "or" centred on the two columns, not on their headers
kit.line(s, W / 2, top, W / 2, top + col_h, kit.tint(PLUM, 0.6), 0.75)
kit.connector_word(s, W / 2, mid_y, "or", d=0.62, pt=16)
qy = top + col_h + 0.4
kit.box(s, M, qy, CW, q_h, PALE)
kit.text(s, M + 0.35, qy + 0.2, 4.0, 0.4, "Questions to consider", 18, PLUM, bold=True)
kit.bullet_list(s, M + 0.35, qy + 0.65, CW - 0.7, q_h - 0.75,
                ["Which roles could start before checks complete, with supervision?",
                 "What would it take to offer induction every week instead of every month?"], pt=18, gap_pt=6)

# ---------------------------------------------------------------- 10 exercise 2: impact / effort
s = slide(10, "Exercise 2", "We'll pick three fixes to test, starting with high impact and low effort", step=3)
gs = 3.0; gw = 5.6
block = gs + 0.7                                             # grid + effort axis and label
top = kit.place_balanced(block)
gx = M + 1.25
cells = [("Quick wins", "Test these first", PLUM, T["white"], "tl"), ("Big bets", "Plan carefully", kit.tint(PLUM, 0.55), T["ink"], "tr"),
         ("Maybe later", "Only if easy", kit.tint(PLUM, 0.85), T["ink"], "bl"), ("Avoid", "Not now", kit.tint(PLUM, 0.92), T["grey"], "br")]
for k_, (name, sub, fill, tc, corner) in enumerate(cells):
    cx_, cy_ = gx + (k_ % 2) * gw / 2, top + (k_ // 2) * gs / 2
    corner_cell(s, cx_ + 0.03, cy_ + 0.03, gw / 2 - 0.06, gs / 2 - 0.06, fill, corner)
    kit.text(s, cx_, cy_, gw / 2, gs / 2, [[(name + "\v", True, tc), (sub, False, tc)]], 18, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
kit.arrow(s, gx - 0.25, top + gs, gx - 0.25, top, T["grey"], 1.25)
kit.text(s, M, top + gs / 2 - 0.2, gx - 0.4 - M, 0.4, "Impact", 14, T["grey"], bold=True, align=PP_ALIGN.RIGHT, smart=False)
kit.arrow(s, gx, top + gs + 0.25, gx + gw, top + gs + 0.25, T["grey"], 1.25)
kit.text(s, gx + gw / 2 - 1, top + gs + 0.33, 2, 0.35, "Effort", 14, T["grey"], bold=True, align=PP_ALIGN.CENTER, smart=False)
rx = gx + gw + 0.8; pw = W - M - rx
kit.chart_header(s, rx, top, pw, "How to vote")
steps10 = ["Each table places its ideas on the grid (10 minutes)",
           "Everyone gets three dots for the ideas they'd test first",
           "The three ideas with the most dots in Quick wins go forward"]
sh10 = [kit.text_height(kit.wrapped(t_, 18, pw - 0.35), 18, pw - 0.35) for t_ in steps10]
avail = gs - 0.6                                             # panel spans the grid: from below its header to the grid's bottom
gap10 = min(2.5 * 0.3, max(0.3, (avail - sum(sh10)) / 2))
kit.bullet_list(s, rx, top + 0.6, pw, avail + 0.1, steps10, pt=18, gap_pt=gap10 * 72)

# ---------------------------------------------------------------- 11 objectives revisited + next steps
s = slide(11, "Wrap-up", "We've agreed the breakpoints and three fixes; here's what's next", step=4)
done = [("Understand", "We lose most people between enquiring and starting"),
        ("Agree", "Breakpoints cluster in the first few weeks"),
        ("Choose", "Weekly induction, a named buddy, and a two-week reply promise")]
nexts = [("Week 1", "Leads share the three tests with their teams"), ("Week 4", "Tests running at two sites"),
         ("Week 8", "Review results; decide what to scale")]
row, rh = 1.1, 0.75
block = 0.6 + 2 * row + rh
top = kit.place_balanced(block)
lw_ = 6.0; nx = M + 6.8; nw = W - M - nx
kit.chart_header(s, M, top, lw_, "What we set out to do")
kit.chart_header(s, nx, top, nw, "Next steps")
for i in range(3):
    y = top + 0.6 + i * row                                  # rows shared by both columns
    verb, rest = done[i]
    kit.badge(s, M + 0.27, y + rh / 2, 0.5, str(i + 1), pt=14)   # the objective's number from slide 2
    kit.text(s, M + 0.75, y, lw_ - 0.75, rh, [[(verb + ": ", True, PLUM), (rest, False, T["ink"])]], 16, anchor=MSO_ANCHOR.MIDDLE)
    when, what = nexts[i]
    kit.box(s, nx, y, 1.2, rh, PLUM)
    kit.text(s, nx, y, 1.2, rh, when, 16, T["white"], bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    kit.text(s, nx + 1.4, y, nw - 1.4, rh, what, 16, T["ink"], anchor=MSO_ANCHOR.MIDDLE)

# ---------------------------------------------------------------- 12 close
s = kit.new_slide()
bookend(s, "Thank you", "Please leave one sticky note: what should we keep doing?", None)

print(kit.save(str(Path(__file__).with_name("volunteer-workshop.pptx"))))
