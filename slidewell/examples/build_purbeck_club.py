"""Example: a short case for action (storyline-examples, Example D): proposal to set up a student club.
Fictional university, club, survey, and budget. Theme: navy + leaf green, square corners, topic-label navigation."""
from _shared import *  # noqa: F401,F403

NAVY, GREEN, PALE, INK = rgb("22356F"), rgb("3FA34D"), rgb("E6EAF3"), rgb("1E2233")
kit = Kit(dict(body="Arial", main="22356F", main_pale="E6EAF3", accent="3FA34D", ink="1E2233", grey="687083",
               light="C9CFDD", tag_labels=("Demand", "Purpose", "Activities", "Fit with campus", "Running the club", "The ask")))
T, CW, M, W, H = kit.T, kit.CW, kit.T["M"], kit.T["W"], kit.T["H"]

# ------------------------------------------------------------ accuracy checks
assert 28 + 20 + 13 == 61                                   # survey responses
assert [round(100 * v / 61) for v in (28, 20, 13)] == [46, 33, 21]
assert round(100 * (28 + 20) / 61) == 79                    # "nearly 8 in 10 interested"
assert 1500 + 600 + 300 == 2400 and 1200 + 700 + 500 == 2400     # "Other" = transport 300 + reserve 200


def slide(n, tg, headline, source=None):
    s = kit.new_slide()
    kit.frame(s, n, None, headline, source)
    tag(kit, s, tg)
    return s


def bookend(title, sub, small):
    s = kit.new_slide()
    kit.shape(s, MSO_SHAPE.RECTANGLE, 0, 0, W, H, NAVY)
    kit.shape(s, MSO_SHAPE.RECTANGLE, 0, H - 0.35, W, 0.35, GREEN)          # motif: the green base band
    kit.text(s, M, 2.6, 10, 1.4, title, 40, T["white"], bold=True)
    kit.text(s, M, 4.1, 10, 0.5, sub, 20, PALE)
    if small:
        kit.text(s, M, 6.55, 10, 0.4, small, 14, kit.tint(NAVY, 0.55))


# 1 title
bookend("Proposal: a Social Impact Club at Purbeck University", "To the Students' Union executive",
        "October 2026  |  Illustrative example with fictional data")

# 2 demand
s = slide(2, "Demand", "There is strong demand for a social impact club: nearly 8 in 10 students surveyed want one",
          "Note: online survey of 61 students across all years, September 2026; peer universities of similar size (fictional data).")
ch_w, ch_h = 7.6, 3.2
top, cy = chart_slide_layout(kit, s, "Would you join a social impact club? (% of 61 students)", ch_w, ch_h, panel_title="At peer universities")
pt_, pb_, ch, _ = bar_chart(kit, s, M, cy, ch_w, ch_h, ["Yes, definitely", "Probably", "No"], [46, 33, 21], horizontal=True,
                            highlight=(0, 1), axis_title="% of respondents", maximum=60, fmt='0"%"', gap=50)
rx = M + ch_w + 0.6; pw = W - M - rx
kit.text(s, rx, pt_ + 0.05, pw, 0.9, "9 of 12", 44, NAVY, bold=True, smart=False)
kit.text(s, rx, pt_ + 0.95, pw, 0.9, "similar-sized universities already have a social impact club", 18, INK)

# 3 purpose: a two-way exchange
s = slide(3, "Purpose", "The club will put students' skills to work for local charities, and bring their experience onto campus")
cols = [("users", "Students give", ["Skills from their courses: marketing, data, design, law", "Time on short, well-scoped projects", "Fresh ideas and energy"]),
        ("heart-handshake", "Charities give", ["Real projects for students' CVs", "Talks on running a cause", "A route into social-sector careers"])]
cw2, gutter = (CW - 1.6) / 2, 1.6
top = kit.place_balanced(2.75)
for j, (ic, title, pts) in enumerate(cols):
    x = M + j * (cw2 + gutter)
    kit.icon(s, ic, x, top, 0.7, circle=NAVY if j == 0 else GREEN)
    kit.text(s, x + 0.95, top + 0.1, cw2 - 0.95, 0.5, title, 22, NAVY if j == 0 else GREEN, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    kit.bullet_list(s, x, top + 1.05, cw2, 2.2, pts, pt=18, gap_pt=10, bullet_color=NAVY if j == 0 else GREEN)
mx = M + cw2 + gutter / 2
kit.arrow(s, mx - 0.45, top + 1.45, mx + 0.45, top + 1.45, NAVY, 1.75)
kit.arrow(s, mx + 0.45, top + 2.15, mx - 0.45, top + 2.15, GREEN, 1.75)

# 4 activities: one year
s = slide(4, "Activities", "We would run volunteering projects, speaker events, and skills workshops through the year")
terms = ["Autumn term", "Spring term", "Summer term"]
acts = [("Volunteering projects", ["Match 6 teams to local charities", "Projects run, mid-point check-in", "Hand-over and showcase"]),
        ("Speaker events", ["Charity founders' panel", "Careers in the social sector", "End-of-year celebration"]),
        ("Skills workshops", ["Scoping a project", "Data for good", "Pitching for funding"])]
c0, cw3 = 2.9, (CW - 2.9) / 3
rh = 1.05
top = kit.place_balanced(0.55 + 3 * rh)
for j, t_ in enumerate(terms):
    kit.box(s, M + c0 + j * cw3 + 0.05, top, cw3 - 0.1, 0.45, NAVY)
    kit.text(s, M + c0 + j * cw3, top, cw3, 0.45, t_, 16, T["white"], bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, smart=False)
for i, (nm, cells) in enumerate(acts):
    y = top + 0.55 + i * rh
    kit.text(s, M, y, c0 - 0.2, rh - 0.1, nm, 18, NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    for j, c_ in enumerate(cells):
        kit.box(s, M + c0 + j * cw3 + 0.05, y, cw3 - 0.1, rh - 0.1, PALE)
        kit.text(s, M + c0 + j * cw3 + 0.25, y, cw3 - 0.5, rh - 0.1, c_, 16, INK, anchor=MSO_ANCHOR.MIDDLE)

# 5 fit with what exists
s = slide(5, "Fit with campus", "The club will work with the careers and volunteering teams rather than duplicate them")
rows = [("Careers service", "Advice, placements, employer events", "Social-sector speakers; projects that count as work experience"),
        ("Volunteering team", "One-off volunteering, safeguarding checks", "Longer, skills-based projects; uses the team's checks and insurance"),
        ("Students' Union societies", "Fundraising for named causes", "Joint events; refers students who want hands-on projects")]
cws = [3.1, (CW - 3.1) / 2, (CW - 3.1) / 2]
rh = 1.0
top = kit.place_balanced(0.55 + 3 * rh)
for k, h_ in enumerate(["", "What they do today", "What the club adds"]):
    if h_:
        kit.text(s, M + sum(cws[:k]), top, cws[k], 0.45, h_, 18, NAVY if k == 1 else GREEN, bold=True, smart=False)
for i, (who, today, adds) in enumerate(rows):
    y = top + 0.55 + i * rh
    if i % 2 == 0:
        kit.box(s, M - 0.1, y, CW + 0.2, rh, PALE)
    kit.text(s, M, y, cws[0] - 0.2, rh, who, 18, NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    kit.text(s, M + cws[0], y, cws[1] - 0.3, rh, today, 16, INK, anchor=MSO_ANCHOR.MIDDLE)
    kit.text(s, M + cws[0] + cws[1], y, cws[2] - 0.2, rh, adds, 16, INK, anchor=MSO_ANCHOR.MIDDLE)

# 6 running the club: budget + continuity
s = slide(6, "Running the club", "The club will run lean, on a £2,400 annual budget, and is set up to last",
          "Note: proposed first-year budget; sponsorship agreed in principle with a local employer (fictional).")
top = kit.place_balanced(3.75)                              # measured: header, two bars, and the panel beside them
kit.chart_header(s, M, top, 8.0, "First-year budget (£)")
bw = 8.0 - 1.6
for r_, (lab, parts) in enumerate([("Income", [("Union grant", 1500, NAVY), ("Sponsorship", 600, kit.tint(NAVY, 0.35)), ("Tickets", 300, kit.tint(NAVY, 0.6))]),
                                    ("Spending", [("Events", 1200, GREEN), ("Workshops", 700, kit.tint(GREEN, 0.3)), ("Other", 500, kit.tint(GREEN, 0.6))])]):
    y = top + 0.75 + r_ * 1.45
    kit.text(s, M, y, 1.4, 0.9, lab, 18, NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    x = M + 1.6
    for nm, v, col in parts:
        w_ = bw * v / 2400
        kit.shape(s, MSO_SHAPE.RECTANGLE, x, y, w_ - 0.03, 0.9, col)
        assert kit.width_in(nm, 14, True) + 0.12 < w_, nm      # label must fit inside its segment (charts.md §5b)
        tc = T["white"] if col in (NAVY, GREEN) else INK
        kit.text(s, x, y, w_ - 0.03, 0.9, [[(nm, True, tc)], [(f"{v:,}", False, tc)]], 14, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, smart=False)
        x += w_
rx = M + 8.6; pw = W - M - rx
kit.chart_header(s, rx, top, pw, "Set up to last")
kit.lead_points(s, rx, top + 0.75, pw, [("Committee of six,", "with first- and second-years on it from the start"),
                                         ("Hand-over each May,", "with a written guide and charity contacts"),
                                         ("Staff adviser", "from the volunteering team")], pt=18, lead_color=NAVY, gap=0.3)

# 7 the ask
s = slide(7, "The ask", "We ask the Union to approve the club and a £1,500 start-up grant")
steps = [("Today", "Union approves the club and grant"), ("November", "Committee elected; first charity partners signed"),
         ("January", "Six project teams start"), ("May", "Showcase and hand-over to next year's committee")]
G = 3.4
top = kit.place_balanced(G)
xs = kit.distribute_x([2.6] * 4, M, W - M)
ty = top + 0.95
kit.shape(s, MSO_SHAPE.RECTANGLE, M, ty - 0.05, CW, 0.1, PALE)
for x, (when, what) in zip(xs, steps):
    cx = x + 1.3
    kit.text(s, x, top, 2.6, 0.45, when, 20, NAVY, bold=True, align=PP_ALIGN.CENTER, smart=False)
    kit.shape(s, MSO_SHAPE.OVAL, cx - 0.12, ty - 0.12, 0.24, 0.24, GREEN, line=T["white"], lw=1.5)
    kit.text(s, x, ty + 0.35, 2.6, 0.8, what, 16, INK, align=PP_ALIGN.CENTER)
kit.bar(s, top + G - 0.85, "Decision: Does the Union approve the Social Impact Club and a £1,500 start-up grant?", fill=NAVY)

print(kit.save(out("purbeck-social-impact-club.pptx")))
