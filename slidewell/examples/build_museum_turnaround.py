"""Example: a findings-and-recommendations deck with three MECE sections and a section tracker
(storyline-examples §3, the museum turnaround). Fictional museum, town, benchmarks, survey, and figures.
Theme: deep navy + ochre accent, square corners, recurring hero image on dividers, small tracker top-centre."""
from _shared import *  # noqa: F401,F403

NAVY, OCHRE, PALE, INK = rgb("1D3557"), rgb("D98E04"), rgb("EDF1F6"), rgb("1F2430")
kit = Kit(dict(body="Arial", main="1D3557", main_pale="EDF1F6", accent="D98E04", ink="1F2430", grey="667080",
               light="C7CED9", tag_labels=("x",), head_top=0.72))
T, CW, M, W, H = kit.T, kit.CW, kit.T["M"], kit.T["W"], kit.T["H"]
SECTIONS = ["Overview", "Reasons to visit", "Audiences", "Value"]
HERO = "Photo: the museum's main hall, visitors looking up at the whale skeleton"

# ------------------------------------------------------------ accuracy checks for every derived figure
visits = [82, 79, 41, 58, 63, 60, 57, 55]                      # thousands, 2018-2025 (2020: closures)
assert round(1 - 55 / 82, 2) == 0.33                           # "down a third since 2018"
assert round(1 - 6 / 9, 2) == 0.33                             # "£6 ticket is a third below the £9 peer median"
assert sorted([8, 9, 10, 12, 9])[2] == 9                       # peer median ticket price
assert 42 - 32 == 10                                           # "10 fewer hours a week"
assert round(65 / 55 - 1, 2) == 0.18                           # visits target +18%
assert round(8.5 / 7.1 - 1, 2) == 0.20                         # revenue per visit +20%


def tracker(s, current):
    """Small section tracker, top-centre; introduced on the executive summary. Fixed-width slots."""
    pt, gap = 12, 0.35
    slot = max(kit.width_in(n, pt, True) for n in SECTIONS) + 0.1
    x = (W - (slot * len(SECTIONS) + gap * (len(SECTIONS) - 1))) / 2
    for i, n in enumerate(SECTIONS):
        on = i == current
        kit.text(s, x, 0.2, slot, 0.28, n, pt, NAVY if on else T["muted_ink"], bold=on, align=PP_ALIGN.CENTER, smart=False)
        if i < len(SECTIONS) - 1:
            kit.shape(s, MSO_SHAPE.OVAL, x + slot + gap / 2 - 0.03, 0.31, 0.06, 0.06, T["light"])
        x += slot + gap
    kit.line(s, M, 0.56, W - M, 0.56, T["light"], 0.5)


def slide(n, sec, headline, source=None):
    s = kit.new_slide()
    kit.frame(s, n, None, headline, source)
    tracker(s, sec)
    return s


def divider(num, title, items):
    """Section divider: the recurring hero image (same photo, same crop), with a gradient overlay for the text."""
    s = kit.new_slide()
    kit.photo(s, 0, 0, W, H, label=HERO)
    kit.overlay(s, 0, 0, W * 0.62, H, NAVY, mode="gradient", opacity=0.92, stops=[(0, 0.92), (0.7, 0.85), (1, 0)])
    kit.text(s, M, 2.2, 3, 1.0, f"{num:02d}", 60, OCHRE, bold=True, smart=False)
    kit.text(s, M, 3.3, 7, 1.0, title, 36, T["white"], bold=True)
    kit.text(s, M, 4.4, 7, 0.5, "  •  ".join(items), 18, PALE, smart=False)
    return s


# 1 title
s = kit.new_slide()
kit.photo(s, 0, 0, W, H, label=HERO)
kit.overlay(s, 0, 0, W * 0.62, H, NAVY, mode="gradient", opacity=0.92, stops=[(0, 0.92), (0.7, 0.85), (1, 0)])
kit.text(s, M, 2.3, 7.2, 1.4, "Bringing visitors back to Kestrel Bay Museum", 40, T["white"], bold=True)
kit.text(s, M, 3.95, 7.2, 0.5, "Findings and recommendations for the board of trustees", 20, PALE)
kit.text(s, M, 6.7, 7.2, 0.35, "March 2027  |  Illustrative example with fictional data", 14, kit.tint(NAVY, 0.55))

# 2 executive summary: labels are mini findings; the three recommendation rows introduce the tracker sections
s = slide(2, 0, "Three moves would bring visitors back and raise income per visit")
rows = [("Visits are down a third", "Kestrel Bay has lost a third of its visitors since 2018, and income has followed", False),
        ("More reasons to visit", "Change exhibitions more often and fix the front of house first", True),
        ("Reach local audiences", "Partner locally, regionally, and nationally so people hear what's on", True),
        ("Capture more value", "Raise prices towards peers with a membership, and open later", True)]
rh, gap = 0.95, 0.2
top = kit.place_balanced(4 * rh + 3 * gap)
for i, (lab, txt, rec) in enumerate(rows):
    y = top + i * (rh + gap)
    kit.shape(s, MSO_SHAPE.RECTANGLE, M, y, 3.6, rh, NAVY if rec else T["light"])
    kit.text(s, M + 0.25, y, 3.2, rh, lab, 18, T["white"] if rec else NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    kit.text(s, M + 3.9, y, CW - 3.9, rh, txt, 20, INK, anchor=MSO_ANCHOR.MIDDLE)

# 3 the problem
s = slide(3, 0, "Kestrel Bay has lost a third of its visitors since 2018",
          "Note: 2020 and 2021 affected by closures. Source: museum ticketing system, 2018–2025 (fictional data)")
CH_W, CH_H = 7.9, 3.4
top, cy = chart_slide_layout(kit, s, "Visits a year, 2018–2025 (thousands)", CH_W, CH_H, "What changed")
pt_, pb, ch, (PX, PY, PW, PH) = bar_chart(kit, s, M, cy, CH_W, CH_H, [str(y) for y in range(2018, 2026)], visits,
                                          highlight=(0, 7), axis_title="Thousand visits", maximum=100)
rx = M + CH_W + 0.5
kit.side_points(s, rx, W - M - rx, pt_, pb, [("Fewer repeat visits:", "local residents now come once every three years"),
                                              ("Same exhibitions:", "the main galleries last changed in 2015"),
                                              ("Income:", "ticket and café income fell with visits")], lead_color=NAVY)

# 4 what we did (a method slide: the headline still states what the work found its way to)
s = slide(4, 0, "We compared 40 museums and surveyed 480 residents to find what brings people back",
          "Note: comparison museums are independent museums of similar size in the UK; survey run online and at the library, January 2027 (fictional).")
steps = [("search", "Compare", "40 similar museums: exhibitions, prices, hours, partnerships"),
         ("users", "Ask", "480 local residents: what they know, want, and would pay"),
         ("bar-chart-3", "Analyse", "Visits and income by month, ticket type, and hour"),
         ("lightbulb", "Recommend", "Three groups of changes, quick wins first")]
cw4 = 2.6
xs = kit.distribute_x([cw4] * 4, M, W - M)
desc_h = max(kit.text_height(kit.wrapped(d, 16, cw4 - 0.2), 16, cw4 - 0.2) for _, _, d in steps)
top = kit.place_balanced(1.5 + desc_h)                      # measured: icon, label, then the tallest description
for i, (ic, t_, d) in enumerate(steps):
    x = xs[i]
    kit.icon(s, ic, x + cw4 / 2 - 0.4, top, 0.8, circle=NAVY)
    kit.text(s, x, top + 1.0, cw4, 0.45, t_, 20, NAVY, bold=True, align=PP_ALIGN.CENTER, smart=False)
    kit.text(s, x, top + 1.5, cw4, 1.3, d, 16, INK, align=PP_ALIGN.CENTER)
    if i < 3:
        g0 = x + cw4; g1 = xs[i + 1]
        mk = kit.shape(s, MSO_SHAPE.ISOSCELES_TRIANGLE, (g0 + g1) / 2 - 0.08, top + 0.4 - 0.1, 0.16, 0.2, T["light"]); mk.rotation = 90

# 5 divider 1
divider(1, "Creating reasons to visit", ["Exhibitions", "Front of house", "Galleries"])

# 6 exhibitions benchmark
s = slide(6, 1, "Growing museums change exhibitions three times as often as we do",
          "Note: average number of new temporary exhibitions a year, 2022–2025, by change in visits over the period. Source: museum annual reports (fictional data)")
top, cy = chart_slide_layout(kit, s, "New exhibitions a year, by visitor trend", CH_W, CH_H, "What growing museums do")
pt_, pb, ch, geo = bar_chart(kit, s, M, cy, CH_W, CH_H, ["Growing museums", "Stable museums", "Declining museums", "Kestrel Bay"],
                             [6.1, 4.0, 2.3, 2.0], highlight=(0,), horizontal=True, axis_title="Exhibitions a year", maximum=8,
                             fmt="0.0", pt_cat=16, colors=[NAVY, None, None, OCHRE])
assert round(6.1 / 2.0) == 3                                   # "three times as often"
kit.side_points(s, rx, W - M - rx, pt_, pb, [("Smaller shows,", "more often: six a year, most under six weeks"),
                                              ("Local stories:", "objects lent by residents and schools"),
                                              ("Events:", "late openings built around each show")], lead_color=NAVY)

# 7 quick vs longer-term changes
s = slide(7, 1, "Quick front-of-house fixes can start now, ahead of gallery changes")
cols = [("Now: front of house", "0–6 months", ["Clear signs and a 'what's on' board outside", "Café open to non-visitors", "Free entry for under-16s with an adult"]),
        ("Next: the galleries", "6–24 months", ["Rotate one gallery every two months", "A family trail through every room", "Seating and better lighting in the main hall"])]
cw7 = (CW - 0.8) / 2
items_h = max(sum(kit.text_height(kit.wrapped(p_, 18, cw7 - 0.8), 18, cw7 - 0.8) for p_ in pts) + 2 * 14 / 72 for _, _, pts in cols)
top = kit.place_balanced(0.55 + 0.45 + items_h)
for j, (t_, when, pts) in enumerate(cols):
    x = M + j * (cw7 + 0.8)
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, top, cw7, 0.55, NAVY if j == 0 else kit.tint(NAVY, 0.4))
    kit.text(s, x + 0.25, top, cw7 - 0.5, 0.55, t_, 18, T["white"], bold=True, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    kit.text(s, x + 0.25, top, cw7 - 0.5, 0.55, when, 16, T["white"], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    kit.bullet_list(s, x + 0.25, top + 0.85, cw7 - 0.5, items_h, pts, pt=18, gap_pt=14)

# 8 divider 2
divider(2, "Connecting with audiences", ["Local", "Regional", "National"])

# 9 partnerships: local -> regional -> national
s = slide(9, 2, "Partnerships at three levels would help people hear what's on and give them reasons to come")
lv = [("map-pin", "Locally", "Schools, the library, and walking groups", ["Term-time school visits with a teacher pack", "A shared events calendar with the library"]),
      ("map", "Regionally", "Other museums and the tourist board", ["A joint pass with two nearby museums", "Listings in the coast's visitor guide"]),
      ("globe", "Nationally", "Museum networks and lenders", ["Touring exhibitions from national collections", "Membership of the national museums association"])]
cw9 = 3.6
xs = kit.distribute_x([cw9] * 3, M, W - M)
bul_h = max(sum(kit.text_height(kit.wrapped(p_, 16, cw9 - 0.3), 16, cw9 - 0.3) for p_ in pts) + 10 / 72 for *_, pts in lv)
top = kit.place_balanced(1.65 + bul_h)
for i, (ic, t_, who, pts) in enumerate(lv):
    x = xs[i]
    kit.icon(s, ic, x, top, 0.7, circle=NAVY)
    kit.text(s, x + 0.9, top, cw9 - 0.9, 0.7, t_, 20, NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    kit.text(s, x, top + 0.9, cw9, 0.6, who, 16, T["grey"])
    kit.bullet_list(s, x, top + 1.65, cw9, 1.8, pts, pt=16, gap_pt=10)

# 10 divider 3
divider(3, "Capturing visitor value", ["Pricing", "Membership", "Opening hours"])

# 11 pricing
s = slide(11, 3, "Our £6 ticket is a third below peers, so it can rise with a membership",
          "Note: standard adult ticket, 2026. Peers: five museums of similar size within 50 miles (fictional data)")
top, cy = chart_slide_layout(kit, s, "Adult ticket price (£)", CH_W, CH_H, "How to raise it")
pt_, pb, ch, geo = bar_chart(kit, s, M, cy, CH_W, CH_H, ["Peer A", "Peer B", "Peer C", "Peer D", "Peer E", "Kestrel Bay"],
                             [8, 9, 10, 12, 9, 6], axis_title="£ per adult ticket", maximum=14, colors=[None] * 5 + [OCHRE])
PX, PY, PW, PH = geo
Y = lambda v: cy + (PY + PH) * CH_H - v / 14 * PH * CH_H
kit.line(s, M + PX * CH_W, Y(9), M + (PX + PW) * CH_W, Y(9), T["grey"], 1.0, dash=4)
kit.text(s, M + (PX + PW) * CH_W - 2.1, Y(9) - 0.34, 2.05, 0.3, "Peer median £9", 14, INK, align=PP_ALIGN.RIGHT, smart=False)
kit.side_points(s, rx, W - M - rx, pt_, pb, [("Step up:", "£6 to £8 in April, with under-16s free"),
                                              ("Membership:", "£30 a year for unlimited visits and late openings"),
                                              ("Residents:", "seven in ten would pay £8 with free return visits")], lead_color=NAVY)

# 12 opening hours
s = slide(12, 3, "We open 10 fewer hours a week than peers, and close before 5pm",
          "Note: standard weekly opening hours outside school holidays, 2026. Peer median of the same five museums (fictional data)")
top, cy = chart_slide_layout(kit, s, "Opening hours a week", CH_W, CH_H, "What we recommend")
pt_, pb, ch, geo = bar_chart(kit, s, M, cy, CH_W, CH_H, ["Peer median", "Kestrel Bay"], [42, 32], horizontal=True,
                             axis_title="Hours a week", maximum=50, pt_cat=16, colors=[NAVY, OCHRE])
kit.side_points(s, rx, W - M - rx, pt_, pb, [("Late Thursdays:", "open until 8pm, with an event each week"),
                                              ("Weekends:", "open at 10am, not 11am")], lead_color=NAVY)

# 13 success in 12 months
s = slide(13, 3, "In 12 months, success would mean more visits, 800 members, and more income per visit")
kpis = [("65,000", "visits a year", "up 18% from 55,000"), ("800", "members", "from none today"), ("£8.50", "income per visit", "up 20% from £7.10")]
tw = 3.5
xs = kit.distribute_x([tw] * 3, M, W - M)
top = kit.place_balanced(2.6)
for i, (num, lab, note) in enumerate(kpis):
    x = xs[i]
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, top, tw, 2.6, PALE)
    kit.text(s, x, top + 0.35, tw, 0.8, num, 40, NAVY, bold=True, align=PP_ALIGN.CENTER, smart=False)
    kit.text(s, x, top + 1.25, tw, 0.45, lab, 20, INK, bold=True, align=PP_ALIGN.CENTER, smart=False)
    kit.text(s, x, top + 1.75, tw, 0.45, note, 16, T["grey"], align=PP_ALIGN.CENTER, smart=False)

# 14 appendix sample
s = slide(14, 0, "Appendix: we chose comparison museums by size, setting, funding",
          "Source: museum association directory, 2026 (fictional)")
crit = [("Size", "20,000–150,000 visits a year"), ("Setting", "Coastal or small-town, outside major cities"),
        ("Funding", "Independent trust, not local-authority run"), ("Result", "40 museums; 5 nearest used for prices and hours")]
rh = 0.8
top = kit.place_balanced(4 * rh)
for i, (k_, v) in enumerate(crit):
    y = top + i * rh
    if i % 2 == 1:                                          # stripe rows 2 and 4, so the table's last row has an edge
        kit.shape(s, MSO_SHAPE.RECTANGLE, M, y, CW, rh, PALE)
    kit.text(s, M + 0.25, y, 2.6, rh, k_, 18, NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    kit.text(s, M + 3.0, y, CW - 3.2, rh, v, 18, INK, anchor=MSO_ANCHOR.MIDDLE)

print(kit.save(out("museum-turnaround.pptx")))
