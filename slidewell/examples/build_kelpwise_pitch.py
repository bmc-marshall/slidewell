"""Example: investment pitch with a strong slide-to-slide flow (storyline-examples, Example A).
Fictional company, people, partners, and figures. Theme: deep sea teal + coral accent, wave-edge motif."""
from _shared import *  # noqa: F401,F403

TEAL, CORAL, PALE, INK = rgb("0E4D64"), rgb("E4572E"), rgb("E3EEF1"), rgb("1F2A30")
kit = Kit(dict(body="Arial", main="0E4D64", main_pale="E3EEF1", accent="E4572E", ink="1F2A30", grey="6A757B",
               light="C8D3D8", tag_labels=("Recommendation", "The problem", "The solution", "The company", "The investment", "Risks")))
T, CW, M, W, H = kit.T, kit.CW, kit.T["M"], kit.T["W"], kit.T["H"]

# ------------------------------------------------------------ accuracy checks for every derived figure
assert 92 > 90                                         # ">90,000 tons" (2025 value, thousands)
assert abs(39 - 38) <= 1                               # seaweed tray within €1 of plastic
assert 800 + 400 + 300 == 1500                         # $1.5M seed round, of which our $300K
assert round(300 / 1500, 2) == 0.2                     # "20% of the round"
assert round(0.9 + 0.5 + 0.4, 6) == 1.8                # €1.8M contracted revenue (round: decimals in floats)


def wave(s, x, y, h, color, flip=False):
    """Motif: a gentle wave edge, used on the title, closing, and the section-marking tag."""
    fb = s.shapes.build_freeform(Inches(x), Inches(y), scale=1.0)
    pts, n = [], 24
    import math
    for i in range(n + 1):
        yy = y + h * i / n
        xx = x + 0.3 * math.sin(i / n * math.pi * 3)
        pts.append((Inches(xx), Inches(yy)))
    pts += [(Inches(x + (3 if not flip else -3)), Inches(y + h)), (Inches(x + (3 if not flip else -3)), Inches(y))]
    fb.add_line_segments(pts[1:], close=True)
    shp = fb.convert_to_shape(); shp.fill.solid(); shp.fill.fore_color.rgb = color; shp.line.fill.background()
    kit.no_effects(shp)
    return shp


def slide(n, tg, headline, source=None):
    s = kit.new_slide()
    kit.frame(s, n, None, headline, source)
    tag(kit, s, tg)
    return s


def bookend(title, sub, small):
    s = kit.new_slide()
    kit.shape(s, MSO_SHAPE.RECTANGLE, 0, 0, W, H, TEAL)
    kit.photo(s, 8.6, 0, W - 8.6, H, label="Photo: seaweed farm at low tide, Portuguese coast")
    wave(s, 8.45, 0, H, TEAL, flip=True)
    kit.text(s, M, 2.5, 7.4, 1.3, title, 40, T["white"], bold=True)
    kit.text(s, M, 3.95, 7.4, 0.5, sub, 20, PALE)
    if small:
        kit.text(s, M, 6.7, 7.4, 0.35, small, 14, kit.tint(TEAL, 0.55))
    return s


# 1 title
bookend("Kelpwise", "Investment recommendation: seaweed packaging for Portugal's coast",
        "Investment committee  |  Illustrative example with fictional data")

# 2 recommendation up front
s = slide(2, "Recommendation", "We suggest offering $300K debt financing to Kelpwise, a seaweed-packaging startup in Portugal")
rows = [("Plastic is choking the coast", "Over 90,000 tons of food packaging reach Portugal's coastal waters every year"),
        ("Kelpwise can replace it", "Seaweed trays at almost the price of plastic, made end to end on the coast"),
        ("The team can deliver", "Experienced founders, partner farms, and €1.8M of contracted sales"),
        ("A $300K loan unlocks the round", "20% of a $1.5M seed round that funds a second processing line")]
rh, gap = 0.85, 0.2
top = kit.place_balanced(4 * rh + 3 * gap)
for i, (lab, txt) in enumerate(rows):
    y = top + i * (rh + gap)
    kit.box(s, M, y, 3.9, rh, TEAL if i < 3 else CORAL)
    kit.text(s, M + 0.3, y, 3.4, rh, lab, 18, T["white"], bold=True, anchor=MSO_ANCHOR.MIDDLE)
    kit.text(s, M + 4.3, y, CW - 4.3, rh, txt, 20, INK, bold=(i == 3), anchor=MSO_ANCHOR.MIDDLE)

# 3 the problem: chart
s = slide(3, "The problem", "Every year, >90,000 tons of single-use food packaging end up in Portugal's coastal waters",
          "Source: coastal waste monitoring programme, 2019–2025 (fictional data)")
ch_h = 3.6
top, cy = chart_slide_layout(kit, s, "Food packaging reaching coastal waters, 2019–2025", CW, ch_h)
pt_, pb_, ch, _ = bar_chart(kit, s, M, cy, CW, ch_h, ["2019", "2020", "2021", "2022", "2023", "2024", "2025"],
                            [71, 74, 78, 81, 85, 88, 92], highlight=(6,), axis_title="Thousand tons", maximum=100)
ch.value_axis.major_unit = 25

# 4 the damage: three pillars
s = slide(4, "The problem", "This waste damages fisheries, tourism income, and marine habitats",
          "Note: illustrative estimates for this example (fictional data).")
pillars = [("ship", "Fisheries", "€40M", "a year in lost catch", ["Plastic fragments in nets and fish stocks", "Smaller catches near river mouths"]),
           ("sun", "Tourism", "€120M", "a year in clean-up and lost visits", ["Beach closures after storms", "Lower visitor ratings on affected coasts"]),
           ("leaf", "Habitats", "1 in 3", "seagrass meadows damaged", ["Nurseries for young fish lost", "Less carbon stored in coastal plants"])]
pw, pg = (CW - 2 * 0.5) / 3, 0.5
top = kit.place_balanced(3.3)
for i, (ic, name, num, lab, pts) in enumerate(pillars):
    x = M + i * (pw + pg)
    kit.icon(s, ic, x, top, 0.6, circle=TEAL)
    kit.text(s, x + 0.8, top + 0.08, pw - 0.8, 0.45, name, 20, TEAL, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    kit.text(s, x, top + 0.85, pw, 0.75, num, 40, CORAL, bold=True, smart=False)
    kit.text(s, x, top + 1.6, pw, 0.4, lab, 16, INK, bold=True)
    kit.bullet_list(s, x, top + 2.2, pw, 1.6, pts, pt=16, gap_pt=8)

# 5 what would fix it: cost parity chart
s = slide(5, "The solution", "To solve this, packaging needs a compostable alternative that costs the same as plastic",
          "Note: cost per 1,000 single-portion food trays, delivered. Source: supplier quotes, 2026 (fictional data)")
ch_w, ch_h = 8.4, 3.6
top, cy = chart_slide_layout(kit, s, "Cost per 1,000 food trays (€)", ch_w, ch_h, panel_title="What it takes")
pt_, pb_, ch, (PX, PY, PW, PH) = bar_chart(kit, s, M, cy, ch_w, ch_h, ["Plastic (today)", "Paperboard", "PLA bioplastic", "Kelpwise seaweed"],
                                            [38, 52, 61, 39], horizontal=True, colors=[T["grey"], None, None, CORAL],
                                            axis_title="€ per 1,000 trays", maximum=70)
rx = M + ch_w + 0.5
kit.side_points(s, rx, W - M - rx, pt_, pb_, [("Compostable:", "breaks down at sea and in home compost"),
                                               ("Same price:", "Kelpwise is within €1 of plastic"),
                                               ("Same use:", "fits existing packing lines")], lead_color=TEAL)

# 6 who can do it: value chain tick table
s = slide(6, "The company", "Kelpwise is the only company in the region to grow, process, and sell seaweed packaging end to end",
          "Note: regional packaging and seaweed businesses, Portugal and Galicia, 2026 (fictional).")
stages = ["Grow", "Harvest", "Process", "Make trays", "Sell"]
players = [("Kelpwise", [1, 1, 1, 1, 1]), ("Seaweed farm co-op", [1, 1, 0, 0, 0]),
           ("Film importer", [0, 0, 1, 1, 0]), ("Paper tray maker", [0, 0, 0, 1, 1])]
c0, cw_ = 3.4, (CW - 3.4) / 5
rh = 0.8
top = kit.place_balanced(0.55 + rh * 4)
for j, st in enumerate(stages):
    kit.text(s, M + c0 + j * cw_, top, cw_, 0.45, st, 18, TEAL, bold=True, align=PP_ALIGN.CENTER, smart=False)
for i, (nm, ticks) in enumerate(players):
    y = top + 0.55 + i * rh
    if i == 0:
        kit.box(s, M - 0.1, y, CW + 0.2, rh, PALE)
    kit.text(s, M, y, c0 - 0.2, rh, nm, 18, INK, bold=(i == 0), anchor=MSO_ANCHOR.MIDDLE, smart=False)
    for j, t_ in enumerate(ticks):
        cx = M + c0 + j * cw_ + cw_ / 2
        if t_:
            kit.icon(s, "circle-check", cx - 0.2, y + rh / 2 - 0.2, 0.4, color=CORAL if i == 0 else T["grey"])
        else:
            kit.shape(s, MSO_SHAPE.OVAL, cx - 0.06, y + rh / 2 - 0.06, 0.12, 0.12, T["light"])

# 7 impact
s = slide(7, "The company", "Kelpwise has the potential to deliver significant environmental and economic impact by 2030",
          "Note: company projections for 2030, reviewed by our team (fictional data).")
nums = [("12,000 t", "of plastic packaging replaced a year"), ("140", "coastal jobs, most in seaweed farming"), ("25", "partner seaweed farms along the coast")]
tw, tg = (CW - 2 * 0.5) / 3, 0.5
top = kit.place_balanced(2.4)
for i, (n_, lab) in enumerate(nums):
    x = M + i * (tw + tg)
    kit.box(s, x, top, tw, 2.4, PALE)
    kit.text(s, x + 0.4, top + 0.45, tw - 0.8, 0.85, n_, 44, TEAL, bold=True, smart=False)
    kit.text(s, x + 0.4, top + 1.35, tw - 0.8, 0.8, lab, 18, INK)

# 8 team
s = slide(8, "The company", "Kelpwise's leadership team has the commitment and expertise to scale across Portugal and into Spain")
team = [("Ana Ribeiro", "Chief executive", "Ran operations for a national fish-processing group for 9 years"),
        ("Tomás Ferreira", "Chief operating officer", "Built and runs the first seaweed processing line"),
        ("Inês Costa", "Head of materials", "PhD in biopolymers; holds the tray-forming patent")]
tw, tg = (CW - 2 * 0.6) / 3, 0.6
top = kit.place_balanced(3.7)
for i, (nm, role, line) in enumerate(team):
    x = M + i * (tw + tg)
    d = 1.6
    kit.photo(s, x + (tw - d) / 2, top, d, d, label="Headshot", crop="circle")
    kit.text(s, x, top + d + 0.25, tw, 0.4, nm, 20, TEAL, bold=True, align=PP_ALIGN.CENTER, smart=False)
    kit.text(s, x, top + d + 0.7, tw, 0.35, role, 16, T["grey"], align=PP_ALIGN.CENTER, smart=False)
    kit.text(s, x + 0.2, top + d + 1.15, tw - 0.4, 0.8, line, 16, INK, align=PP_ALIGN.CENTER)

# 9 partnerships: chart
s = slide(9, "The company", "Kelpwise has also signed supply partnerships that will bolster its sales and revenue",
          "Note: signed supply agreements, revenue over three years from 2027 (fictional data).")
ch_w, ch_h = 8.4, 3.6
top, cy = chart_slide_layout(kit, s, "Contracted revenue by partner, 2027–2029 (€M)", ch_w, ch_h, panel_title="Why it matters")
pt_, pb_, ch, _ = bar_chart(kit, s, M, cy, ch_w, ch_h, ["Mar Azul Supermarkets", "Porto Fresh Fish Co-op", "Atlântica Catering"],
                            [0.9, 0.5, 0.4], horizontal=True, highlight=(0, 1, 2), axis_title="€ millions", maximum=1.2, fmt='0.0')
rx = M + ch_w + 0.5
kit.side_points(s, rx, W - M - rx, pt_, pb_, [("€1.8M", "contracted before the new line opens"),
                                               ("Three sectors:", "retail, wholesale, and catering")], lead_color=TEAL)

# 10 the recommendation
s = slide(10, "The investment", "We recommend a $300K debt investment in Kelpwise as part of its $1.5M seed round",
          "Note: proposed terms subject to due diligence (fictional).")
parts = [("Lead venture fund", 800, T["light"]), ("Impact fund", 400, T["light"]), ("Our loan", 300, CORAL)]
top = kit.place_balanced(0.55 + 0.95 + 0.45 + 0.85 + 0.45 + 0.85)
kit.chart_header(s, M, top, CW, "$1.5M seed round ($ thousands)")
x, y, bw = M, top + 0.6, CW
for nm, v, col in parts:
    w_ = bw * v / 1500
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, y, w_ - 0.04, 0.95, col)
    tc = T["white"] if col == CORAL else INK
    kit.text(s, x, y, w_ - 0.04, 0.95, [[(nm, True, tc)], [(f"{v:,}", False, tc)]], 16, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    x += w_
terms = [("3-year loan", "at 8% interest, repaid from year 2"), ("Use of funds", "second processing line and three new partner farms"),
         ("Board observer", "seat for our investment lead")]
ty = top + 0.55 + 0.95 + 0.45
tw = (CW - 2 * 0.5) / 3
for i, (a_, b_) in enumerate(terms):
    kit.text(s, M + i * (tw + 0.5), ty, tw, 0.85, [[(a_, True, TEAL)], [(b_, False, INK)]], 18)
kit.bar(s, ty + 0.85 + 0.45, "Decision: Does the committee approve a $300K loan to Kelpwise as part of its seed round?", fill=TEAL)

# 11 results: timeline
s = slide(11, "The investment", "The seed round would deliver tangible results within 18 months",
          "Note: milestones from the company's plan, tested in our diligence (fictional).")
ms = [("Q2 2027", "Second processing line opens"), ("Q4 2027", "Three new partner farms harvesting"),
      ("Q1 2028", "Capacity reaches 4,000 tons a year"), ("Q3 2028", "Break-even on monthly sales")]
top = kit.place_balanced(2.0)
xs = kit.distribute_x([2.6] * 4, M, W - M)
ty = top + 0.9
trk = kit.shape(s, MSO_SHAPE.PENTAGON, M, ty - 0.08, CW, 0.16, PALE); trk.adjustments[0] = 0.5
for x, (when, what) in zip(xs, ms):
    cx = x + 1.3
    kit.text(s, x, top, 2.6, 0.45, when, 20, TEAL, bold=True, align=PP_ALIGN.CENTER, smart=False)
    kit.shape(s, MSO_SHAPE.OVAL, cx - 0.12, ty - 0.12, 0.24, 0.24, CORAL, line=T["white"], lw=1.5)
    kit.text(s, x, ty + 0.4, 2.6, 1.0, what, 18, INK, align=PP_ALIGN.CENTER)

# 12 risks
s = slide(12, "Risks", "This investment isn't risk-free, but its main risks can be mitigated",
          "Note: likelihood after mitigation, team judgement (fictional).")
risks = [("Seaweed harvests fall in a bad year", "Contracts with farms along 200 km of coast spread the weather risk"),
         ("Plastic prices drop", "Retail partners have committed to compostable packaging by 2028"),
         ("Scaling the new line takes longer", "Repayments start only in year 2; equipment supplier guarantees output"),
         ("A larger competitor enters", "Tray-forming patent and long-term farm contracts")]
rh = 0.85
top = kit.place_balanced(0.5 + 4 * rh - 0.25)   # last row has no band: measure visible text
kit.text(s, M, top, 5.0, 0.4, "Risk", 18, TEAL, bold=True, smart=False)
kit.text(s, M + 5.4, top, CW - 5.4, 0.4, "How it's mitigated", 18, TEAL, bold=True, smart=False)
for i, (r_, m_) in enumerate(risks):
    y = top + 0.5 + i * rh
    if i % 2 == 0:
        kit.box(s, M - 0.1, y, CW + 0.2, rh, PALE)
    kit.text(s, M, y, 5.0, rh, r_, 18, INK, anchor=MSO_ANCHOR.MIDDLE)
    kit.text(s, M + 5.4, y, CW - 5.6, rh, m_, 18, INK, anchor=MSO_ANCHOR.MIDDLE)

# 13 close
bookend("Thank you", "Questions and discussion", None)

print(kit.save(out("kelpwise-pitch.pptx")))
