"""Example: a proposal (pitching for work), storytelling §5b and layout-patterns §I. Fictional consultancy, client,
people, clients on the logo wall, quote, and figures; the logos are plain-text wordmarks made up for the example.
Theme: navy + client teal + amber highlight, square corners; co-branded lockup bottom-left on every slide;
a recurring hero image on section dividers; photo placeholders where the user would drop in real photos."""
from _shared import *  # noqa: F401,F403

NAVY, TEAL, AMBER, PALE, INK = rgb("1B2F5B"), rgb("0B7A75"), rgb("E9A23B"), rgb("EAF1F4"), rgb("1E2430")
kit = Kit(dict(body="Arial", main="1B2F5B", main_pale="EAF1F4", accent="E9A23B", ink="1E2430", grey="667085",
               light="C9D2DC", tag_labels=("x",), foot_top=6.45, page_num_top=6.94))
T, CW, M, W, H = kit.T, kit.CW, kit.T["M"], kit.T["W"], kit.T["H"]
HERO = "Photo: parent and child reading together"

# ------------------------------------------------------------ accuracy checks for every derived figure
SPRINT_WEEKS, CORE_SPRINTS = 2, 4
assert SPRINT_WEEKS * CORE_SPRINTS == 8                         # "8-week core phase"
assert 14 * CORE_SPRINTS == 56 and 9 * CORE_SPRINTS == 36       # Regular £56K, Light £36K for the core phase
assert round(56 / 0.8) == 70 and round(36 / 0.8) == 45          # both include the 20% charity discount
assert 2 * 3100 == 6200                                         # "double the children reached" (3,100 → 6,200)


def lockup(s, x, y, dark=False):
    """Co-branded lockup (visual-identity §4c): client first, thin divider, then the consultancy; one baseline."""
    c1, c2 = (T["white"], T["white"]) if dark else (TEAL, NAVY)
    kit.text(s, x, y, 1.15, 0.3, [[("TIDEWATER", True, c1)]], 12, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    kit.line(s, x + 1.25, y + 0.04, x + 1.25, y + 0.26, kit.tint(NAVY, 0.6) if not dark else PALE, 0.75)
    kit.text(s, x + 1.35, y, 1.2, 0.3, [[("Northlight", True, c2)]], 12, anchor=MSO_ANCHOR.MIDDLE, smart=False)


def slide(n, headline, source=None):
    s = kit.new_slide()
    kit.frame(s, n, None, headline, source)
    lockup(s, 0.3, H - 0.3 - 0.3)                    # bottom-left corner, same gap to both edges
    return s


def divider(num, title, sub):
    s = kit.new_slide()
    kit.photo(s, 0, 0, W, H, label=HERO)
    kit.overlay(s, 0, 0, W * 0.6, H, NAVY, mode="gradient", opacity=0.92, stops=[(0, 0.92), (0.7, 0.85), (1, 0)])
    kit.text(s, M, 2.3, 3, 1.0, f"{num:02d}", 60, AMBER, bold=True, smart=False)
    kit.text(s, M, 3.4, 7, 0.9, title, 36, T["white"], bold=True)
    kit.text(s, M, 4.35, 6.5, 0.9, sub, 18, PALE)
    lockup(s, 0.3, H - 0.6, dark=True)
    return s


# 1 title: half-bleed photo
s = kit.new_slide()
kit.photo(s, 6.4, 0, W - 6.4, H, label=HERO)
title = "A growth plan to reach twice as many children by 2035"
th = kit.text_height(kit.wrapped(title, 32, 5.4, True), 32, 5.4, True)
y = (H - (0.4 + 0.15 + th + 0.25 + 0.4 + 0.1 + 0.35)) / 2          # title block centred vertically
kit.text(s, M, y, 5.4, 0.4, "PROPOSAL", 14, TEAL, bold=True, smart=False); y += 0.55
kit.text(s, M, y, 5.4, th + 0.05, title, 32, NAVY, bold=True); y += th + 0.25
kit.text(s, M, y, 5.4, 0.4, "For Tidewater Children's Trust  |  January 2027", 18, INK, smart=False); y += 0.5
kit.text(s, M, y, 5.4, 0.35, "Illustrative example with fictional content", 14, T["grey"], smart=False)
lockup(s, 0.3, H - 0.6)

# 2 contents with a circle-cropped photo bleeding off the left edge
s = kit.new_slide()
kit.photo(s, -2.4, 0.9, 6.0, 6.0, label=HERO, crop="circle")
kit.text(s, 4.4, 1.3, W - M - 4.4, 1.3, "We'd be delighted to help you build a growth plan that doubles the number of children you reach by 2035.", 22, NAVY, bold=True)
kit.text(s, 4.4, 2.95, 4, 0.4, "Contents", 18, TEAL, bold=True, smart=False)
items = [("1", "Our understanding", "3"), ("2", "Our proposed approach", "5"), ("3", "Commercials", "8"), ("4", "Why us", "11")]
for i, (n_, t_, pg) in enumerate(items):
    y = 3.5 + i * 0.55
    kit.text(s, 4.4, y, 0.4, 0.4, n_, 18, TEAL, bold=True, smart=False)
    kit.text(s, 4.85, y, 5.5, 0.4, t_, 18, INK, smart=False)
    kit.text(s, W - M - 1.0, y, 1.0, 0.4, pg, 18, T["grey"], align=PP_ALIGN.RIGHT, smart=False)
    kit.line(s, 4.4, y + 0.47, W - M, y + 0.47, T["light"], 0.5)
lockup(s, 0.3, H - 0.6)
kit.text(s, W - M - 0.6, T["page_num_top"], 0.6, 0.3, "2", T["small_pt"], T["grey"], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.BOTTOM, smart=False)

# 3 divider
divider(1, "Our understanding", "Where you are, what you want to achieve, and what success looks like")

# 4 understanding: three columns headed by photo strips
s = slide(4, "You want a plan to reach twice as many children by 2035, without losing what makes you trusted")
cols = [("Your context", "Photo: a family support session", ["Reaching 3,100 children a year across four towns", "Demand rising faster than funding", "Strong local reputation"]),
        ("What you've asked for", "Photo: staff planning at a whiteboard", ["A growth plan to 2035", "Options to reach new areas", "A funding model that can scale"]),
        ("What success looks like", "Photo: children at an after-school club", ["6,200 children a year by 2035", "Two new areas, chosen on need", "Funding that grows with reach"])]
cw4 = (CW - 2 * 0.4) / 3
blh = max(sum(kit.text_height(kit.wrapped(p_, 16, cw4 - 0.55), 16, cw4 - 0.55) for p_ in c[2]) + 2 * 8 / 72 for c in cols)
block = 1.4 + 0.55 + 0.25 + blh
top = kit.place_balanced(block)
for j, (t_, ph, pts) in enumerate(cols):
    x = M + j * (cw4 + 0.4)
    kit.photo(s, x, top, cw4, 1.4, label=ph)
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, top + 1.4, cw4, 0.55, NAVY if j < 2 else TEAL)
    kit.text(s, x, top + 1.4, cw4, 0.55, t_, 18, T["white"], bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    kit.bullet_list(s, x + 0.1, top + 2.2, cw4 - 0.2, blh + 0.1, pts, pt=16, gap_pt=8)

# 5 phased offer: core + options
s = divider(2, "Our proposed approach", "An 8-week core phase, with two optional extensions to June")
s = slide(6, "We recommend an 8-week core phase, plus two optional extensions")
tiles = [("rocket", "Core", "8 weeks", "A growth plan with options for new areas and a funding model", NAVY, False),
         ("network", "Advanced", "4 weeks, optional", "Detailed business cases for the two chosen areas", TEAL, True),
         ("refresh-cw", "Embed", "4 weeks, optional", "Hand-over, training, and a plan for the first year of delivery", kit.tint(TEAL, 0.35), True)]
tw6 = 3.4; th6 = 3.3
xs = kit.distribute_x([tw6] * 3, M, W - M)
top = kit.place_balanced(th6)
for i, (ic, name, dur, d, col, opt) in enumerate(tiles):
    x = xs[i]
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, top, tw6, th6, col)
    kit.text(s, x, top + 0.2, tw6, 0.35, f"STAGE {i + 1}" + (" · OPTIONAL" if opt else ""), 14, T["white"], bold=True, align=PP_ALIGN.CENTER, smart=False)
    kit.icon(s, ic, x + tw6 / 2 - 0.35, top + 0.65, 0.7, color="FFFFFF")
    kit.text(s, x, top + 1.45, tw6, 0.5, name, 24, T["white"], bold=True, align=PP_ALIGN.CENTER, smart=False)
    kit.text(s, x, top + 1.95, tw6, 0.35, dur, 16, T["white"], align=PP_ALIGN.CENTER, smart=False)
    kit.text(s, x + 0.3, top + 2.35, tw6 - 0.6, 0.8, d, 16, T["white"], align=PP_ALIGN.CENTER)
    if i < 2:
        g0, g1 = x + tw6, xs[i + 1]
        mk = kit.shape(s, MSO_SHAPE.ISOSCELES_TRIANGLE, (g0 + g1) / 2 - 0.08, top + th6 / 2 - 0.1, 0.16, 0.2, T["light"]); mk.rotation = 90

# 7 delivery plan: a Gantt in two-week sprints
s = slide(7, "We'd work in two-week sprints, so you can review progress and steer",
          "Note: optional stages start only if you choose them at the end of the core phase.")
sprints = 8
lab_w = 2.4; gx = M + lab_w; gw = CW - lab_w; sw = gw / sprints
rows = [("Research", [(0, 2), (4, 5)]), ("Engagement", [(1, 3), (5, 6), (6, 8)]), ("Analysis", [(2, 4), (4, 6)]), ("Mobilisation", [(3, 4), (6, 8)])]
rh7 = 0.62
block = 0.45 + 0.4 + len(rows) * rh7
top = kit.place_balanced(block)
for k_, (name, a_, b_, col) in enumerate([("Core", 0, 4, NAVY), ("Advanced (optional)", 4, 6, TEAL), ("Embed (optional)", 6, 8, kit.tint(TEAL, 0.35))]):
    kit.shape(s, MSO_SHAPE.RECTANGLE, gx + a_ * sw + 0.02, top, (b_ - a_) * sw - 0.04, 0.4, col)
    kit.text(s, gx + a_ * sw, top, (b_ - a_) * sw, 0.4, name, 14, T["white"], bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, smart=False)
for i in range(sprints):
    kit.text(s, gx + i * sw, top + 0.45, sw, 0.35, f"Sprint {i + 1}", 14, T["grey"], align=PP_ALIGN.CENTER, smart=False)
for r_, (name, bars) in enumerate(rows):
    y = top + 0.85 + r_ * rh7
    if r_ % 2 == 0:
        kit.shape(s, MSO_SHAPE.RECTANGLE, M, y, CW, rh7, PALE)
    kit.text(s, M + 0.15, y, lab_w - 0.2, rh7, name, 16, NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    for a_, b_ in bars:
        col = NAVY if b_ <= 4 else TEAL
        kit.shape(s, MSO_SHAPE.RECTANGLE, gx + a_ * sw + 0.02, y + 0.17, (b_ - a_) * sw - 0.04, rh7 - 0.34, col)

# 8 divider
divider(3, "Commercials", "Two ways to staff the work, both with our charity discount")

# 9 pricing options + assumptions on a photo panel
s = slide(9, "Two team options let you choose how much support you want during the core phase",
          "Note: prices exclude VAT and include a 20% discount for registered charities (fictional).")
pw_ = 3.9
panel_x = M + CW - pw_
top = kit.place_balanced(3.7)
kit.photo(s, panel_x, top, pw_, 3.7, label="Photo: hands holding a mug, warm light")
kit.overlay(s, panel_x, top, pw_, 3.7, NAVY, mode="panel", opacity=0.85)
kit.text(s, panel_x + 0.3, top + 0.3, pw_ - 0.6, 0.4, "Our estimates assume", 18, T["white"], bold=True, smart=False)
kit.bullet_list(s, panel_x + 0.3, top + 0.85, pw_ - 0.6, 2.8, ["Staff and data available in the first two weeks", "One steering meeting every sprint", "Final scope agreed together at kick-off"],
                pt=16, color=T["white"], bullet_color=AMBER, gap_pt=10)
opts = [("users", "Regular", "Director, manager, and two consultants", "£14K", "£56K"), ("user", "Light", "Director, manager, and one consultant", "£9K", "£36K")]
cw9 = (panel_x - M - 0.5 - 0.3) / 2
for i, (ic, name, team, per, total) in enumerate(opts):
    x = M + i * (cw9 + 0.3)
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, top, cw9, 3.7, PALE)
    kit.icon(s, ic, x + 0.3, top + 0.3, 0.6, circle=NAVY)
    kit.text(s, x + 1.05, top + 0.3, cw9 - 1.2, 0.6, name, 22, NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    kit.text(s, x + 0.3, top + 1.15, cw9 - 0.6, 0.8, team, 16, INK)
    kit.text(s, x + 0.3, top + 2.0, cw9 - 0.6, 0.4, "Per two-week sprint", 14, T["grey"], smart=False)
    kit.text(s, x + 0.3, top + 2.35, cw9 - 0.6, 0.5, per, 24, NAVY, bold=True, smart=False)
    kit.text(s, x + 0.3, top + 2.95, cw9 - 0.6, 0.4, "Core phase, 8 weeks", 14, T["grey"], smart=False)
    kit.text(s, x + 0.3, top + 3.25, cw9 - 0.6, 0.5, total, 24, TEAL, bold=True, smart=False)

# 10 example deliverables
s = slide(10, "You'd come away with tools your team can keep using, not a report")
dels = [("layout-grid", "Growth options map"), ("pie-chart", "Funding model"), ("map", "Area needs analysis"),
        ("bar-chart-3", "Impact dashboard"), ("workflow", "Operating model"), ("calendar-check", "First-year delivery plan")]
dw, dh = 3.6, 1.75
xs = kit.distribute_x([dw] * 3, M, W - M)
top = kit.place_balanced(2 * dh + 0.3)
for i, (ic, t_) in enumerate(dels):
    x = xs[i % 3]; y = top + (i // 3) * (dh + 0.3)
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, y, dw, 0.5, NAVY)
    kit.text(s, x, y, dw, 0.5, t_, 16, T["white"], bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, y + 0.5, dw, dh - 0.5, PALE)
    kit.icon(s, ic, x + dw / 2 - 0.35, y + 0.5 + (dh - 0.5) / 2 - 0.35, 0.7, color=str(TEAL))

# 11 divider
divider(4, "Why us", "Results we've delivered, the team you'd work with, and what clients say")

# 12 impact numbers + statement panel
s = slide(12, "Our recent work has helped charities and public bodies grow their reach and income")
nums = [("£12M", "new funding secured for a youth charity"), ("40%", "more families reached by a parenting service"), ("3x", "growth in regular givers for a food bank network"),
        ("18", "local authorities supported on early help"), ("£4M", "annual savings for a hospital trust"), ("92%", "of clients would hire us again")]
gx = M + 3.9
nw = (W - M - gx - 2 * 0.3) / 3; nh = 1.7
top = kit.place_balanced(2 * nh + 0.3)
kit.photo(s, M, top, 3.6, 2 * nh + 0.3, label="Photo: a youth worker with a teenager")
kit.overlay(s, M, top, 3.6, 2 * nh + 0.3, NAVY, mode="panel", opacity=0.85)
kit.text(s, M + 0.3, top + 0.3, 3.0, 2 * nh - 0.3, "Results that matter are measured in lives, not just money: every number here is about more people reached or more funding for the front line.", 18, T["white"], bold=True)
for i, (n_, c_) in enumerate(nums):
    x = gx + (i % 3) * (nw + 0.3); y = top + (i // 3) * (nh + 0.3)
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, y, nw, nh, PALE)
    kit.text(s, x + 0.25, y + 0.2, nw - 0.5, 0.65, n_, 32, TEAL, bold=True, smart=False)
    kit.text(s, x + 0.25, y + 0.85, nw - 0.5, 0.8, c_, 14, INK)

# 13 team
s = slide(13, "You'd work with a small senior team who have done this before")
team = [("Ruth Okafor", "Director", "Led growth plans for six national charities"), ("James Patel", "Senior manager", "Ran a children's services charity for eight years"),
        ("Elena Varga", "Consultant", "Specialist in funding and financial models"), ("Sam Whitfield", "Consultant", "Data analyst; builds needs and impact dashboards")]
tw13 = 2.6
xs = kit.distribute_x([tw13] * 4, M, W - M)
top = kit.place_balanced(3.2)
for i, (nm, role, line) in enumerate(team):
    x = xs[i]
    kit.photo(s, x + tw13 / 2 - 0.75, top, 1.5, 1.5, label="Headshot", crop="circle")
    kit.text(s, x, top + 1.7, tw13, 0.4, nm, 18, NAVY, bold=True, align=PP_ALIGN.CENTER, smart=False)
    kit.text(s, x, top + 2.1, tw13, 0.35, role, 14, TEAL, bold=True, align=PP_ALIGN.CENTER, smart=False)
    kit.text(s, x + 0.1, top + 2.5, tw13 - 0.2, 0.7, line, 14, INK, align=PP_ALIGN.CENTER)

# 14 case study
s = slide(14, "Case study: we helped a family charity decide where to grow next",
          "Note: client name withheld; illustrative example with fictional details.")
blocks = [("users", "The client", "A regional family-support charity working in three cities"),
          ("target", "The challenge", "Demand doubling, and no agreed view on which areas to grow into first"),
          ("lightbulb", "Our solution", "Needs mapping across 30 areas, three growth options, and a funding model the board could test")]
lw = 7.2; ph_x = M + lw + 0.5; ph_w = W - M - ph_x
top = kit.place_balanced(3.9)
for i, (ic, k_, v) in enumerate(blocks):
    y = top + i * 1.35
    kit.icon(s, ic, M, y, 0.6, circle=TEAL)
    kit.text(s, M + 0.85, y, lw - 0.85, 0.4, k_, 18, NAVY, bold=True, smart=False)
    kit.text(s, M + 0.85, y + 0.42, lw - 0.85, 0.9, v, 16, INK)
kit.photo(s, ph_x, top, ph_w, 3.9, label="Photo: a family at a community centre")
kit.overlay(s, ph_x, top, ph_w, 3.9, NAVY, mode="panel", opacity=0.85)
kit.text(s, ph_x + 0.3, top + 0.3, ph_w - 0.6, 0.45, "The outcomes", 20, T["white"], bold=True, smart=False)
kit.bullet_list(s, ph_x + 0.3, top + 0.9, ph_w - 0.6, 3.1, ["Board agreed two new areas in one meeting", "£2.5M bid built on the needs analysis", "First new service opened within a year"],
                pt=16, color=T["white"], bullet_color=AMBER, gap_pt=12)

# 15 client quote (fictional, for the example only: real decks use real quotes with permission)
s = slide(15, "Clients tell us we make complex decisions feel manageable")
top = kit.place_balanced(3.0)
kit.photo(s, M, top, 3.0, 3.0, label="Portrait of the quoted client")
kit.text(s, M + 3.5, top - 0.2, 1.0, 1.0, "“", 72, AMBER, bold=True, smart=False)
kit.text(s, M + 3.5, top + 0.7, CW - 3.8, 1.5, "Northlight helped our board see the choices clearly and agree a plan we all believe in. We started delivering it the following month.", 22, NAVY)
kit.text(s, M + 3.5, top + 2.4, CW - 3.8, 0.4, "Chief executive, a regional family charity (fictional quote)", 16, T["grey"], smart=False)

# 16 logo wall (made-up wordmarks, all one tone, equal weight)
s = slide(16, "We've worked with charities, funders, and public bodies of every size")
names = ["Brightside Trust", "Harbour Health", "Moorland Council", "Kinfolk Foundation", "Lantern Youth", "Oakway Housing",
         "Riverside NHS Trust", "Stepping Stones", "Wyvern Schools", "Northgate Fund", "Seacroft Care", "Tilly's Kitchen"]
lw_, lh = 2.6, 0.9
xs = kit.distribute_x([lw_] * 4, M, W - M)
top = kit.place_balanced(3 * lh + 2 * 0.35)
for i, nm in enumerate(names):
    x = xs[i % 4]; y = top + (i // 4) * (lh + 0.35)
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, y, lw_, lh, PALE)
    kit.text(s, x, y, lw_, lh, nm, 16, kit.tint(NAVY, 0.25), bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, smart=False)

# 17 close
s = kit.new_slide()
kit.photo(s, 6.4, 0, W - 6.4, H, label=HERO)
kit.text(s, M, 2.5, 5.4, 0.8, "Thank you", 40, NAVY, bold=True, smart=False)
kit.text(s, M, 3.4, 5.4, 0.9, "We'd love to talk through this proposal and shape it with you.", 20, INK)
kit.text(s, M, 4.45, 5.4, 0.9, [[("Ruth Okafor, Director", True, NAVY)], [("ruth.okafor@example.com", False, INK)]], 16)
lockup(s, 0.3, H - 0.6)

print(kit.save(out("consulting-proposal.pptx")))
