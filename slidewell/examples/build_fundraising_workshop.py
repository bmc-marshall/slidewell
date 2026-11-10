"""Example: a facilitated workshop with a stepped method (storyline-examples §4, the workshop arc):
open → why it matters → the method, one step at a time → what it means for us → close. Fictional charity, people,
and figures. The method is a generic campaign-planning sequence. Theme: deep green + warm orange accent, square
corners; navigation: the agenda tracked small at the bottom centre; presented live, so the method steps are
animated (fade in on click, earlier steps stay at full strength because they are equal)."""
from _shared import *  # noqa: F401,F403

GREEN, ORANGE, PALE, INK = rgb("1B4D3E"), rgb("E07A1F"), rgb("E8F0EC"), rgb("1E2622")
kit = Kit(dict(body="Arial", main="1B4D3E", main_pale="E8F0EC", accent="E07A1F", ink="1E2622", grey="66716C",
               light="C6D3CC", muted_ink="A6AFAA", tag_labels=("x",), foot_top=6.5, page_num_top=6.94, footer_top=6.98))
T, CW, M, W, H = kit.T, kit.CW, kit.T["M"], kit.T["W"], kit.T["H"]
AGENDA = [("10:00", "Welcome", "Introductions and warm-up"), ("10:15", "Why now", "What the numbers tell us"),
          ("10:30", "The method", "Four steps to a campaign plan"), ("11:00", "Our plan", "Draft it together"),
          ("11:40", "Next steps", "Who does what, by when")]
STEPS = [("target", "Goal", "How much, by when, and for what?"), ("users", "Audience", "Who gives, and why would they?"),
         ("megaphone", "Message", "One story, told the same way everywhere"), ("calendar-check", "Plan", "Channels, dates, and who leads each one")]

# ------------------------------------------------------------ accuracy checks for every derived figure
mins = lambda t: int(t[:2]) * 60 + int(t[3:])
assert mins("12:00") - mins("10:00") == 120                    # "2 hours"
assert mins("11:40") - mins("11:00") == 40                     # "Our plan: 40 minutes", the longest slot
assert max(mins(b) - mins(a) for a, b in zip(["10:00", "10:15", "10:30", "11:00", "11:40"],
                                               ["10:15", "10:30", "11:00", "11:40", "12:00"])) == 40
donations = {"Regular givers": 38, "One-off appeals": 24, "Events": 22, "Grants": 16}
assert sum(donations.values()) == 100
assert round(38 / 100 * 120) == 46                             # £46K from regular givers out of £120K


def tracker(s, current, on_dark=False):
    pt, gap = 12, 0.3
    names = [a[1] for a in AGENDA]
    slot = max(kit.width_in(n, pt, True) for n in names) + 0.1
    x = (W - (slot * len(names) + gap * (len(names) - 1))) / 2
    for i, n in enumerate(names):
        on = i == current
        col = (T["white"] if on else kit.tint(GREEN, 0.5)) if on_dark else (GREEN if on else T["muted_ink"])
        kit.text(s, x, 6.98, slot, 0.26, n, pt, col, bold=on, align=PP_ALIGN.CENTER, smart=False)
        if i < len(names) - 1:
            kit.shape(s, MSO_SHAPE.OVAL, x + slot + gap / 2 - 0.03, 7.08, 0.06, 0.06, kit.tint(GREEN, 0.5) if on_dark else T["light"])
        x += slot + gap


def slide(n, label, headline, source=None, step=None):
    s = kit.new_slide()
    kit.frame(s, n, None, headline, source)
    kit.text(s, M, 0.26, 6, 0.32, label.upper(), 14, T["grey"], bold=True, smart=False)    # grey: distinct from headline
    if step is not None:
        tracker(s, step)
    return s


def bookend(title, sub, small):
    s = kit.new_slide()
    kit.shape(s, MSO_SHAPE.RECTANGLE, 0, 0, W, H, GREEN)
    kit.photo(s, 7.6, 0, W - 7.6, H, label="Photo: volunteers packing food boxes")
    kit.shape(s, MSO_SHAPE.RECTANGLE, 7.6, 0, 0.12, H, ORANGE)                                 # motif: the orange edge
    kit.text(s, M, 2.5, 6.6, 1.4, title, 40, T["white"], bold=True)
    kit.text(s, M, 4.0, 6.6, 0.9, sub, 20, PALE)
    if small:
        kit.text(s, M, 6.7, 6.6, 0.35, small, 14, kit.tint(GREEN, 0.55))
    return s


# 1 title
bookend("Planning our first fundraising campaign", "Workshop with trustees and volunteers  |  2 hours",
        "Riverbank Food Hub  |  Illustrative example with fictional content")

# 2 objectives
s = slide(2, "Objectives", "By midday we'll have a draft plan for our first campaign, and owners for each part")
objs = [("Agree", "why a campaign, and why now"), ("Learn", "a simple four-step method for planning one"), ("Draft", "our own plan, with a goal, an audience, a message, and dates")]
row = 1.0
top = kit.place_balanced(0.8 + 2 * row + 0.6)
kit.chart_header(s, M, top, CW, "By the end of this session, we will:")
for i, (verb, rest) in enumerate(objs):
    y = top + 0.8 + i * row
    kit.badge(s, M + 0.3, y + 0.3, 0.6, str(i + 1), pt=16)
    kit.text(s, M + 0.9, y, CW - 1.5, 0.6, [[(verb + " ", True, GREEN), (rest, False, INK)]], 20, anchor=MSO_ANCHOR.MIDDLE)

# 3 agenda (introduces the tracker)
s = slide(3, "Agenda", "We'll spend most of the morning drafting our own plan together")
COLW = 2.2
xs = kit.distribute_x([COLW] * len(AGENDA))
block = 0.2 + 0.4 + 0.3 + 0.36 + 0.3 + 0.4 + 0.65 + 0.15 + 0.55 + 0.15
top = kit.place_balanced(block)
ty_time = top + 0.2
ty = ty_time + 0.4 + 0.3 + 0.18
trk = kit.shape(s, MSO_SHAPE.PENTAGON, M, ty - 0.07, CW, 0.14, PALE); trk.adjustments[0] = 0.5
fx0, fx1 = xs[3] - 0.12, xs[3] + COLW + 0.12
kit.shape(s, MSO_SHAPE.RECTANGLE, fx0, top, fx1 - fx0, block, None, line=kit.tint(GREEN, 0.45), lw=1.5)
for i, (t_, name, desc) in enumerate(AGENDA):
    x = xs[i]; cx = x + COLW / 2
    focus = i == 3
    kit.text(s, x, ty_time, COLW, 0.4, t_, 18, GREEN, bold=focus, align=PP_ALIGN.CENTER, smart=False)
    kit.badge(s, cx, ty, 0.36, "", on=True)
    kit.text(s, x, ty + 0.48, COLW, 0.4, name, 18, INK, bold=True, align=PP_ALIGN.CENTER, smart=False)
    kit.text(s, x + 0.1, ty + 0.9, COLW - 0.2, 0.65, desc, 14, T["grey"], align=PP_ALIGN.CENTER)
kit.text(s, fx0, top + block - 0.68, fx1 - fx0, 0.6, "11:00–11:40\v40 minutes", 16, GREEN, bold=True, align=PP_ALIGN.CENTER, smart=False)

# 4 warm-up question to the room (full bleed)
s = kit.new_slide()
kit.shape(s, MSO_SHAPE.RECTANGLE, 0, 0, W, H, GREEN)
kit.text(s, M, 0.26, 6, 0.32, "WARM-UP", 14, kit.tint(GREEN, 0.6), bold=True, smart=False)
q = "Think of a cause you've given to. What made you give?"
qh = kit.text_height(kit.wrapped(q, 36, 10.5, True), 36, 10.5, True)
y = (H - (qh + 0.25 + 0.4)) / 2
kit.text(s, (W - 10.5) / 2, y, 10.5, qh + 0.05, q, 36, T["white"], bold=True, align=PP_ALIGN.CENTER)
kit.text(s, (W - 10.5) / 2, y + qh + 0.25, 10.5, 0.4, "Two minutes with your neighbour, then three answers to the room", 18, PALE, align=PP_ALIGN.CENTER)
tracker(s, 0, on_dark=True)

# 5 why now: income by source
s = slide(5, "Why now", "Regular givers bring in 38% of our income, but we've never asked",
          "Note: share of the £120K raised in the year to March 2027. Source: hub accounts (fictional data)", step=1)
CH_W, CH_H = 7.6, 3.3
top, cy = chart_slide_layout(kit, s, "Income by source, year to March 2027 (%)", CH_W, CH_H, "What it tells us")
pt_, pb, ch, geo = bar_chart(kit, s, M, cy, CH_W, CH_H, list(donations), list(donations.values()), horizontal=True,
                             axis_title="% of income", maximum=50, pt_cat=16, colors=[GREEN, None, None, None])
rx = M + CH_W + 0.5
kit.side_points(s, rx, W - M - rx, pt_, pb, [("Regular givers:", "£46K a year, from 310 people, with no campaign behind it"),
                                              ("One-off appeals:", "raised whenever the fridge broke, never planned"),
                                              ("The gap:", "demand for food parcels is rising faster than income")], lead_color=GREEN)

# 6 the method: four steps, animated (equal weight: fade in, no dimming)
s = slide(6, "The method", "A campaign plan answers four questions, in order", step=2)
tw = 2.75
xs = kit.distribute_x([tw] * 4, M, W - M)
th = 2.7
top = kit.place_balanced(th)
groups = []
for i, (ic, name, qn) in enumerate(STEPS):
    x = xs[i]
    g = [kit.box(s, x, top, tw, th, PALE)]
    g.append(kit.icon(s, ic, x + tw / 2 - 0.35, top + 0.35, 0.7, circle=GREEN))
    g.append(kit.text(s, x, top + 1.2, tw, 0.45, f"{i + 1}. {name}", 20, GREEN, bold=True, align=PP_ALIGN.CENTER, smart=False))
    g.append(kit.text(s, x + 0.2, top + 1.7, tw - 0.4, 0.9, qn, 16, INK, align=PP_ALIGN.CENTER))
    groups.append(g)
kit.animate_on_click(s, groups)                    # click sequence: step 1, 2, 3, 4

# 7 exercise
s = slide(7, "Exercise", "In four groups, draft one step each, then we'll put the plan together", step=3)
specs = [("Task", "Answer your step's question for our first campaign, on the template"), ("Groups", "Four tables, one step each"),
         ("Time", "20 minutes, then three minutes per table to share"), ("Output", "One sheet per step, ready to join up")]
G_H = 3.9
top = kit.place_balanced(G_H)
rh = G_H / 4
for i, (k_, v) in enumerate(specs):
    y = top + i * rh
    kit.text(s, M, y, 1.4, rh, k_.upper(), 14, GREEN, bold=True, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    kit.text(s, M + 1.5, y, 4.6, rh, v, 18, INK, anchor=MSO_ANCHOR.MIDDLE)
    if i:
        kit.line(s, M, y, M + 6.1, y, PALE, 1)
gx = M + 6.6; gw = (W - M - gx - 3 * 0.2) / 4
for i, (_, name, qn) in enumerate(STEPS):
    x = gx + i * (gw + 0.2)
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, top, gw, 0.5, GREEN)
    kit.text(s, x, top, gw, 0.5, name, 16, T["white"], bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, top + 0.5, gw, G_H - 0.5, PALE)

# 8 our plan on a page (what it means for us)
s = slide(8, "Our plan", "Our first campaign would aim to raise £15K from 150 new regular givers by December", step=3)
assert 150 * 10 * 10 == 15000                       # 150 givers × £10 a month × 10 months (March to December)
cells = [("Goal", "£15K by December: 150 new regular givers at £10 a month"), ("Audience", "Parents at the school gate and local businesses"),
         ("Message", "\"A full plate for every child this winter\""), ("Plan", "Launch at the September fair; posts weekly; letters to 40 businesses")]
cw8 = (CW - 0.3) / 2; ch8 = 1.45
top = kit.place_balanced(2 * ch8 + 0.3)
for i, (k_, v) in enumerate(cells):
    x = M + (i % 2) * (cw8 + 0.3); y = top + (i // 2) * (ch8 + 0.3)
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, y, cw8, ch8, PALE)
    kit.shape(s, MSO_SHAPE.RECTANGLE, x, y, 0.12, ch8, GREEN)
    kit.text(s, x + 0.4, y + 0.2, cw8 - 0.6, 0.4, k_.upper(), 14, GREEN, bold=True, smart=False)
    kit.text(s, x + 0.4, y + 0.6, cw8 - 0.6, 0.8, v, 18, INK)

# 9 next steps
s = slide(9, "Next steps", "Three owners will take the plan forward before trustees next meet", step=4)
nexts = [("Week 1", "Amira", "Turns today's sheets into a one-page plan"), ("Week 3", "Tom", "Sets up the regular-giving page"),
         ("Week 6", "Priya", "Books the September fair stall and drafts the business letter")]
rh = 0.85
top = kit.place_balanced(3 * rh + 2 * 0.25)
for i, (when, who, what) in enumerate(nexts):
    y = top + i * (rh + 0.25)
    kit.shape(s, MSO_SHAPE.RECTANGLE, M, y, 1.5, rh, GREEN)
    kit.text(s, M, y, 1.5, rh, when, 18, T["white"], bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    kit.text(s, M + 1.8, y, 2.0, rh, who, 18, GREEN, bold=True, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    kit.text(s, M + 3.9, y, CW - 3.9, rh, what, 18, INK, anchor=MSO_ANCHOR.MIDDLE)

# 10 close
bookend("Thank you", "Please leave one sticky note: what would make you give to us?", None)

print(kit.save(out("fundraising-workshop.pptx")))
