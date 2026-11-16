"""Example: a decision workshop comparing two options in parallel (inspired by a brand-creation workshop deck):
open over a photo → the agenda over a photo → what we're deciding → numbered sections with a wavy-edged divider →
the challenge (chart) and who we're designing for (persona) → each concept on the SAME two slides (idea + mood board;
who it wins and what it risks) → both scored on the same criteria → next steps.
Fictional dairy, agency, people, survey and figures. Theme: deep green + sun accent, subtle rounded corners.
Motif: rolling hills. Large on the title and closing slides; scaled down into a shallow footer band that holds the
co-branded lockup and the page number on one centre line; as the wavy edge of the section dividers. Section labels
are kickers above the headline (no boxed tag beside a curved motif). Presented live; one fade transition."""
import sys, math, datetime as dt
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from slidekit import Kit, rgb, CategoryChartData
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION

DEEP, MEADOW, HILL, OAT, SUN, INK = rgb("1F4D3A"), rgb("7FB069"), rgb("DCEAD3"), rgb("F4EDE0"), rgb("F2A541"), rgb("1E2A24")
LEAFTXT = rgb("4F8A5B")                       # kicker colour: the lighter brand green, readable on white
FOOT_MID = 7.2                                # centre line of the footer band (logos + page number)
kit = Kit(dict(body="Arial", display="Cambria", main="1F4D3A", main_pale="E8F0E9", accent="F2A541", ink="1E2A24",
               grey="5E6B64", light="CBD5CE", corners="rounded", radius=0.1, header_style="kicker",
               kicker_color=LEAFTXT, head_top=0.95, foot_top=6.36, footer_mid=FOOT_MID, head_gap=0.6,
               tag_labels=("x",)))
T, CW, M, W, H = kit.T, kit.CW, kit.T["M"], kit.T["W"], kit.T["H"]
WHITE, GREY, PALE, LIGHT = T["white"], T["grey"], T["main_pale"], T["light"]
TOP, BOT = 2.42, 6.2                          # content zone on every content slide

# ------------------------------------------------------------------ accuracy checks
under, over = [68, 49, 12, 18], [52, 74, 21, 63]
assert under[3] == 18 and over[3] == 63 and under[1] >= 49            # "only 18% ... against 63%"; "half trust it"
scores = [("Appeals to under-35s", 2, 3, "Sprout & Co. did better in two early groups"),
          ("Builds on Brightfield's trust", 3, 1, "Fieldwork carries the family name"),
          ("Stands out on the shelf", 1, 3, "Earth tones blend in; bright colours don't"),
          ("Room to stretch to new products", 2, 3, "A new name stretches to snacks and drinks"),
          ("Affordable to launch", 3, 1, "A new name costs about £400K more")]
assert sum(s[1] for s in scores) == sum(s[2] for s in scores) == 11      # "score evenly": 11 of 15 each
WORKSHOP = dt.date(2026, 12, 1); assert WORKSHOP.weekday() == 1          # Tuesday
assert dt.date(2026, 12, 4).weekday() == 4 and dt.date(2027, 1, 11).weekday() == 0 and dt.date(2027, 1, 29).weekday() == 4


# ------------------------------------------------------------------ shared pieces
def fit_h(s, pt, w, bold=False, font=None):
    """Height for text measured at 95% of the width: room to spare if PowerPoint wraps a word earlier."""
    return kit.text_height(kit.wrapped(s, pt, w * 0.95, bold, font), pt, w * 0.95, bold, font)


def hills(s, y, amp, waves, fill, phase=0.0, x0=0, x1=None):
    """A band whose top edge rolls like hills, filled to the bottom of the slide."""
    x1 = W if x1 is None else x1
    fb = s.shapes.build_freeform(Inches(x0), Inches(H), scale=1.0)
    n = 90
    pts = [(Inches(x0 + (x1 - x0) * i / n), Inches(y + amp * math.sin(phase + 2 * math.pi * waves * i / n))) for i in range(n + 1)]
    pts.append((Inches(x1), Inches(H)))
    fb.add_line_segments(pts, close=True)
    shp = fb.convert_to_shape(); shp.fill.solid(); shp.fill.fore_color.rgb = fill; shp.line.fill.background()
    shp.name = "Hills"
    return shp


def footer_band(s):
    """The motif scaled down: a shallow two-layer hills band holding the logos and the page number."""
    hills(s, 6.82, 0.05, 1.3, MEADOW, phase=0.4)
    hills(s, 6.9, 0.05, 1.7, HILL, phase=2.1)


def lockup(s, x=M, y_mid=FOOT_MID, pt=13, dark=False):
    """Client first, in its serif; the divider sits exactly midway (kit.lockup measures both names)."""
    return kit.lockup(s, x, y_mid, [("Brightfield", pt, True, WHITE if dark else DEEP, "Cambria"),
                                    ("Tandem Studio", pt, True, WHITE if dark else INK, None)],
                      divider_color=MEADOW if dark else kit.tint(DEEP, 0.55))


def slide(n, kicker, headline, source=None, dark=False):
    s = kit.new_slide()
    footer_band(s)
    kit.frame(s, n, kicker, headline, source=source)
    lockup(s)
    kit.transition(s)
    return s


def wave_panel(s, xe, amp, waves, fill, phase=0.0):
    """Section-divider panel from the left edge whose right edge is a vertical wave (the motif as an edge)."""
    fb = s.shapes.build_freeform(Inches(0), Inches(0), scale=1.0)
    n = 90
    pts = [(Inches(xe + amp * math.sin(phase + 2 * math.pi * waves * i / n)), Inches(H * i / n)) for i in range(n + 1)]
    pts.append((Inches(0), Inches(H)))
    fb.add_line_segments(pts, close=True)
    shp = fb.convert_to_shape(); shp.fill.solid(); shp.fill.fore_color.rgb = fill; shp.line.fill.background()
    shp.name = "Divider panel"
    return shp


def divider(num, title, sub, photo_label):
    s = kit.new_slide()
    kit.photo(s, 5.6, 0, W - 5.6, H, label=f"Photo: {photo_label}", label_at="top")
    wave_panel(s, 6.6, 0.22, 1.1, MEADOW, phase=1.2)
    wave_panel(s, 6.35, 0.22, 1.1, DEEP, phase=1.2)
    kit.text(s, M, 1.7, 5.4, 1.5, num, 96, SUN, bold=True, font="Cambria", smart=False)
    kit.text(s, M, 3.25, 5.4, 0.8, title, 40, WHITE, bold=True, font="Cambria", smart=False)
    kit.text(s, M, 4.15, 5.2, 0.9, sub, 18, HILL)
    lockup(s, dark=True)
    kit.transition(s)
    return s


def items_h(items, pt, w, gap):
    return sum(fit_h(t, pt, w - 0.28) for t in items) + gap / 72 * (len(items) - 1)


# ================================================================== 1 title
s = kit.new_slide()
kit.photo(s, 0, 0, W, H, label="Photo: a young family at a sunny farm-shop counter, from the right", label_at="top")
kit.overlay(s, 0, 0, W, H, str(DEEP), mode="gradient", opacity=0.94, angle=0, stops=[(0, 0.94), (0.5, 0.8), (0.8, 0.15), (1, 0)])
hills(s, 6.55, 0.12, 1.2, MEADOW, phase=0.4)
hills(s, 6.8, 0.1, 1.6, HILL, phase=2.1)
lockup(s, M, 0.85, pt=20, dark=True)                         # clearly larger, part of the title composition
kit.text(s, M, 2.0, 7.2, 0.35, "BRAND WORKSHOP  |  1 DECEMBER 2026", 14, SUN, bold=True, smart=False)
kit.text(s, M, 2.45, 7.0, 1.6, "Which brand should take Brightfield's plant range to a new generation?", 36, WHITE, bold=True, font="Cambria")
kit.text(s, M, 4.3, 6.6, 0.8, "A morning to compare two concepts and choose one to take into consumer testing", 18, HILL)
kit.transition(s)

# ================================================================== 2 agenda over a photo
s = kit.new_slide()
kit.photo(s, 0, 0, W, H, label="Photo: oat field at golden hour, horizon on the right", label_at="top")
kit.overlay(s, 0, 0, W, H, str(DEEP), mode="gradient", opacity=0.95, angle=0, stops=[(0, 0.95), (0.55, 0.88), (0.85, 0.2), (1, 0)])
footer_band(s)
kit.frame(s, 2, None, "We'll spend most of the morning on the two concepts, then choose one of them together before lunch", head_color=WHITE)
kit.text(s, M, T["head_top"] - 0.42, CW, 0.3, "AGENDA", 13, MEADOW, bold=True, anchor=MSO_ANCHOR.BOTTOM, smart=False)
lockup(s)
agenda = [("9:30", "Welcome", "what we're deciding today"), ("9:45", "The challenge", "what we've learned about under-35s"),
          ("10:15", "Concept A: Fieldwork", "the idea, and who it wins"), ("10:45", "Concept B: Sprout & Co.", "the idea, and who it wins"),
          ("11:15", "Break", ""), ("11:30", "Scoring and choosing", "both concepts on five criteria"),
          ("12:15", "Next steps", "what testing has to prove")]
rh = (BOT - TOP) / len(agenda)
for i, (tm, what, note) in enumerate(agenda):
    y = TOP + i * rh
    kit.text(s, M, y, 1.0, rh, tm, 18, SUN, bold=True, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    para = [(what, True, WHITE)] + ([("   " + note, False, HILL)] if note else [])
    kit.text(s, M + 1.2, y, 6.6, rh, [para], 18, anchor=MSO_ANCHOR.MIDDLE)
kit.transition(s)

# ================================================================== 3 what we're deciding
s = slide(3, "Today", "By 12:30 we'll have chosen the one concept we take into consumer testing with shoppers in January")
tiles = [("target", "Agree what the brand must do", "Win under-35s without losing the shoppers who trust Brightfield today"),
         ("scale", "Compare both concepts fairly", "The same two slides for each, then the same five criteria"),
         ("flag", "Choose one, and say what to test", "Agree what testing has to prove before we commit")]
strip_h = 0.7; gap_s = 0.3
tg = 0.35; tw = (CW - 2 * tg) / 3; th = BOT - TOP - strip_h - gap_s
for i, (ic, head, body) in enumerate(tiles):
    x = M + i * (tw + tg)
    kit.box(s, x, TOP, tw, th, PALE)
    kit.icon(s, ic, x + 0.35, TOP + 0.3, 0.55, circle=DEEP)
    hh = fit_h(head, 20, tw - 0.7, True)
    kit.text(s, x + 0.35, TOP + 1.05, tw - 0.7, hh + 0.05, head, 20, DEEP, bold=True)
    kit.text(s, x + 0.35, TOP + 1.05 + hh + 0.15, tw - 0.7, th - 1.35 - hh, body, 16, INK)
sy = BOT - strip_h
kit.box(s, M, sy, CW, strip_h, OAT)
kit.text(s, M + 0.35, sy, CW - 0.7, strip_h, [[("Not today: ", True, DEEP), ("the final wording of the name, packaging design, or price", False, INK)]],
         16, anchor=MSO_ANCHOR.MIDDLE)

# ================================================================== 4 divider
divider("01", "The challenge", "Why Brightfield needs a new brand for its plant-based range", "rows of oats, close up")

# ================================================================== 5 the challenge (chart)
s = slide(5, "The challenge", "Under-35s see Brightfield as traditional, and only 18% of them think it's for people like them",
          source="Note: Brightfield shopper survey, September 2026, 1,200 UK adults. Market growth: retail panel, year to August 2026.")
cats = ["Traditional", "Trustworthy", "Modern", "For people like me"]
cx, cw = M, 7.5
kit.chart_title(s, cx, TOP, cw - 2.2, "Shoppers who describe Brightfield as…", left_axis_title=False)
# key: small, above the top-right of the chart
kx = cx + cw - 2.35
for j, (lab, col) in enumerate([("Under-35s", SUN), ("Over-55s", LIGHT)]):
    xx = kx + j * 1.2
    kit.shape(s, MSO_SHAPE.RECTANGLE, xx, TOP + 0.1, 0.16, 0.16, col)
    kit.text(s, xx + 0.24, TOP, 0.95, 0.36, lab, 12, GREY, anchor=MSO_ANCHOR.MIDDLE, smart=False)
cd = CategoryChartData(); cd.categories = cats
cd.add_series("Under-35s", under); cd.add_series("Over-55s", over)
ch_top = TOP + 0.55; ch_h = BOT - 0.3 - ch_top
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(cx), Inches(ch_top), Inches(cw), Inches(ch_h), cd)
ch = gf.chart; kit.style_chart(ch); kit.transparent_chart(ch)
kit.plot_layout(ch, 2.0 / cw, 0.02, (cw - 2.3) / cw, 1 - 0.45 / ch_h)
ch.category_axis.reverse_order = True
kit.axis(ch.category_axis)
va = ch.value_axis; va.minimum_scale, va.maximum_scale, va.major_unit = 0, 100, 25
va.tick_labels.number_format = '0"%"'; va.tick_labels.number_format_is_linked = False
va.has_major_gridlines = False
kit.axis(va, title="% of age group")
pl = ch.plots[0]; pl.gap_width = 55; pl.overlap = -8
pl.has_data_labels = True
dl = pl.data_labels; dl.number_format = '0"%"'; dl.number_format_is_linked = False; dl.position = XL_LABEL_POSITION.OUTSIDE_END
dl.font.size = Pt(14); dl.font.color.rgb = INK
for ser, col in zip(pl.series, [SUN, LIGHT]):
    ser.format.fill.solid(); ser.format.fill.fore_color.rgb = col
px = cx + cw + 0.5
kit.side_points(s, px, W - M - px, ch_top + 0.05, BOT - 0.75, [
    ("Plant-based yoghurt sales grew 18%", "last year, mostly among under-35s"),
    ("Only 18% of under-35s", "say Brightfield is 'for people like me', against 63% of over-55s"),
    ("Half of under-35s still trust it,", "so the name is an asset to use carefully"),
], 16, lead_color=DEEP)

# ================================================================== 6 persona
s = slide(6, "The challenge", "Maya, 29, wants food that's good for her and for the planet, but doesn't want a lecture with it")
pw = 4.1
kit.box(s, M, TOP, pw, BOT - TOP, OAT)
d = 1.55
kit.photo(s, M + 0.35, TOP + 0.35, d, d, label="Photo: Maya (a stock portrait)", crop="circle")
kit.text(s, M + 0.35 + d + 0.25, TOP + 0.55, pw - d - 0.95, 0.5, "Maya, 29", 26, DEEP, bold=True, font="Cambria", smart=False)
kit.text(s, M + 0.35 + d + 0.25, TOP + 1.1, pw - d - 0.95, 0.75, "Bristol; shops at Co-op and Aldi", 14, GREY)
quote = "“I want food that's good for me and the planet. I just don't want to be preached at.”"
kit.text(s, M + 0.35, TOP + d + 0.65, pw - 0.7, BOT - TOP - d - 0.9, quote, 20, DEEP, font="Cambria")
cols = [("shopping-bag", "What she buys", ["Plant-based milk and yoghurt three times a week", "Own-label first, brands when on offer"]),
        ("heart", "What she wants", ["Food that tastes good first", "Brands with a sense of humour", "Proof, not promises, about farming"]),
        ("triangle-alert", "What puts her off", ["Packs that look like a health shop", "Guilt-trip messages about the planet", "Brands her parents buy"])]
cx0 = M + pw + 0.45; cg = 0.35; cwid = (W - M - cx0 - 2 * cg) / 3
for i, (ic, head, items) in enumerate(cols):
    x = cx0 + i * (cwid + cg)
    kit.icon(s, ic, x, TOP + 0.3, 0.55, circle=DEEP)
    kit.text(s, x, TOP + 1.05, cwid, 0.4, head, 18, DEEP, bold=True, smart=False)
    kit.line(s, x, TOP + 1.55, x + cwid, TOP + 1.55, LIGHT, 0.75)
    kit.bullet_list(s, x, TOP + 1.75, cwid, BOT - TOP - 1.9, items, pt=16, color=INK, bullet_color=MEADOW, gap_pt=10)

# ================================================================== 7 divider
divider("02", "Two concepts", "Fieldwork and Sprout & Co., each shown on the same two slides", "a market stall of fresh produce")


# ------------------------------------------------------------------ the two concepts: identical layouts
def concept_idea(n, letter, name, tagline, idea, pillars, moods, swatches, headline):
    s = slide(n, "Two concepts", headline)
    lw = 5.9
    kit.text(s, M, TOP, lw, 0.3, f"CONCEPT {letter}", 13, GREY, bold=True, smart=False)
    kit.text(s, M, TOP + 0.3, lw, 0.65, name, 36, DEEP, bold=True, font="Cambria", smart=False)
    kit.text(s, M, TOP + 0.95, lw, 0.4, tagline, 18, GREY, font="Cambria")
    ih = fit_h(idea, 16, lw)
    kit.text(s, M, TOP + 1.45, lw, ih + 0.05, idea, 16, INK)
    # pillar rows sit at the same, fixed spacing on both concept slides (parallel slides share one layout),
    # measured up from the bottom; the idea paragraph must fit above them
    ph = 0.52
    py = BOT - len(pillars) * ph
    assert TOP + 1.45 + ih + 0.15 <= py, f"idea for {name} is too long for the shared layout: shorten it"
    for k, (ic, lead, rest) in enumerate(pillars):
        y = py + k * ph
        kit.icon(s, ic, M, y + (ph - 0.42) / 2, 0.42, circle=MEADOW)
        kit.text(s, M + 0.62, y, lw - 0.62, ph, [[(lead + ": ", True, DEEP), (rest, False, INK)]], 16, anchor=MSO_ANCHOR.MIDDLE)
    mx = M + lw + 0.5; mw = W - M - mx; g = 0.12
    sw_h = 0.75
    mh = BOT - TOP - sw_h - 0.15
    cw2, ch2 = (mw - g) / 2, (mh - g) / 2
    for k, lab in enumerate(moods):
        kit.photo(s, mx + (k % 2) * (cw2 + g), TOP + (k // 2) * (ch2 + g), cw2, ch2, label=f"Photo: {lab}")
    sw = (mw - 3 * g) / 4
    for k, (nm, hx) in enumerate(swatches):
        x = mx + k * (sw + g)
        kit.box(s, x, BOT - sw_h, sw, 0.4, rgb(hx))
        kit.text(s, x, BOT - sw_h + 0.45, sw, 0.3, nm, 12, GREY, align=PP_ALIGN.CENTER, smart=False)
    return s


def concept_fit(n, wins, risks, test, headline):
    s = slide(n, "Two concepts", headline)
    strip_h = 0.75; gap = 0.35
    cg = 0.4; cwid = (CW - cg) / 2; ch_ = BOT - TOP - strip_h - gap; hh = 0.65
    for k, (ic, head, items) in enumerate([("thumbs-up", "Who it wins", wins), ("triangle-alert", "What it risks", risks)]):
        x = M + k * (cwid + cg)
        kit.box(s, x, TOP, cwid, hh, DEEP, corners="top")
        kit.icon(s, ic, x + 0.3, TOP + (hh - 0.38) / 2, 0.38, color=WHITE)
        kit.text(s, x + 0.85, TOP, cwid - 1.1, hh, head, 20, WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE, smart=False)
        kit.box(s, x, TOP + hh, cwid, ch_ - hh, PALE, corners="bottom")
        bh = items_h(items, 16, cwid - 0.7, 14)
        kit.bullet_list(s, x + 0.35, TOP + hh + (ch_ - hh - bh) / 2, cwid - 0.7, bh + 0.15, items, pt=16, color=INK, bullet_color=MEADOW, gap_pt=14)
    sy = BOT - strip_h
    kit.box(s, M, sy, CW, strip_h, OAT)
    kit.icon(s, "flag", M + 0.3, sy + (strip_h - 0.38) / 2, 0.38, color=DEEP)
    kit.text(s, M + 0.85, sy, CW - 1.15, strip_h, [[("In testing we'd check ", True, DEEP), (test, False, INK)]], 16, anchor=MSO_ANCHOR.MIDDLE)
    return s


concept_idea(8, "A", "Fieldwork", "Grown by people who know the land",
             "Plant-based food that looks and sounds like it comes from a working farm, under the Brightfield name.",
             [("leaf", "Honest ingredients", "British oats and peas"), ("users", "Real growers", "farmers named on every pack"),
              ("shield-check", "A trusted name", "'from the Brightfield family'")],
             ["a farmer in an oat field", "hands holding fresh peas", "a farmhouse breakfast", "sunrise over farmland"],
             [("Wheat", "E3C68A"), ("Soil", "6B4F3A"), ("Sky", "9CC5D9"), ("Leaf", "5E8C4A")],
             "Fieldwork makes plant-based feel honest by leaning on Brightfield's roots as a working family farm")
concept_fit(9, ["Over-35s who already trust Brightfield", "Lapsed dairy shoppers moving to plant-based", "Farm shops and independent grocers"],
            ["Under-35s may see it as more of the same", "Earthy packs can blend in on a crowded shelf", "Keeps the range tied to dairy for some shoppers"],
            "whether under-35s see it as modern enough to try.",
            "Fieldwork would win back lapsed dairy shoppers fastest, but may not feel new enough to under-35s")
concept_idea(10, "B", "Sprout & Co.", "Good stuff, no lecture",
             "A bright, playful new brand that stands apart from dairy, with jokes rather than guilt about the planet.",
             [("sparkles", "A playful look", "bold colours, hand-drawn type"), ("smile", "A light tone", "humour, not lectures"),
              ("rocket", "Room to grow", "stretches to snacks and drinks")],
             ["friends at brunch", "a colourful street mural", "a smoothie bowl", "cycling in a city park"],
             [("Lime", "C6E04A"), ("Pink", "F28DB2"), ("Violet", "6C4AB6"), ("Sun", "FFD23F")],
             "Sprout & Co. makes plant-based feel fun with a bright new brand that stands apart from dairy")
concept_fit(11, ["Under-35s who don't buy Brightfield today", "Shoppers who choose plant-based for fun, not only health", "Supermarket aisles where colour stands out"],
            ["No borrowed trust: the name starts from zero", "A higher launch cost for a brand-new name", "Could look like a trend that fades"],
            "whether shoppers will trust a new name with no family story.",
            "Sprout & Co. would stand out most with under-35s, but it starts without Brightfield's trust")

# ================================================================== 12 divider
divider("03", "Choosing", "Both concepts scored on the same five criteria", "a family picnic on a hillside")

# ================================================================== 13 scoring grid
s = slide(13, "Choosing", "The two concepts score evenly, so the choice comes down to trust versus standing out")
c1, c2, c3 = 3.9, 2.0, 2.0
c4x = M + c1 + c2 + c3 + 0.3; c4 = W - M - c4x
bar_h = 0.7; bar_y = BOT - bar_h
hh = 0.6                                         # two-line column headers: concept name, then its total
for x, w, nm, tot in [(M + c1, c2, "Fieldwork", 11), (M + c1 + c2, c3, "Sprout & Co.", 11)]:
    kit.text(s, x, TOP, w, hh, [[(nm, True, DEEP)], [(f"{tot} of 15", False, GREY)]], 16, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM, smart=False)
kit.text(s, c4x, TOP, c4, hh, "What decides it", 16, DEEP, bold=True, anchor=MSO_ANCHOR.BOTTOM, smart=False)
ry0 = TOP + hh + 0.12; rows_bot = bar_y - 0.3
rh = (rows_bot - ry0) / len(scores)              # rows spread over all the height there is
DOT = 0.24


def dots(x, w, y, n):
    """Three same-size circles: filled for the score, white with the same outline for the rest."""
    g = 0.12; tot = 3 * DOT + 2 * g; x0 = x + (w - tot) / 2
    for k in range(3):
        kit.shape(s, MSO_SHAPE.OVAL, x0 + k * (DOT + g), y - DOT / 2, DOT, DOT, DEEP if k < n else WHITE, line=DEEP, lw=1.5)


for i, (crit, a, b, why) in enumerate(scores):
    y = ry0 + i * rh
    kit.line(s, M, y, W - M, y, LIGHT, 0.75)
    kit.text(s, M, y, c1 - 0.2, rh, crit, 16, INK, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    dots(M + c1, c2, y + rh / 2, a)
    dots(M + c1 + c2, c3, y + rh / 2, b)
    kit.text(s, c4x, y, c4, rh, why, 14, GREY, anchor=MSO_ANCHOR.MIDDLE)
kit.line(s, M, rows_bot, W - M, rows_bot, LIGHT, 0.75)
kit.box(s, M, bar_y, CW, bar_h, DEEP)
kit.text(s, M + 0.35, bar_y, CW - 0.7, bar_h, "To decide: for this range, which matters more, Brightfield's trust or standing out on the shelf?",
         18, WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)

# ================================================================== 14 close
s = kit.new_slide()
kit.photo(s, 0, 0, W, H, label="Photo: friends sharing breakfast outdoors, from the right", label_at="top")
kit.overlay(s, 0, 0, W, H, str(DEEP), mode="gradient", opacity=0.94, angle=0, stops=[(0, 0.94), (0.5, 0.82), (0.8, 0.15), (1, 0)])
hills(s, 6.55, 0.12, 1.2, MEADOW, phase=0.4)
hills(s, 6.8, 0.1, 1.6, HILL, phase=2.1)
lockup(s, M, 0.85, pt=20, dark=True)
kit.text(s, M, 1.55, 7.2, 0.35, "NEXT STEPS", 14, SUN, bold=True, smart=False)
kit.text(s, M, 2.0, 7.0, 1.3, "We'll test the chosen concept with 600 shoppers in January", 32, WHITE, bold=True, font="Cambria")
steps = [("4 Dec", "Write-up and a brief for the chosen concept"), ("11 Jan", "Online test with 600 shoppers aged 18 to 45"),
         ("29 Jan", "Results, and a go or no-go for the launch plan")]
for i, (dte, what) in enumerate(steps):
    y = 3.55 + i * 0.55
    kit.text(s, M, y, 1.2, 0.45, dte, 18, SUN, bold=True, anchor=MSO_ANCHOR.MIDDLE, smart=False)
    kit.text(s, M + 1.3, y, 5.8, 0.45, what, 18, WHITE, anchor=MSO_ANCHOR.MIDDLE)
kit.text(s, M, 5.4, 7.0, 0.65, [[("Priya Nair, Tandem Studio", True, WHITE)], [("priya.nair@tandem.example", False, HILL)]], 16)
kit.transition(s)

print(kit.save(str(Path(__file__).with_name("brand-workshop.pptx"))))
