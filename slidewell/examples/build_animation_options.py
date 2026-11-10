"""Animation demo: the same slide revealed three ways, for the user to compare in PowerPoint's slide show.
A: builds as successive slides (n/N), earlier items faded along their own hue (the skill's default).
B: native PowerPoint animation, each tile fades in on click.
C: native animation, each tile fades in and the previous one dims."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from slidekit import Kit
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

kit = Kit(dict(main="0F4C5C", main_pale="E2EEF0", accent="E36414", ink="1E1E1E", grey="6B7378", light="C9D3D6",
               tag_labels=("Option A: builds as slides (3/3)", "Option B: native fade-in", "Option C: native fade-in + dim", "How to review")))
T, CW, M, W, H = kit.T, kit.CW, kit.T["M"], kit.T["W"], kit.T["H"]
HEAD = "A demand-led timetable lets our staff support ~20% more visitors"
TILES = [("Demand has moved to evenings", "62% of visits now happen after 4pm, up from 35% in 2019"),
         ("A new timetable fits staff to visitors", "Moving two daytime shifts into the evening covers the peak in full"),
         ("Two decisions for the board", "Approve a six-month pilot from March; agree three measures of success")]
TW, TH = 3.6, 3.0


def tiles(s, shown, faded=()):
    """Draw tiles 0..shown-1; tiles in `faded` are muted along their own hue. Returns the shapes per tile."""
    xs = kit.distribute_x([TW] * 3, M, W - M)
    top = kit.place_balanced(TH)
    groups = []
    for i in range(shown):
        on = i not in faded
        g = [kit.shape(s, MSO_SHAPE.RECTANGLE, xs[i], top, TW, TH, T["main_pale"] if on else T["muted_fill"])]
        g.append(kit.text(s, xs[i] + 0.3, top + 0.35, TW - 0.6, 0.6, str(i + 1), 32, T["main"] if on else T["muted_main"], bold=True, smart=False))
        g.append(kit.text(s, xs[i] + 0.3, top + 1.05, TW - 0.6, 0.9, TILES[i][0], 20, T["main"] if on else T["muted_main"], bold=True))
        g.append(kit.text(s, xs[i] + 0.3, top + 1.95, TW - 0.6, 0.9, TILES[i][1], 16, T["ink"] if on else T["muted_ink"]))
        groups.append(g)
    return groups


def slide(n, tag, headline=HEAD):
    s = kit.new_slide()
    kit.frame(s, n, tag, headline)
    return s


# 1 title
s = kit.new_slide()
kit.shape(s, MSO_SHAPE.RECTANGLE, 0, 0, W, H, T["main"])
kit.text(s, M, 2.6, 10, 0.8, "Animation options", 40, T["white"], bold=True, smart=False)
kit.text(s, M, 3.5, 10, 0.5, "The same slide, revealed three ways. Open in PowerPoint and use Slide Show to compare.", 20, T["main_pale"])

# 2-4 option A: builds as slides
for k in range(1, 4):
    s = slide(k + 1, f"Option A: builds as slides ({k}/3)")
    tiles(s, k, faded=range(k - 1))

# 5 option B: native fade in
s = slide(5, "Option B: native fade-in")
g = tiles(s, 3)
kit.animate_on_click(s, g)

# 6 option C: native fade in + dim previous
s = slide(6, "Option C: native fade-in + dim")
g = tiles(s, 3)
kit.animate_on_click(s, g, dim_previous=True)

# 7 how to review
s = slide(7, "How to review", "Each option suits a different way the deck will be used")
rows = [("Option A: builds as slides", "Works everywhere: slide show, PDF, print, and when someone flicks through the file. Each step is a separate slide with a counter (1/3), and earlier items fade along their own colour. This is the skill's default."),
        ("Option B: native fade-in", "One slide; each tile fades in on a click. Smoother when presenting live, but a PDF or print-out shows only the finished slide, and a reader scrolling the file sees everything at once."),
        ("Option C: native fade-in + dim", "As B, but the previous tile dims as the next one appears, so attention moves with you. The most polished live, and the most fragile: test it in PowerPoint before presenting.")]
rh = 1.25
top = kit.place_balanced(3 * rh - 0.2)
for i, (a, b) in enumerate(rows):
    y = top + i * rh
    kit.text(s, M, y, 3.6, 1.0, a, 18, T["main"], bold=True)
    kit.text(s, M + 3.9, y, CW - 3.9, 1.0, b, 16, T["ink"])

print(kit.save(str(Path(__file__).with_name("animation-options.pptx"))))
