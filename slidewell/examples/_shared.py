"""Small helpers shared by the example build scripts (all content in the examples is fictional)."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from slidekit import Kit, CategoryChartData, XyChartData, rgb  # noqa: F401
from pptx.util import Inches, Pt  # noqa: F401
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR  # noqa: F401
from pptx.enum.shapes import MSO_SHAPE  # noqa: F401
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_TICK_LABEL_POSITION  # noqa: F401

HERE = Path(__file__).resolve().parent
AX = 14 * 1.2 / 72 + 0.12           # chart titles start at the axis labels, not at the rotated axis title


def out(name):
    return str(HERE / name)


def tag(kit, s, label, fill=None):
    """Fixed-size header tag, left-aligned text (finishing-standard §4)."""
    T = kit.T
    kit.shape(s, MSO_SHAPE.RECTANGLE, 0, 0, kit.tag_w, T["tag_h"], fill or T["main"])
    kit.text(s, T["M"], 0, kit.tag_w - T["M"], T["tag_h"], label.upper(), 11, T["white"], bold=True,
             anchor=MSO_ANCHOR.MIDDLE, smart=False)


def bar_chart(kit, s, x, y, w, h, cats, vals, highlight=(), horizontal=False, axis_title=None, maximum=None,
              unit=None, fmt='0', pt_cat=14, pt_val=12, colors=None, gap=60):
    """Bar or column chart in the deck style: grey bars, highlighted bars in the main colour, outside-end labels.
    Returns (plot_top, plot_bottom) in inches."""
    T = kit.T
    cd = CategoryChartData()
    cd.categories = list(reversed(cats)) if horizontal else cats
    cd.add_series("Series", list(reversed(vals)) if horizontal else vals)
    kind = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
    ch = s.shapes.add_chart(kind, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
    kit.style_chart(ch)
    PX, PY = (0.36, 0.02) if horizontal else (0.1, 0.06)
    PW, PH = (0.56, 1 - 0.02 - 0.75 / h) if horizontal else (0.88, 1 - 0.06 - 0.75 / h)
    kit.plot_layout(ch, PX, PY, PW, PH)
    plot = ch.plots[0]; plot.gap_width = gap; plot.vary_by_categories = False
    ser = plot.series[0]; ser.format.fill.solid(); ser.format.fill.fore_color.rgb = T["light"]
    n = len(cats)
    for i in range(n):
        k = (n - 1 - i) if horizontal else i
        c = (colors[i] if colors else (T["main"] if i in highlight else None))
        if c is not None:
            p_ = ser.points[k]; p_.format.fill.solid(); p_.format.fill.fore_color.rgb = c
    dl = ser.data_labels; dl.show_value = True; dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.number_format = fmt; dl.number_format_is_linked = False; dl.font.size = Pt(14)
    va = ch.value_axis; va.has_major_gridlines = False; va.minimum_scale = 0
    if maximum:
        va.maximum_scale = maximum
    kit.axis(va, axis_title, pt=pt_val); kit.axis(ch.category_axis, pt=pt_cat)
    return y + PY * h, y + (PY + PH) * h, ch, (PX, PY, PW, PH)


def chart_slide_layout(kit, s, title, chart_w, ch_h, panel_title=None, title_gap=0.7, source_lines=None):
    """Distribute a chart group (title + chart) between headline and notes; return geometry."""
    T = kit.T
    top = kit.place_balanced(title_gap + ch_h, source_lines=source_lines)
    kit.chart_header(s, T["M"] + AX, top, chart_w - AX, title)
    if panel_title:
        rx = T["M"] + chart_w + 0.5
        kit.chart_header(s, rx, top, T["W"] - T["M"] - rx, panel_title)
    return top, top + title_gap
