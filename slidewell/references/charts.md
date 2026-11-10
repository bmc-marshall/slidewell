# Charts

**Covers:** everything about charts: whether to use one, which one, how to build it natively, the conventions, the finishing details, and the chart checklist.

**Doesn't cover:** general text, colour, and spacing rules that apply to charts as to everything else (`finishing-standard.md`).

Read this whenever a slide contains numbers. AI-built decks tend to fall back on stating numbers in text when a chart would make the point faster and more convincingly. Treat "should this be a chart?" as a question you answer deliberately for every number on every slide.

## Contents
1. Chart or text?
2. Choosing the chart: start from the message
3. Decision guide by relationship
4. Building each chart type in PowerPoint
5. Conventions (titles, axes, labels, numbers, colour, annotation)
5b. Rules and finishing details
6. Chart checklist

---

## 1. Chart or text?

| What you have | Use |
|---|---|
| One number | Text, or a hero number if it's a standalone headline fact |
| Two or more values **being compared** (across items, over time, as parts of a whole, against a target) | **A chart** |
| Several numbers that measure *different* things (KPIs of different metrics, figures supporting different points, a number in a pillar header) | Individual numbers are fine; don't force unrelated measures onto one chart |
| Exact lookup of many values across several dimensions | A table (see finishing-standard §8), with the key comparison also visualized or called out |

Hero-number tiles are for standalone facts that aren't being compared with each other, and should be **rare**. If you find yourself writing "A grew from 12 to 19 while B fell from 15 to 9" in a bullet, that sentence is a chart.

## 2. Choosing the chart: start from the message

1. Write the slide headline first. The chart's job is to prove it.
2. Name the **relationship** the headline is about: comparison among items, change over time, composition (parts of a whole), distribution, relationship between variables, ranking, deviation from a reference, flow, or geography. The same data can need different charts depending on the message (e.g., monthly revenue by plan: a line chart if the point is one plan's growth; a bar chart if the point is which plan earns most).
3. Decide the level of detail: granular (exact values, outliers matter) or aggregated (the overall direction matters).
4. Pick the simplest chart in §3 that shows that relationship.

## 3. Decision guide by relationship

Adapted from A. Abela's *Chart Suggestions: A Thought-Starter* (2006) and Flourish's chart-type guidance, with this skill's own conventions applied (no pie charts, no 3D).

**Comparison among items** (sizes, magnitudes, rankings)
- One value per item, few items → **column** chart.
- One value per item, many items or long names → **horizontal bar** chart, sorted by value.
- Two values per item → **clustered bar/column**; or a **dumbbell** (two dots joined by a line) when the gap between the two values is the point.
- Many items × many categories, needs lookup → **table**, optionally with embedded mini-bars or shaded cells (heatmap).
- Ranking where position is the point → sorted **bar**; a **lollipop** draws more attention to individual values.
- Item size on two dimensions (e.g., share and growth) → **variable-width (Marimekko) column** (build with shapes).

**Change over time**
- Few periods, one or few series → **column** chart.
- Few periods, many series → **line** chart (or small multiples).
- Many periods → **line** chart.
- Two points in time, several items → **slope** chart.
- Projections with uncertainty → line with a **fan** (shaded range) or **scenario lines** (see §5 forecasts).
- Cyclical data (e.g., months of the year) → line chart by cycle; radar charts should rarely be needed.

**Composition (parts of a whole)**
- Static, simple share of a total → **100% stacked bar** (a single bar works well), or a sorted bar of shares. **Never a pie.**
- Static, accumulation or subtraction to a total (bridges, budgets, savings) → **waterfall**.
- Static, components of components → **100% stacked column with sub-components**, or a **treemap** (build as image or shapes).
- Changing over time, few periods → **stacked column** (absolute and relative differences matter) or **100% stacked column** (only shares matter).
- Changing over time, many periods → **line** chart of each component; stacked area only when the total matters and series are few.

**Distribution**
- One variable, few data points → **column histogram** (zero-gap columns).
- One variable, many data points → **line histogram** / density curve, or a **box/violin**-style range chart built with shapes.
- Emphasizing individual points → **strip/beeswarm** (dots).
- Two variables → **scatter** plot.
- Age/sex or two-sided distributions → **population pyramid** (back-to-back bars).

**Relationship between variables**
- Two variables → **scatter** plot (add a trend line only if it's meaningful).
- Three variables → **bubble** chart (bubble area = third variable; label key bubbles).

**Deviation from a reference** (target, average, zero, last year)
- **Diverging bar** chart (bars left/right of the reference line); **diverging stacked bar** for survey agreement scales; a reference line on a bar or line chart.

**Flow and connection**
- Volumes moving between stages or states → **Sankey** (image from a specialist tool, native annotations).
- Connections between entities → **network** diagram; for simple cases build with shapes.
- Process steps → see layout-patterns §E (not a chart).

**Geography**
- Only when the geographic pattern matters more than anything else; having a location column is not a reason to map.
- Totals by place → **proportional symbol** map; rates or results by region → **choropleth**. Usually inserted as an image with native title, annotations, and source.

## 4. Building each chart type in PowerPoint

Prefer native, editable charts (python-pptx `chart_data` + `add_chart`):

| Chart | How |
|---|---|
| Bar, column, clustered, stacked, 100% stacked | Native (`BAR_CLUSTERED`, `COLUMN_STACKED`, `BAR_STACKED_100`, …) |
| Line, slope, scenario lines | Native `LINE`; slope = two categories |
| Scatter, bubble | Native `XY_SCATTER`, `BUBBLE` |
| Waterfall | Native **stacked column** with an invisible (no-fill) base series; or shapes |
| Histogram | Native column with gap width 0–5% |
| Diverging bar | Native bar with negative values and a reference line |
| Population pyramid | Native bar with one series negated and the axis labels formatted as absolute values |
| Dumbbell, lollipop, Marimekko, strip | Shapes (lines + circles/rectangles) from a computed scale |
| Heatmap table | Native table with conditional cell fills |
| Sankey, treemap, network, maps | Image from a specialist tool; title, annotations, and source stay native |

When you build a chart from shapes, compute positions from one scale function so marks, labels, and gridlines can't drift apart.

## 5. Conventions

**Titles and units**
- Chart title is **descriptive**: what, scope, period ("Daily traffic by corridor, 2025"). The slide headline carries the message. Chart titles ≥16 pt. Labels that name the data (category labels, series labels) ≥14 pt, ideally 16; a value axis that only gives the scale (ticks and a units title like "People") can be a step smaller (12 pt) so the data labels lead.
- **Chart title position:** left-aligned with the axis tick labels (the visible edge of the chart), not with the rotated axis title; same position on every chart slide (finishing-standard §5).
- The **unit goes in the value-axis title** ("Vehicles per day", "Revenue ($M)"), not in the chart title.
- Leave clear space between the title and the plot.

**Axes**
- Keep the value axis (line and tick labels) even when every mark has a data label. The one exception is a **sparkline**: a small trend line inside a KPI tile, there to show direction rather than values, has no axes. *(Checked: the checker allows hidden axes only on line charts under 1.7 in tall with no category axis.)*
- Axis line always visible; tick labels close to the axis without touching; edge labels centered on their gridline.
- Value axes start at zero. A non-zero start is allowed on any chart **only if it's clearly labeled** (an axis break or a note).

**Gridlines**
- Usually none. Use thin, light-grey gridlines on the value axis only when the marks have **no** data labels **and** readers need to read specific values or ranges off the axis. If the general trend is enough, leave them out.

**Identifying series**
- Line charts (lines running broadly horizontally) → **direct labels at the right end of each line**, colored to match.
- Stacked bars/columns → **legend to the right of the rightmost bar, each entry vertically aligned with its segment**.
- Everything else → direct labels; a legend to the right of the plot only if direct labels won't fit.
- Keys for tables and diagrams (rating scales, symbols): small (10–12 pt), italic, grey, just above the top-right corner, right-aligned. Better still, label the highlighted thing directly (§5b).
- Data labels inside bars use the centre position. Forecast dividers sit on the last actual value, where dashed lines begin (§5b).

**Category labels (horizontal bars)**
- Left-aligned in a label column. Each label's text box is vertically centered on its bar, **and its text is vertically centered within the box**, so a wrapped label stays centered on the bar. Wrap long labels deliberately rather than shrinking them.

**Data labels**
- Regular bars/columns → **outside end**.
- Stacked bars/columns → **inside each segment** (contrasting text; omit labels on segments too small to hold them and note it).
- High-level charts → label only the key values (e.g., first, last, highlighted).
- Line charts → label the latest value (and other values the headline mentions) with a small dot.

**Numbers**
- Percentages: whole numbers; **1 dp** for values under 5% and for CAGRs/growth rates.
- Large numbers: state the scale once in the axis title ("Revenue ($M)", "Vehicles per day (thousands)") and show plain numbers in labels, without K/M suffixes: 1 dp below 10, no decimals at 10 and above. (User's decision.)
- Be consistent across every chart in the deck.

**Sorting**
- Sort by the value the headline is about, descending, unless there's a natural order (time, bands, scales; "Other" last) or the user specifies the source order.

**Colour**
- Colour only differentiates categories that can't be distinguished otherwise, or highlights the point the headline makes.
- If the commentary highlights something → the focal bar/series in the main or accent color; everything else neutral grey or a light tint.
- If the chart shows all data evenly → brand palette in a fixed order, used the same way on every chart.
- Positive/negative values (waterfalls, variances) → two brand colors that read as a pair (e.g., a brand blue and red); otherwise blue for positive and red for negative, or one brand color plus grey. Totals in the main color. Signs on labels (+/−).
- Never rely on color alone; pair it with labels.

**Levels, thresholds, and gaps**
- Show capacities, targets, and thresholds as **dotted horizontal reference lines** with plain labels at their ends ("Current capacity", "New capacity", "Target"). Don't join two levels with a vertical step line; it reads as a change in the data.
- Show a **gap** (unmet demand, shortfall, overspend) as a **shaded area** between the two lines, in a light tint of the accent, labelled in words ("Unmet demand today").
- Avoid cryptic marks: a bare number on a dashed line asks the reader to work out what it means.

**Annotations**
- Match the deck's style: either short accent-colored text (≤10 words) with a thin leader line or circle, or filled callout boxes with a pointer. About three per chart at most, placed outside the plot where possible, ≥10 pt, and native/editable. A **curved arrow** distinguishes an annotation from the chart's own lines; make it **two-toned** where it crosses backgrounds (dark over the slide, white over a shaded area).

**Forecasts and estimates**
- Use all four signals together: dashed lines or lighter fills for projected periods; "E"/"F" suffix on period labels (2027E) with a footnote; a vertical divider where actuals end; scenario names at line ends (base / high / low).

**Time labels**
- Shortest clear form: full years (2025); months as "Jan" with the year shown at the first month and at year changes; fiscal years "FY25".

**Marks**
- Bars wider than the gaps between them (gap ≈ 50–80% of bar width), consistent across the deck. Bars never touch the axis baseline.
- Lines about 2 pt, no markers except a dot on labeled points; the focal line slightly thicker.

**Multiple charts**
- At most two charts per slide, except special cases (maps, report screenshots, small multiples). Charts of the same measure share scale, size, and alignment.

**Sources and notes**
- In the slide's standard source box (bottom-left): "Note: …" first, then "Source: Organization, Title (year)". Numbered footnotes for specific values.

## 5b. Rules and finishing details

**Rules every chart meets** *(the checker tests several of these)*
- **Chart type** fits the message: name the relationship the headline is about and pick the simplest chart for it (§3). Two or more values being compared get a chart, not text. No pie charts; no 3D.
- **Every chart has:** a descriptive title; units in the value-axis title; a visible axis line with tick labels; a zero-based value axis unless the exception is clearly labelled; sensible tick marks; gridlines only when readers must read values off the axis (thin, light, value axis only).
- **Legend** only if it adds interpretive value; when present, aligned with the chart. Direct labels are often better.
- **Tick labels** align unambiguously to their gridlines or tick marks; edge labels (e.g., the maximum) sit centred on the final gridline, not detached.
- **Label space:** reserve room for category, tick, and data labels. Labels never overlap bars, axes, gridlines, or each other, and never intrude into the plot area. If they spill, reduce label size (≥10 pt), wrap deliberately, widen the label area, or abbreviate clearly.
- **Title clearance:** clear space between the chart title/subtitle and the plot, so the title doesn't crowd the first bar, gridline, label, or axis (about 0.25–0.35 in; finishing-standard §5).
- **Axis proximity:** tick labels and axis titles close enough to their axis that the relationship is obvious, without touching.
- **Bars** never touch or overlap the axis baseline.
- **Native first:** build charts as native, editable PowerPoint charts wherever practical. If a visual can't reasonably be rebuilt, place it as an image and keep its title, annotations, and source native.

**Finishing details** *(from the author's reviews of built decks)*
- **Line labels sit on their lines.** A label at the end of a line (or a reference line) has its text box vertically centred on the line's y-position and its text vertically centred in the box, so the label reads as part of the line.
- **Reference lines span the plot:** capacity, target, and threshold lines run from the value axis to the end of the plot, so they don't suddenly appear partway along. If a level only applies to part of the range, say so in the label ("Evening capacity").
- **Place annotations where they don't clutter:** usually in empty space next to or beneath what they describe, not across the data. An annotation's text uses a colour that doesn't match any data series (usually near-black). *Why:* an annotation in the same colour as a line reads as part of that series.
- **Scatter labels: choose a side for each point, then check for collisions.** Default to the right of the point; put a label on the left, above, or below when its neighbour, the zone label, or the plot edge is in the way. After rendering, check that no label touches another label or point. Use word-wrapped label boxes wide enough for one line, not no-wrap boxes: some renderers re-centre no-wrap boxes, so their labels drift away from their points.
- **Zone labels belong to their zone.** Put the label for a shaded zone ("Target zone") *inside* the zone, in its emptiest corner (often top-right), a size smaller than the data labels, in the same case style as other labels (title or sentence case, or all caps if that's the deck's label style). Only if there's no empty corner, place it just outside, touching the edge. *Why:* a label floating near a box doesn't clearly belong to it.
- **Labels that don't fit inside a bar or segment go outside it,** never broken across lines mid-word: widen the bar, combine small segments ("Other"), shorten the label, or place the label just outside, centred on the segment. *(Checked.)*
- **Labels inside bars are centred:** horizontally and vertically within the bar (use the chart's centre label position, not "inside end"). *Why:* a label pushed to one end looks misplaced, especially on short floating bars like waterfall steps.
- **Forecasts: the divider sits on the last actual value,** and the dashed forecast lines start exactly there. Never let a dashed segment start before the divider. Label the divider "Forecast →" on the forecast side. *Why:* the dash means "estimated"; a dashed line before the divider makes actual data look estimated.
- **Commentary must match the chart.** Side points or notes must talk about what the chart shows. If they add detail (profit when the chart shows cash), that detail belongs on the chart instead. If they explain individual bars, put them as short notes directly under those bars, centred on each one, rather than in a separate panel. *Why:* a chart saying one thing beside bullets about something else makes the reader reconcile two stories.
- **Takeaways beside a chart line up with the plotted data.** The panel header sits level with the chart title. The first point starts level with the top of the plotted data (the 100% line, the first bar), not higher up near the title. Space the points evenly; let the gaps grow to spread the points down the plot, from about one line to about 2.5 lines between points. If even 2.5-line gaps can't fill the plot (only two short points, say), keep the points top-aligned with the data rather than spreading them further. *Why:* the points read as commentary on the data next to them; points crammed at the top of a tall chart leave the panel looking unfinished, and points spread too thin stop reading as a group.
- **Chart titles align with the axis, not the axis title.** The chart title's left edge lines up with the left edge of the axis tick labels (the visible start of the chart), not with the rotated axis title or the invisible chart frame. Use the same x on every chart slide, even one without a left axis title. *Why:* the axis is what the eye reads as the chart's edge; aligning to it makes title and chart read as one unit.
- **Keep chart styles consistent across the deck.** Charts look like one coherent body of work (fonts, colours, label sizes, axis treatment). A second chart about the same subject (payback after a profit bridge) should reuse the first one's style: the same bar types, colours for the same meanings (cost, profit), label treatment. A one-off style (hollow bars, bars hanging below zero) looks like it came from a different deck and makes the reader relearn how to read it.
- **Prefer direct labels to keys.** Say what a highlight means on the highlight itself: e.g., "RECOMMENDED" in small caps, rotated 90° and centred vertically along the left (or right) side of the outlined option; or above the outline, left-aligned, if that space is free. *Why:* a key ("Gold outline = recommended") makes the reader decode; a label just tells them. `slidekit.vertical_label()` builds the rotated version.
- **When a key or legend is needed, keep it small and close:** 10–12 pt, italic, grey, right next to what it explains, usually just above the top-right corner of the table or chart, right-aligned; not below the bottom where it reads like a footnote.

## 6. Chart checklist

Before a chart slide passes QA:
1. Is a chart the right form (§1), and is it the simplest chart for the relationship (§3)?
2. Does the chart prove the headline at a glance? Is the focal element highlighted?
3. Descriptive chart title; unit in the axis title; clear space between title and plot.
4. Axis line visible, zero baseline (or labeled exception), sensible ticks, gridlines only when needed.
5. Series identified by the right method; category labels aligned and vertically centered; data labels positioned correctly and not overlapping anything.
6. Numbers rounded per §5, consistent across the deck; sorted sensibly.
7. Forecasts, estimates, and dummy data clearly marked.
8. Source and notes present, ≥10 pt.
9. Consistent with every other chart in the deck.
