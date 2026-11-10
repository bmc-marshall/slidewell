# QA: working method, procedure, and checklists

**Covers:** how to work on any build or revision so quality doesn't slip, the QA loop that gates every delivery, the accuracy pass for numbers, and the per-slide and deck-level checklists.

**Doesn't cover:** the rules themselves. The checklists point to `finishing-standard.md` and `charts.md`, which hold each rule and its reason.

QA is a delivery gate, not an optional review. Never describe a slide as presentation-ready, QA'd, or compliant unless this full procedure has actually been run on it.

## 1. Working method

These exist because they are exactly where AI-built slides have failed before.

- **Corrections are cumulative.** Every requirement and every accepted fix becomes a locked constraint. Change only what the new request needs; never reintroduce a fixed issue. Before delivering any revision, re-check the whole constraints log, not just the latest request. Most regressions come from treating a fix as local.
- **Inspect before editing.** Look at the target slide *and* the slides that define the deck's visual language. Preserve recurring elements exactly unless asked to change them.
- **Copy, don't approximate.** Use established elements (from the template or earlier slides) as formatting templates: copy their geometry and formatting rather than recreating them. Don't introduce a new font size, padding, alignment, corner, bevel, border, or fill unless it serves a deliberate hierarchy purpose. Recreated boxes are where inconsistent bevels come from.
- **Judge the render, not the coordinates.** An element being programmatically placed doesn't make it visually correct. Render and look at normal presentation scale.
- **Compare, don't just check.** Where elements are parallel (cards, columns, labels, boxes), compare them against each other directly. Where sections sit side by side, evaluate them as one composition, not column by column; this is how uneven bottoms and unbalanced margins get caught.

## 2. The QA loop

1. **Automated check:** `python scripts/check_deck.py deck.pptx` (add `--theme theme.json` if you wrote one). Fix every error; review every warning. A warning you decide to keep needs a reason (e.g., the user asked for two slides to share positions even though one is then slightly off-balance).
2. **Render:** `python scripts/render.py deck.pptx` (use `--slides` to re-render only what changed). Open **every slide image on its own at full size** (the default 150 dpi, or at least 100) and inspect it; the contact sheet is only for checking consistency across slides, never a substitute for looking at each slide. Don't deliver after the first render. **A clean checker run is necessary, not sufficient:** the checker can't see a label colliding with a placeholder, a price pressed against the bottom of its card, a chart floating in empty space, or a deck that looks like every other deck. Those are found only by looking.
3. **Hunt for failures:** go through the per-slide checklist below, **actively looking for a failure** rather than confirming the slide looks fine. Inspect every object, including ones not changed in the latest revision. Compare parallel elements side by side, and evaluate side-by-side sections as one composition.
4. **Accuracy pass:** run `python scripts/check_deck.py deck.pptx --numbers` to list every number on every slide (text and chart data). Re-derive every figure that is computed from others, in code, not by eye: sums and differences, percentages and shares, growth rates, ratios, durations between times, totals of agenda slots, payback periods, and any claim like "more than half" or "twice as fast". Put the checks as `assert`s in the build script so a later edit can't silently break them. Compare decimals with rounding (`round(a + b, 6) == c`) or `math.isclose`, since 0.9 + 0.5 + 0.4 isn't exactly 1.8 in floating point. Also check that each derived number is unambiguous on the slide: a duration label states the range it covers ("10:45–11:25, 40 minutes"), and a highlight box covers exactly the items its label counts. Express durations in natural units: "2 years, 8 months", not "2.7 years" (only halves may be fractional, written "1½").
5. **Fix and restart:** if anything fails, or differs without an intentional reason, fix it, re-render, and restart the checklist **for the whole slide**. Repeat until every applicable item passes.
6. **Regressions:** re-check the constraints log; after a revision, run `scripts/compare_versions.py` (SKILL.md mode B).
7. **Deck-level checklist** before final delivery.

Remember that renders here may substitute fonts; distinguish real issues from substitution artifacts by checking the text box size against the content, not just the picture.

## 2b. Review rounds (the internal reviewer)

After the QA loop passes, run review rounds: an internal version of a demanding human reviewer going through the deck slide by slide. It catches what checkers can't: things that look off, crowded, generic, or inconsistent.

1. **Reviewer.** Where you can start a separate agent (a subagent), do: give it only the rendered slide images (full size, one per slide), `references/review-brief.md`, `finishing-standard.md`, and the approved storyline; not the build script or your reasoning. A reviewer that didn't build the deck isn't marking its own work. Where you can't, do the review yourself as a separate pass: re-read the brief first, then look at each slide image fresh, as if someone else made it.
2. **Findings.** The reviewer returns a numbered list, worst first: slide, what's wrong, which rule, how serious (*must fix*: breaks a rule or looks wrong; *should fix*: noticeably weaker; *polish*). It must look at every slide individually and say "no issues" for a slide only after checking it against the brief.
3. **Fix.** Fix every must-fix and should-fix item, re-render the changed slides, re-run `check_deck.py` (a fix can break something else), and update the constraints log.
4. **Repeat** with a new review of the changed slides and their neighbours, until a round finds no must-fix or should-fix items, or after three rounds. Don't loop forever: after three rounds, list what's still open in the QA report.
5. **Record** each round (findings and fixes) for the QA report at delivery.

## 3. Per-slide checklist

## 3. Per-slide checklist

1. Spelling and terminology
2. Recurring elements preserved (position and formatting)
3. Headline: one message, ≤2 lines, fits the horizontal logic; all content supports it; top line used fully, no 1–2 word second line
4. Clear hierarchy among nearby text, with separated vertical bands
5. Content group distributed between headline and notes, with ≥0.6 in breathing room; clear space between each header and what it heads
6. Consistent margins, gutters, vertical gaps, internal padding; rows of items evenly distributed
7. Balanced left/right weight; parallel sections aligned top and bottom; side panels aligned with the plotted data
8. Outer margins balanced; gutters wide enough to separate sections
9. No overflow, clipping, collisions, or spillover; no orphans or short fragments
10. Container geometry consistent (corners, bevels, borders, fills, padding); one corner style
11. Equivalent elements identical unless deliberately differentiated; one text colour per fill
12. Text visually aligned to the objects it describes; labels centred on their shapes
13. Short callouts vertically centred
14. Font sizes meet the tiers; at most two fonts; no unintended italics/underline/shadow
15. Palette ≤6 brand colours (+ labelled RAG); colours used consistently; topic label distinct from the headline
16. Sources ≥10 pt in the standard position; qualifiers present; page number low in its corner
17. Charts: title, axis line, axis title/units, zero baseline, ticks, light gridlines (`charts.md` §6)
18. Charts: title clearance, label clearance, axis-label proximity, tick alignment, no bar/axis overlap; commentary matches the chart
19. Numbers correct (accuracy pass), rounded sensibly; abbreviations clear; status markers where needed; durations in natural units
20. No invented facts and no over-claiming; no meaningless design machinery
21. No regression from any previously accepted version

## 4. Deck-level checklist

1. Leads with the key message: executive summary (findings, recommendation, options), status at a glance (update), or objectives + agenda (workshop)
2. Sections grouped and ordered logically; flow trackable via trackers, breadcrumbs, or dividers; any tracker introduced first and shown on every slide in its span
3. Headlines read in sequence tell a coherent story
4. Clear which slides stand alone and which need a presenter, and what the audience should do with each
5. Recurring elements and inner margins consistent across all slides; elements that do the same job (chart titles, charts, panel headers, decision bars) sit in the same place on every slide. Flip through the renders in order and watch for anything that jumps
6. Charts consistent with each other in style, colours for the same meanings, and label treatment
7. Animated slides (if any) listed with their click sequence in the delivery note
8. Remind the user to review in presentation mode on the machine they'll present from, since fonts and rendering can differ between computers
