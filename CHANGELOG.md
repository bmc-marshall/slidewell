# Changelog

## 1.0.2 — November 2026

From a real assignment: rebuilding two slides and adding three to an existing deck.

- **Revise mode asks first.** Before redesigning or adding slides to an existing deck, a short intake: audience and use, one version or three options, whether to keep the existing format (if not, proposed format changes are signed off and applied to the existing slides too), and how hands-on to be. The "options?" question moves into intake Round 1, so the short route asks it too.
- **Take things off.** After vertical logic, a deliberate pass to remove what supports the headline but isn't needed: a strap repeating the headline, the same numbers as both chart and table, comparisons outside the decision's scope, context aimed at a different audience.
- **Supplied text is used word for word** (storyboard bullets, quotes, labels); suggested edits go in the reply.
- **New slides match the deck:** the deck's own text sizes and headline formatting where they meet the standard; containers sized to their content; replacement versions go on the template's placeholder slides. `kit.fill(..., like=headline)` copies an existing headline's formatting.
- **Headlines read the emphasis in the material:** highlighted rows and circled points are the author saying what the slide is for.
- **Sources:** never invent one to satisfy the checker; write "Source: [to be confirmed]" and flag it. Units that look implausible are flagged, not corrected; derived figures are labelled.
- **Charts:** chart titles and panel headers describe, never conclude; stacked legends are a vertical list in stack order, with text centred on each key (`kit.legend`); totals sit above stacked columns, never below the axis; bars start exactly on the baseline (the old wording could be read as leaving a gap); column + line on two visible axes (`kit.combo_chart`) with guidance on when not to use it.
- **Checker:** data labels with K/M/bn suffixes or decimals on values of 10 or more; a chart and a table of numbers on one slide.
- **Review:** an independent reviewer is required wherever one can be started, and is given the audience and the supplied text; new questions on what can come off and whether supplied wording survived.
- **Missing files:** if only SKILL.md is loaded (some hosts sync nothing else), the skill fetches its references and scripts from this release, asks for the zip, or says the checker didn't run.

## 1.0.1 — November 2026

- Checker: small trend lines inside KPI tiles (line charts under 1.7 in tall with no category axis) may hide their axes; the chart guide explains the exception.
- Fixed a line in the build library that only worked on Python 3.12 and later; tests now run on Python 3.10 and 3.12.
- Showcase slides moved out of the skill to `docs/showcase/`, where they build the README and social-preview images.

## 1.0.0 — November 2026

First public release.

- Five modes: new deck, revise, options, share version, template
- Storyline approval before building; a lighter route for decks of five slides or fewer
- Template mode: builds inside a company's own .pptx/.potx template (layouts, placeholders, theme fonts and colours, page numbers)
- Checker with 63 checks, including sources on data slides and alt text on charts and pictures
- Works on macOS, Windows and Linux: no fontconfig, SVG tools or LibreOffice required (LibreOffice adds the visual review)
- Ten example decks with their build scripts, rebuilt and checked by the test suite
- Installs on claude.ai (zip upload), Claude Code (`/plugin`), and other agents (`npx skills add`); the checker installs with pip
