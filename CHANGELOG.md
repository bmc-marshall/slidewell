# Changelog

## 1.1.0 — November 2026

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

Getting it ready to publish (10 November 2026):
- Privacy pass: scanned for real client names, cleared author fields in every example deck, kept working notes out of the repo.
- `requirements.txt` with pinned ranges; the visual review reports `VISUAL REVIEW UNAVAILABLE` and the rest of QA carries on when LibreOffice or Poppler is missing.
- Description narrowed to PowerPoint and consulting-style decks; Google Slides, Keynote and HTML slides excluded.
- Template mode built out: `Kit(template=...)` reads a template's colours, fonts, headline position and footer, and builds with its own layouts and placeholders. New example: a steering-group update inside a fictional corporate template (ten examples in all).
- Tests that rebuild and check every example; a GitHub Action that runs them and attaches `slidewell.zip` to each release.
- Claude Code plugin listing and `npx skills add` support.
- A side-by-side with Anthropic's pptx skill on the same brief, with the brief included so anyone can rerun it.
- New checks for sources on data slides and alt text on charts and pictures.
- Time estimate re-measured: a 16-slide proposal took about 11 minutes, including three fix rounds.

---

## Before 1.0: building the skill (reconstructed)

The rounds below happened in conversation before Slidewell was on GitHub, so they have no commits or tags. They're listed in the order they happened, numbered retroactively; dates are approximate (early November 2026). Each round came from building or reviewing real decks and turning the feedback into general rules.

## 0.13.1 — PowerPoint repair fix

- Fixed decks that PowerPoint offered to repair: a line chart hid its markers by adding a second `<c:marker>` by hand. New `kit.no_markers(series)` does it through python-pptx, and the market-entry example uses it.
- Checker: `file-integrity` errors for hand-edited XML that is out of schema order or duplicated (chart series and markers, fills and outlines, bullets, runs, transitions). LibreOffice renders such files without complaint, so renders can't catch this.
- SKILL.md: the file must open cleanly in PowerPoint; edit XML only in schema order.

## 0.13.0 — Time estimates

- Before building, say how long it will take, as a range with a confidence level, and explain any overrun. Calibrated on a 16-slide reading deck built with the skill.

## 0.12.0 — Proposals and five more examples

- New deck type: **proposal** (pitching for work), with its own section-menu column and storyline: the client's goals first, then approach, plan, commercials, team, credentials last. Layout patterns for understanding, phased offers, delivery plans, pricing options, team pages and case studies.
- Five new fictional example decks (consulting proposal, museum turnaround, investment pitch, university club case for action, fundraising workshop): nine in all.
- Checker: words wider than their box, palette counted by colour family (tints don't count), photo bleeds allowed, margins skipped on half-bleed slides, small navigation labels allowed in the top and bottom bands, money to 2 dp allowed.
- slidekit: the orphan fix relaxes the middle-line rule rather than leave an orphan; large photo placeholders are labelled bottom-right.
- Charts: labels that don't fit inside a bar go outside it, or small segments are combined.

## 0.11.0 — Checkpoints and review rounds

- Intake asks how hands-on the user wants to be: review the storyline and the look, the storyline only, or build straight through, with the trade-off of each.
- Review rounds after the QA loop: a reviewer that didn't build the deck looks at full-size slide images against a new `review-brief.md` (built from the author's most frequent feedback), up to three rounds.
- Every delivery includes a QA report: what ran, what was found and fixed, what's still open.

## 0.10.0 — Deck types and template mode

- Deck types as presets: findings, recommendation, options appraisal, update and workshop, each with its own column of pre-ticked slides in the section menu and a deck-types table in `storytelling.md`.
- No agenda in findings, recommendation or options decks (the executive summary is the map); agendas for workshops and updates; icebreaker for workshops.
- New mode E: a styled **template** of the slides the user needs, with bracketed placeholder text; intake asks written deck or template.
- More storyline examples (an investment pitch, a case for action, a steering-group update); principles kept in `storytelling.md`, examples in `storyline-examples.md`.

## 0.9.0 — One rule, one place

- `finishing-standard.md` replaces `design-rules.md` and `craft-notes.md`: every finishing rule once, with its reason and whether the checker tests it. A sentence-by-sentence check confirmed nothing was dropped.
- `qa.md` holds the working method, QA loop, accuracy pass and checklists; `charts.md` holds every chart rule.
- Every reference file opens with what it covers and where to look for anything it doesn't.
- README, MIT licence, and the first four fictional example decks with their build scripts.

## 0.8.0 — Animation

- Animation is an intake question: animated if presented live, not if sent to read. Three styles: builds as successive slides, fade in on click, or fade in and dim earlier items when each supersedes the last. `slidekit.animate_on_click()`.
- Labels on shapes are centred on the shape (checked).
- Storyline examples anonymised.

## 0.7.0 — Market-entry deck round

- One text colour per solid fill across the deck (checked, including labels inside bars).
- Durations in natural units ("2 years, 8 months"; halves as "1½").
- Headlines use as much of the top line as possible; a second line of one or two words is reworded (checked).
- Large numbers: the scale goes in the axis title, not K/M suffixes on labels.

## 0.6.0 — Workshop deck round: "make it cumulative"

- Accuracy pass in QA: `check_deck.py --numbers` lists every number; derived figures are recomputed and asserted in the build script; range labels state their endpoints.
- Breathing room of at least 0.6–0.8 in above and below the content, with the content group vertically balanced (`place_balanced`; checked).
- Page number stays low in its corner; rules and guidelines lists lead with a bold key phrase; chart labels that name data are larger than scale-only axis text.
- Orphan check measures bold and regular words separately.

## 0.5.0 — slidekit

- The build helpers proven on the worked examples move into `scripts/slidekit.py`, so every deck starts with them: phrase-aware wrapping without orphans, real bullets with hanging indents, shapes without shadows, small open arrowheads, opaque badges, a fixed header tag, chart and label helpers.

## 0.4.0 — The finishing standard with reasons

- New `craft-notes.md` from a slide-by-slide review of a worked example: wrapping and orphans, hanging bullets, colour roles, fixed left-aligned headers, muting in a paler shade of the same hue, no shadows, open arrowheads, top-aligned headers, exact dividers, reference lines, timeline labels, "Contents" naming. Each rule with its reason.
- Checker: shadows (including theme-inherited ones), typed bullet characters, orphaned last words.

## 0.3.0 — Visual identity

- New `visual-identity.md` from five more of the author's decks: palette strategies, brand motifs, brand typography, engaging spoken decks, lessons from an animated explainer.
- More storyline examples (a case for action, a recommendation weighing options) and an explainer-talk arc.
- Checker: handles all-caps title placeholders used as breadcrumbs; balance and chart checks tuned so the author's best decks don't produce false alarms.

## 0.2.0 — Charts, checker and storyline examples

- New `charts.md`: chart or text, a decision guide by relationship (after Abela and Flourish), and the author's chart conventions.
- First scripts: `check_deck.py` (tested on the author's decks and a deck with planted errors), `render.py` and `compare_versions.py`, wired into the QA and revise steps.
- New `storyline-examples.md`, drawn from about 480 headlines across 18 of the author's decks: strong sequences, weak patterns with fixes, and the storyline-proposal format.

## 0.1.0 — First draft

- Named slidesmith while drafting, then renamed **Slidewell** after checking for clashes.
- `SKILL.md` with four modes: new deck, revise, options, share version (after a workshop).
- Story first: intake in two short rounds with sample answers, then storyline approval before building.
- `storytelling.md` (pyramid principle, MECE, SCQA, headline tests), `layout-patterns.md` (about 50 patterns and a density benchmark), `design-rules.md` (the author's 32 design rules and QA checklist).
- Lessons from the author's prompt log built in: fixes are cumulative (a constraints log), judge the rendered slide not the coordinates, copy established formatting rather than recreating it, never overwrite the original, three structurally different options rather than one.
