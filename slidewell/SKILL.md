---
name: slidewell
description: Build and revise consulting-grade PowerPoint (.pptx) decks that tell a clear story and pass a strict finishing check: recommendations, findings, options appraisals, board and steering-group papers, proposals, pitches and facilitated workshops, or a styled template of the slides needed. Use when the user wants a PowerPoint deck or .pptx slides written, restructured, redesigned, checked or fixed, matched to a brand or .pptx/.potx template, offered as alternative versions, or built for a workshop (agenda, exercises, capture slides, share version). Also use for "make me a deck" or "slides for the board" when no other format is named. Not for Google Slides, Keynote, HTML or web slides, or a deck the user asked to make in another named tool or format.
license: MIT (see LICENSE). Copyright (c) 2026 Ben Marshall
---

# Slidewell

Slidewell builds decks the way a strong strategy consultant does: **story first, then design, then relentless checking.** The two failure modes it exists to prevent are (1) well-formatted decks with the wrong or muddled story, and (2) AI-built slides that look roughly fine but are full of small inconsistencies, and that get *worse* as they are revised. Read this file fully, then read the references when each step tells you to.

## References and scripts

| File | Read when |
|---|---|
| `references/intake.md` | Before asking the user anything |
| `references/storytelling.md` | Before drafting the storyline or any headline: the principles (pyramid, grouping and order, SCQA, deck types, headline tests) |
| `references/storyline-examples.md` | With storytelling.md: the same principles shown in worked examples (strong and weak headline sequences, the flow check, grouping in practice, workshop arcs) and the exact storyline-proposal format |
| `references/layout-patterns.md` | When storyboarding and when building each slide |
| `references/visual-identity.md` | At the theme step when there's no template: palettes, brand motifs, fonts, engaging decks, animation |
| `references/charts.md` | Whenever a slide contains numbers: deciding chart vs text, choosing and building charts |
| `references/finishing-standard.md` | Before building, and during every visual QA pass: every rule for finishing a slide (content, text and wrapping, colour, recurring elements, space and alignment, shapes, lines and timelines, tables, numbers, naming, editability), each with its reason |
| `references/review-brief.md` | At each review round: the brief the reviewer works from, including the feedback this skill's author gives most often |
| `references/qa.md` | Before delivering anything: the working method, the QA loop, the accuracy pass, and the per-slide and deck checklists |
| `scripts/check_deck.py` | Run on every .pptx before every delivery |
| `scripts/render.py` | Every visual QA pass: slide images + a contact sheet of the whole deck |
| `scripts/compare_versions.py` | After every revision: lists every change, flags changes on slides you weren't asked to touch |
| `scripts/slidekit.py` | When building: helpers that apply the finishing standard automatically (wrapping, bullets, distribution, fixed positions, corners, labels, animation). Read its docstring first |
| `examples/` | When you want a concrete reference: eleven fictional decks (pitch, case for action, findings and recommendations, proposal, a steering-group update built inside a corporate template, two workshops, a decision workshop comparing two concepts on identical slides, animation options, and more) with the build scripts that made them; the README lists what each shows |

### If these files aren't there
Some hosts load only this SKILL.md, without the folders beside it. **Before step 1, check that `references/` and `scripts/` exist next to this file.** If they don't, get them before doing anything else, because the rest of this skill depends on them:

1. **Fetch the matching release** into a scratch folder, and use its `slidewell/` folder as the base directory for every path in the table above:
   `git clone --depth 1 --branch v1.3.0 https://github.com/bmc-marshall/slidewell <scratch>/slidewell-src`
   Then install the requirements if python-pptx is missing: `pip install -r <scratch>/slidewell-src/slidewell/requirements.txt`. Don't mention this step to the user unless it fails.
2. **If the clone fails** (no network, or GitHub blocked), ask the user to attach the slidewell .zip, unzip it into its own scratch folder, and use that.
3. **If neither works**, tell the user in one line that the references and checker aren't available in this session, and follow this file alone. Stand in for the scripts by hand: render with `soffice --headless --convert-to pdf` then `pdftoppm -png` and look at every changed slide; check revisions by listing each slide's shapes (name, position, size, text) before and after and confirming only the intended slides changed. Say in the QA report that `check_deck.py` did not run. Never say a deck passed the checker when it didn't.

*Maintainer note: when you release a new version, update the tag in step 1 to match.*

## Choose the mode

| The user wants… | Mode |
|---|---|
| A new deck (findings, recommendation, options appraisal, update, workshop, or a mix) | **A. New deck** |
| Changes to existing slides, or new slides added to an existing deck | **B. Revise** |
| Several alternative versions of a slide | **C. Options** |
| A post-workshop version with the session's input | **D. Share version** |
| A styled, empty template of the slides they'll need, to fill in themselves | **E. Template** |

Modes combine (e.g., a new deck with options for the executive summary).

---

## A. New deck

### Short decks: the fast path
For **five content slides or fewer** (a one-pager, a short update, a single decision slide, "just three slides on X"), keep the same gates but a lighter route, because the full reference set costs more than the deck is worth:
- **Intake:** ask only Round 1 of `references/intake.md`, in one message; skip Round 2 unless the answer leaves the look or audience unclear.
- **Storyline:** propose the headlines and what proves each, in a few lines. Read `references/storytelling.md` §2 (pyramid) and §6 (headlines) only; skip storyline-examples.md.
- **Build:** read only the sections of `references/layout-patterns.md` and `references/charts.md` for the patterns you'll use. Read finishing-standard.md §1–2 and §9 (content, text, numbers); `slidekit.py` applies most rules automatically.
- **QA:** run `check_deck.py` (including `--numbers`) and one render-and-look pass; one review round instead of up to three.
- Tell the user you're using the short route, and offer the full route if the deck is high-stakes.

### 1. Intake
Read `references/intake.md` and run its two short rounds **before proposing or mocking up anything**. Don't ask what the user has already stated, but never skip a Round 1 question on an inference: if the request doesn't say whether the deck will be read or presented, how long it is, or who it's for, ask. Do offer your best guess as the proposed answer, so the user can just confirm it. Assumptions are allowed only when the user has said to just go or isn't there to answer.

### Checkpoints
Intake Q5 asks how hands-on the user wants to be. Honour the answer at the two gates below:

| Choice | Storyline gate (step 2) | Look gate (step 3) | Implication to tell the user |
|---|---|---|---|
| Review story and look *(recommended for important decks)* | Wait for approval | Show the theme (palette, fonts, motif mock-up, one sample slide) and wait | One or two extra messages; far fewer revisions |
| Review story only | Wait for approval | Decide and build; describe the look in one line at delivery | Faster; the look may need a revision round |
| Build straight through | List assumptions at the top of your reply and continue | Decide and build | Fastest; more changes afterwards, and a wrong story means a bigger rebuild |

If the user didn't answer, recommend the first option for anything client-facing or high-stakes, and the third for quick internal drafts, and say which you chose. If the user is away (a scheduled or unattended run), build straight through and say so.

### 2. Storyline (approval gate)
Read `references/storytelling.md` and `references/storyline-examples.md`. Then give the user, in the conversation, a proposal in the format of storyline-examples §5:
- the governing message (one sentence) and the SCQA opening;
- the **headline sequence**: every slide's headline, in order, grouped by section;
- for each slide, one line on what will prove it (the layout pattern or chart type and the evidence), and whether it's for presenting, discussing, deciding, or doing an exercise;
- the building blocks you've included (from the intake menu);
- any assumptions or gaps in the material.

Test it before sending: run the flow check in storyline-examples §1b (read the headlines aloud, name the link between each pair, fix any break); do they tell the whole story? Is each grouping MECE? Does any headline make a claim the material can't support?

**Wait for approval or edits before building** unless the user chose to build straight through (Checkpoints). Fixing a storyline costs one message; fixing a built deck costs an hour.

### 3. Theme
Decide the visual system once, write it down as a spec (a small JSON or Python dict of constants), and build every slide from it. A single spec is what makes recurring elements identical instead of approximately similar.

The theme must be **specific to this deck**: drawn from the subject's brand, motif, and world, not a default navy-and-grey that would fit any deck. If you're building several decks, each gets its own theme step; reusing one generic look across them is how decks start to look templated. Shared code is for mechanics (wrapping, positions, checks), never for layouts or looks.

**If the user gives a reference deck to follow or be "inspired by"**, ask which they want, and show both at the look checkpoint: (a) **close to the reference**: its layout grid, header treatment, colour roles, imagery density, and icon style, with only the brand, names, and content changed; or (b) **a new look** for this deck. Don't assume (b) because the content is new: people who share a deck they like usually want its style kept.

Source the theme in this order:
1. **The user's template** (.pptx/.potx): use its masters, layouts, and placeholders. Read its existing slides for conventions (breadcrumb style, source position, page numbers, box styles) and copy them. Build with `Kit(theme, template="their.pptx")`: it removes the template's slides, keeps its masters and layouts, and reads its colours, fonts, headline position and footer into the theme. Then `kit.new_slide("Title only")` uses a named layout, `kit.frame()` puts the headline in the title placeholder, `kit.fill(slide, "body", bullets=[...])` keeps the template's own bullet styles, and `kit.move()` repositions a placeholder safely. See `examples/build_brackwater_update.py`.
2. **The subject's brand**: if there's no template, match the brand of the organization the deck is *about or for* (this is how the best decks feel native to the client). Follow `references/visual-identity.md`: choose a palette strategy, find a brand motif (ask the user; if they have none in mind, mock up two or three options as images and let them choose), pair fonts, and bookend the deck. Use the logo colors, brand guidelines, or website the user provides. If they give none (including when they skip or don't answer the brand questions) and the organization is real, **extract the colours yourself** (below); don't pick them from memory or by eye. Choose 1–2 main colors and up to 4 accents, a headline font and a body font (or one family), and an image style. Use logos only if the user supplies them or confirms they may be used. For photos, logos, and icons, follow visual-identity §4b–4c: every photo has a role, text on photos gets an overlay, and missing photos become labelled placeholder frames at their final size.

   **Extracting brand colours when the user gives none.** Do this before the look checkpoint, and record every hex value with where it came from.
   1. **Find the official site** with a web search (the organization's own domain, not a directory, news story or reseller). Fetch the home page, and its brand, press or media page if there is one: published brand guidelines or a press kit give official hex values, and those beat anything you sample.
   2. **Read the logo.** Find it in the page header (the `<img>`, inline `<svg>`, or `og:image` / `apple-touch-icon` link). If it's an SVG, read the `fill` and `stroke` hex values directly. If it's a raster image you can download, sample it with Pillow: ignore transparent pixels, near-white, near-black and greys, group what's left by colour, and take the two to four most common groups as hex. If you can't download it, render or screenshot it and sample that.
   3. **Read the website's colours** as a cross-check: CSS custom properties with names like `--primary`, `--brand`, `--accent`, the `theme-color` meta tag, and the fills of the header, buttons and links. Ignore colours from photos, cookie banners, and third-party widgets.
   4. **Decide.** The logo's dominant colour is the main colour unless published guidelines say otherwise; the website supplies secondary and accent colours where the logo has only one. If logo and site disagree, prefer the logo and mention it. Then choose a palette strategy (visual-identity §2) from these values.
   5. **Show the result at the look checkpoint** (or at the top of the build if building straight through): a row of swatches with each hex value and its source ("#0B3D91, logo SVG fill"; "#F26B21, website button colour"), so the user can correct them in one reply.
   6. **If you can't reach the site or logo**, say so in one line and ask for the logo, a brand-guidelines link, or hex values; if the user isn't there to ask, use the neutral default (3) and say so. If the organization is fictional, or you can't find a website for it, use the neutral default or build colours from the deck's subject, and say which.

   Extracting colours from the logo doesn't mean using the logo: it still appears on slides only if the user supplies it or confirms it may be used.
3. **Neutral default**: white background, one deep main color (e.g., navy), one accent, a clean sans-serif (e.g., Inter or Arial), generous margins.

The spec must include: slide size (16:9 default), outer margins, the headline band (position, size, color, max two lines), breadcrumb style and position, body text sizes per tier, source/footnote position and size, page-number position, palette with each color's purpose, box styles (fill, border, corner, padding), the shape vocabulary (which shape or style means what, e.g. highlight pills, callouts, rating markers; see visual-identity §4b), chart styling, and the gutter unit.

**Fonts:** check which fonts are available for rendering (`fc-list`, or the system font folder on a Mac or Windows). If the brand font isn't installed, renders will substitute it, which changes text widths and can create false overflow or hide real overflow. Either use a font available on both sides, or keep generous slack in text boxes (~10–15%) and tell the user to check text fit on their own machine. Ask (intake Q8) whether the brand font is on the presenting machine. When you have a free choice, prefer fonts the user is likely to have *and* that have exact-metric stand-ins for rendering here (Arial ↔ Liberation Sans, Times New Roman ↔ Liberation Serif, Cambria ↔ Caladea): text then wraps identically in your checks and on their machine.

### 4. Build
**Tell the user how long it will take and roughly how much usage it needs, before you start.** Give a range and how confident you are. Base it on measured builds, not on how long the work feels:
- *Measured:* a 16-slide proposal (6 complex slides: chart, phases, Gantt plan, pricing, case study, photo layouts) took **about 11 minutes** from "make the deck" to delivery, including three fix rounds. It used about 70K output tokens, and about 8M input tokens read from cache, because every step re-reads the whole conversation.
- *Rule of thumb:* about 10–20 minutes for 12–20 slides, more for decks with many charts or a brand template to match, less for a template or a repeat of an earlier deck. Medium confidence: the environment's speed varies.
- *Usage:* the cost is driven mostly by **how long the conversation already is**, not by the deck: each build step re-reads everything before it. A build at the end of a long chat can cost several times more than the same build in a fresh one. If the conversation is long and the user is watching their limits, suggest building in a new chat with the approved storyline pasted in. Anthropic doesn't publish plan limits as token numbers, so don't convert the estimate into a share of a Free or Pro allowance; point the user to Settings > Usage to see where they stand.
Say it in one line ("16 slides, about 10–20 minutes including checks; medium confidence"). If you report how long a build took, measure it (timestamps in the log, or the clock), never estimate it after the fact. If the build runs past the top of the range, say why and how much is left.

Build native, editable .pptx with python-pptx (`pip install python-pptx` if missing; if a pptx skill is available, its guidance on mechanics applies too). Follow `references/finishing-standard.md` §11 (editability) throughout.

**The file must open cleanly in PowerPoint.** Use python-pptx's API (or a slidekit helper) for anything it supports. When you must edit XML directly, insert each element in schema order and replace an existing child rather than adding a second one. PowerPoint refuses out-of-order or duplicated elements and offers to "repair" the deck, which is the worst possible first impression. LibreOffice renders the same file without complaint, so the renders won't show the problem; `check_deck.py` reports these as `file-integrity` errors. Don't copy raw-XML snippets from old build scripts without checking what python-pptx already wrote.

- Choose each slide's pattern from `references/layout-patterns.md` by the job the slide does. Respect the content-density benchmark (§H).
- Write headlines exactly as approved; if building reveals that a headline can't be supported, flag it rather than quietly changing it.
- **Use text the user supplied word for word.** Bullets, labels, quotes, and headlines from a storyboard, brief, or spreadsheet go on the slide exactly as given (fixing only obvious typos, and saying so). If you think a supplied line should change (it's too long, unsupported, or would read better with a number), keep the original on the slide and suggest the edit in your reply. *Why:* the user, or their boss, has already agreed that wording; a silent rewrite breaks that agreement and is hard to spot.
- Draw every recurring element (breadcrumb, headline, source, page number, logo) from the theme spec, using the same function every time.
- Build parallel elements (columns, cards, pillars) in a loop from one definition, so they cannot drift apart.
- Numbers: for every slide with two or more values being compared, use a chart, not text (`references/charts.md` §1). Choose and build it per `references/charts.md`; annotations stay native and editable.
- Workshop decks: every interaction slide makes clear what the room should do.
- Reveals: follow the animation answer from intake (visual-identity §6). Not animated → builds as successive slides (n/N). Animated → fade in on click; dim earlier items only when the newest item supersedes them.
- Mark the status of content honestly: "Illustrative", "[DUMMY DATA]", "WIP", assumptions slides.
- For long decks, build section by section and run QA as you go.

### 5. QA loop (delivery gate)
Run the full loop in `references/qa.md` on every slide before delivering anything. In short:
1. `python scripts/check_deck.py deck.pptx`: fix every error; review every warning.
2. `python scripts/render.py deck.pptx`: look at every slide, then the contact sheet. If it exits with VISUAL REVIEW UNAVAILABLE (LibreOffice or Poppler missing, common in Claude Code on a laptop), don't try to install them unasked: finish every other check, deliver, and tell the user in one line that the visual review was skipped and how to enable it.
3. Hunt for failures against the per-slide checklist (`qa.md`), `references/finishing-standard.md`, and `references/charts.md`.
4. Accuracy pass: `check_deck.py --numbers`, then re-derive every computed figure in code and assert it in the build script.
5. Fix, re-render, restart the checklist for the whole slide; repeat until clean. Then the deck-level checklist.

6. **Review rounds** (`qa.md` §2b): a reviewer that hasn't seen the build looks at the full-size slide images against `references/review-brief.md`, lists issues worst first; you fix, re-render, and review again, up to three rounds or until nothing serious remains.

Never say a deck is presentation-ready or QA'd unless this loop actually ran.

### 6. Deliver
Send the .pptx with a short **QA report**:
- what the deck contains, in two or three lines (don't describe every slide);
- the review rounds: how many ran, the main issues found and fixed, and anything left open with the reason;
- any rule you deliberately flexed and why;
- what the user should check themselves (fonts on their machine, data you marked as assumed, animated slides in Slide Show, photos to drop into placeholders).

A missing or thin QA report tells the user QA didn't really happen, so never skip it.

---

## B. Revise

Revisions are where AI-built decks usually get worse. Work like this:

0. **Ask before you build, briefly.** Unless the user has already answered these, ask in one message, with your best guess as the proposed answer so they can just confirm:
   - **Audience and use:** who will see it, and will it be presented live or read? (This decides what belongs on a slide: a quote that persuades an external audience is clutter for an internal one.)
   - **One version or options?** For each slide you're redesigning or adding: "one recommended version, or three genuinely different options to choose from (mode C)?" Recommend options for a slide that carries a key message or that the user has said is important.
   - **Keep the existing format?** "Are you happy with how the deck looks now (fonts, sizes, colours, layout), so new and revised slides should match it?" Before asking, look at the existing slides against the finishing standard, and if anything falls short, name the two or three biggest gaps in the question ("the body text is 12 pt, below the 16 pt minimum; the charts have no axis titles"), so the user is choosing with that in mind.
     - **Yes:** new and revised slides match the existing format (step 3).
     - **No, or the user wants it improved:** propose the format changes as a short list (what changes, on which slides, and why, citing the finishing standard), and wait for sign-off before building. Once agreed, apply them to the existing slides as well as the new ones, so the deck stays consistent, and record each agreed change in the constraints log. Changing slides the user didn't ask about is only allowed with this sign-off.
   - **How hands-on?** Review a mock-up of the first new slide before the rest, or build straight through (the Checkpoints table).
   Small fixes (a typo, a moved label, one colour) don't need this; a redesigned or new slide does. If the user is away or said to just go, state your assumptions at the top of the reply instead.
1. **Start a constraints log** (`<deck>_constraints.md`) the first time you revise a deck, and keep it for the whole conversation. Record every requirement the user has stated and every fix they've accepted, one line each. Every later revision must satisfy *all* of it, not just the newest request.
2. **Inspect first:** render the target slide and the slides around it that define the visual language. Note the recurring elements and conventions you must preserve.
   - **Read `references/finishing-standard.md` in full before building any new or redesigned slide**, and run at least one review round (`qa.md` §2b) before delivering. *Why:* revisions feel small, so the standard gets skipped, and most revision mistakes (mismatched arrowheads, unneeded borders, a note not centred in its space) break rules it already states.
   - **Name the slide each new slide mirrors, if there is one, and copy it element by element.** If new slide 16 restates slide 15 with different inputs, its left column copies slide 15's left column: the same text size, icons, spacing, and borders (or lack of them). Don't copy from a different slide or from a generic box style.
3. **Change only what's asked.** Copy existing formatting and geometry from established elements; never recreate a box, breadcrumb, or chart style approximately.
   - **New slides take the deck's own sizes** (when the user is keeping the existing format; otherwise use the format they signed off in step 0). Before building a new slide in an existing deck, measure the comparable slides: body and bullet size (e.g., the executive summary's 18 pt bullets), headline size and weight, box fills and corners, chart label sizes. Where those meet the finishing standard, use them exactly, even if the standard would allow smaller. Copy a headline's run formatting from an existing headline (`kit.fill(..., like=existing_title)`): a title placeholder often inherits a much larger size than the deck's slides actually use.
   - **Size containers to their content.** A chevron, header bar, or box is only as tall as its text plus normal padding; spare height goes to the content below, not into the shape.
   - **Working from a sketch or marked-up printout:**
     - **Draw only what's drawn, and keep what's drawn.** Don't add boxes, outlines, borders, or fills the sketch doesn't show. Where the sketch deliberately draws an element (a box around a note, an arrow type), that takes priority over what the deck usually does. Everything the sketch only shows roughly (header styles, fonts, colours, box fills) comes from the deck, copied exactly. *Be careful:* a drawn element is the *what*, not licence to restyle. In one revision the agent rightly kept a deck's header style but wrongly changed the drawn arrow shapes.
     - **Follow the drawn structure exactly:** the arrow type (angular or curved, not straight diagonals), which item points to which, and the order. Don't add bends the layout doesn't need: once items are lined up, an arrow between two level items is straight, even if the hand-drawn line wobbled or jogged. If a mapping or word is hard to read, ask while the user is there. Otherwise build your best reading and list it as uncertain. Never swap in a mapping you think is more logical, or reorder items to make the lines neater.
     - **Transcribe handwriting carefully.** Check each uncertain word against the deck's own wording and any source photos or notes ("family or friend", not "family or client"), and list in the reply the words you weren't sure of.
     - **New devices use the deck's existing look:** its callout style, arrowheads, and line weights (finishing-standard §5 and §7 cover callout pointers, mapping arrows, and spacing free-standing notes).
   - **Placeholder slides:** if the deck has slides marked to be replaced ("Replace this slide with your revised version of slide 12"), build the new version there and leave the original untouched as the "before". Leave slides meant for the user's own writing (a reflection, their assessment) for them, unless they ask you to draft them.
4. **Save a new version** (`deck_v2.pptx`) instead of overwriting, unless the user asks otherwise.
5. Run `python scripts/compare_versions.py old.pptx new.pptx --focus <slides you were asked to change>`. Every UNEXPECTED change is a regression until proven otherwise: undo it. Then run the full QA loop on every changed slide and its neighbours, **plus** a check of each line in the constraints log.
6. Report what changed, and any request you couldn't satisfy without breaking another constraint (and the trade-off), rather than silently choosing.

## C. Options

1. Understand the deck's storyline and the slide's role in it first.
2. Create **three genuinely different options**: same evidence and message, different *information architecture* (e.g., chart-led vs. structured text vs. framework diagram; comparison vs. synthesis; different grouping of evidence). Don't create variety through color or decoration, and don't force three chart types the evidence doesn't support.
3. Each option is complete and independently credible, follows every design rule, and passes the full QA loop.
4. Place them right after the original slide (keep the original untouched), each labelled in its breadcrumb (e.g., "OPTION A: CHART-LED"). Leave the rest of the deck unchanged.
5. Where the material supports more than fits one clean slide, keep the main design clean and put the extra in an optional extra slide rather than overcrowding.
6. In the reply, give one line per option on its strength and trade-off, and say which you'd choose, or which hybrid would beat all three.

## D. Share version

After a workshop, the user provides notes, photos of flipcharts, or a list of what the room said.
1. Copy the live deck; keep the original.
2. Fill the capture slides with the room's input in a distinct highlight color (default: a clear blue, or another color outside the brand palette), and state in a small note on the first capture slide that this color marks session input.
3. Collapse reveal pairs and n/N builds where the final state alone reads clearly; keep sequences that teach something.
4. Update "objectives revisited" with what was actually covered and agreed, and add next steps if any were decided.
5. Run the full QA loop.

## E. Template

When the user wants a styled template rather than a written deck (intake Q3):
1. Run intake round 1 and the theme step as for a new deck (brand, motif, fonts, corners, navigation). Skip the storyline gate: there's no story yet.
2. Build one example slide for each building block ticked in the section menu for their deck type (e.g. executive summary, evidence slide with chart and takeaways, recommendation, timeline, options matrix, risks, decision bar, next steps), using the matching layout pattern.
3. Fill every slot with placeholder text that says what goes there and how, in square brackets: "[Headline: the one message of this slide, ≤2 lines]", "[Takeaway: bold lead-in, then one line of evidence]", "[Source: …]". Charts get clearly fictional sample data, labelled "Sample data".
4. Keep every element native and editable, positioned exactly as the finished slides would be, so the user's content drops in without re-layout. Add a first slide listing the slide types included and the theme rules (colours and their roles, fonts and sizes) so others can follow them.
5. Run the full QA loop (the accuracy pass doesn't apply to sample data, but mark it as sample).

---

## Principles to keep in mind

- **Direction beats detail.** A deck with the right story and adequate formatting beats a polished deck with the wrong story. That's why the storyline gate exists.
- **Every device must mean something.** Color, numbering, borders, and emphasis imply structure; don't add structure that isn't there.
- **Never invent.** No facts, figures, sources, or precision that aren't in the material. Flag gaps instead.
- **Slides, not documents.** If a slide only works when read slowly, it's a memo; split it or move detail to an appendix. If it only works with a presenter, make sure it's meant to.
- **Keep it editable.** The user will finish the deck by hand; make that easy.
