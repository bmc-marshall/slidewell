# Finishing standard

**Covers:** how every slide is finished: content on the slide, text, colour, recurring elements, space and alignment, shapes and effects, lines and timelines, tables, numbers, naming, and editability. Each rule says what to do and, where it isn't obvious, *why*, so you can apply it to cases not listed.

**Doesn't cover:** chart rules (`charts.md`), the overall look of the deck (`visual-identity.md`), which layout to use (`layout-patterns.md`), or how to run QA (`qa.md`).

Read this before building and again during every visual QA pass. The rules come from the skill author's own guidance and from line-by-line reviews of built decks. These are acceptance criteria, not general guidance: a slide is not finished because it looks broadly acceptable or is better than before. Every applicable rule must pass individually **in the rendered output**. Rules marked *(flex)* have a defined exception; everything else is hard.

The underlying goal: a deck that looks polished and calm, where nothing moves, clashes, or competes unless it means something. Visual "noise" costs the audience attention. Where `check_deck.py` checks a rule automatically, the rule says so.

## Contents
1. Content on the slide
2. Text, type, and wrapping
3. Colour
4. Recurring elements and navigation
5. Space, alignment, and composition
6. Shapes, corners, emphasis, and effects
7. Lines, arrows, markers, and timelines
8. Tables
9. Numbers and wording
10. Naming
11. Editability and reveals

---

## 1. Content on the slide

- **One point per slide**, stated in the headline (see `storytelling.md` for headline tests).
- **Headline:** a message, not a topic; ≤2 lines; no exceptions. A large question or statement centred in the slide body (e.g., a question to the room) is body content, not the headline.
- **Vertical logic:** every chart, textbox, table, callout, and source supports the headline. Remove anything that doesn't.
- **Horizontal logic:** headlines read in sequence tell the story.
- **Visible hierarchy:** use at most one subhead, one subordinate comment, and one chart/table title per content block. Don't stack similar-weight textboxes that compete. Give section headlines, explanatory comments, column headers, and data/content distinct vertical bands. A reader should be able to tell which text is which level from its position and styling alone, without reading it.
- **Sources, notes, and qualifiers** wherever a claim or figure needs one. Every slide with a chart or a table of numbers carries a source or note line, even if it only says the data is illustrative. *Why:* a number without a source is the first thing a sceptical reader questions. *(Checked.)*
- **Make the slide's job clear:** it should be obvious whether the audience is meant to absorb, discuss, decide, brainstorm, or ask questions.
- **Don't invent.** Use only facts, data, sources, and conclusions supported by the materials. Don't add false precision, and don't let commentary claim more than the evidence shows (e.g., a causal claim from purely descriptive data).
- **Don't add design machinery.** Every visual device must mean something. Numbering, colour coding, badges, borders, background fills, and extra text weights all imply structure; if the structure isn't real (e.g., ranking markers implying a difference between items that are equal), remove them.

## 2. Text, type, and wrapping

**Fonts and sizes**
- **Fonts:** at most two, and prefer **one family across each slide** (headlines in the body font, bold). A second, display font is best kept for the title and closing slides. **Never use Calibri** (too common to look considered), and avoid Papyrus and decorative fonts unless the template uses them intentionally. *(Checked.)*
- **No underlines, italics, or highlighting** unless the template uses them intentionally. Use bold for key phrases within bullets. Small italic text (≤12 pt) is fine for keys and legends (`charts.md`). *(Checked.)*
- **Minimum sizes (tiered):** *(checked)*
  - Main slide text: **≥16 pt**
  - Supporting text and labels inside boxes: **≥14 pt** (larger wherever space allows)
  - Sources, footnotes, and chart annotations, tick, category, and data labels: **≥10 pt**, kept consistent with comparable charts in the deck
  - Within a chart, text that names the data (category labels, line and point labels) is larger than text that only gives the scale (value-axis ticks and a units-only title like "People"), e.g. 16 pt vs 12 pt
  - Chart titles and panel labels ("What changed") are at least 16 pt.

  Never shrink the main narrative to make content fit; shorten, split, or move detail to an appendix instead.
- **These are floors, not targets.** When a slide has room (e.g., six short bullets), size text up (18–20 pt) rather than leaving empty space: sparse slides get bigger text, not more white space.
- **Abbreviations** must be understandable in their immediate context (e.g., "19K" in a bullet with no chart context fails).
- **No typos;** terminology matches the deck and the client's naming conventions.

**Wrapping**
- **No orphans.** A paragraph or bullet should never end with a single word (or a tiny fragment like "in full") on its last line. *Why:* a lone word looks like an accident and makes the block look ragged. Fix by rewording, or by breaking the line earlier at a natural phrase boundary (before "that", "and", "to", a comma). Insert a manual line break if needed. Check the line *above* the break too: it must not become the new fragment (keep at least two words and a reasonable length; in a paragraph of three or more lines, no middle line should be much shorter than the others). *(Checked.)*
- **Break lines at phrase boundaries,** not wherever the box edge falls. *Why:* each line should read as a unit; "would cover / that peak in full" reads better than "would cover that peak / in full". The same applies when the last line is merely short (a fragment like "peak in full"): if a phrase boundary is within a word or two, break there.
- **Headlines are the exception: use as much of the top line as possible.** Let the headline run across the full width before it wraps. The only problem case is a second line of just one or two words. Then, in order: (1) reword or abbreviate so the headline fits on one line ("between Mar and Sep 2027" instead of "between March and September 2027"; drop a "The") or so it fills more of line two; (2) if that isn't possible and the words are short, move one word from the top line down; (3) otherwise leave it. Never pull a headline back to an earlier phrase boundary just because one exists; the phrase-boundary rule above is for body text. *Why:* the headline is the widest, most important line on the slide; breaking it early wastes that width and makes it look as though there was no room. (`slidekit` fills headlines automatically and prints a warning when line two has only 1–2 words, so you can reword. *Checked.*)
- **Top-anchored headlines:** anchor headline text to the top of its box so one-line and two-line headlines start at the same height on every slide.
- **Never split a hyphenated word across lines** ("warm- / up"): use a non-breaking hyphen (slidekit does this automatically).
- **Keep text clear of the right margin.** Don't let lines run right up to the margin when a slightly earlier wrap gives a better shape. *Why:* text that touches the edge looks crammed, even if technically inside the margin.

**Bullets and emphasis in text**
- **Hanging indents.** Wrapped lines of a bullet align with the *text* above, not with the bullet. Use real paragraph bullets with a hanging indent, never a typed "•" character. *Why:* the bullet should stand out to the left as a marker; text under it looks like a new item. *(Checked.)*
- **Bullet colour.** Colour bullets in a brand colour (e.g., the dark teal) rather than the text colour. *Why:* black bullets on black text look default and heavy; a brand-coloured bullet ties the slide to the identity. If the brand has a near-black colour, it can be the text colour, with a lighter or accent brand colour for the bullets. Don't use a mid-tone brand colour (like a headline teal) for all body text; it's too light to read in quantity.
- **Make the key words of rules and guidelines stand out.** When a list gives rules, principles, or instructions ("Be present", "Build on ideas", "Park anything off-topic"), bold the key phrase at the start of each item in the main brand colour, and keep the rest regular. *Why:* the audience should be able to take in the rules from the key words alone. (`bullet_list` accepts (key phrase, rest) pairs.)

## 3. Colour

- **Palette:** brand colours only: 1–2 main colours plus up to 4 accents (≤6 total). *(Checked.)*
- **Consistent meaning:** the same colour means the same thing throughout the deck.
- **Body text:** near-black, or a near-black brand colour.
- **Headlines:** the main brand colour.
- **Topic labels and breadcrumbs above the headline** must look different from the headline: put them in a box (the header tag), or make them grey, or use a complementary brand colour. Never the headline's colour. *Why:* a short label and a long headline in the same colour read as one block, and the eye can't tell the signpost from the message. *(Checked.)*
- **Highlight text** (chart titles, panel labels like "What changed", column headers, bold lead-ins): the main brand colour or another *muted* brand colour, never black and never a bright accent. *Why:* they're signposts, so they should be coloured like the brand's structure; black makes them compete with body text, and bright accents (orange) are reserved for the data point the slide is about.
- **Accent (bright) colour:** only for the one thing the slide wants noticed, usually in the chart.
- **One text colour per fill, everywhere.** Decide once which text colour goes on each solid fill (e.g. white on gold, white on the dark brand colour, dark text on cream) and use it in every box, band, bar label, and data label in the deck. Prefer white on solid brand colours. *Why:* the same gold box with white text on one slide and dark text on the next looks like two different components. *(Checked.)*
- **Red/green** must not be adjacent or imply the wrong meaning. *(flex)* RAG ratings may use red/amber/green outside the six-colour limit, but every RAG cell also carries a word or symbol so the rating doesn't depend on colour alone. *(Checked.)*
- Colour in charts: see `charts.md`.

## 4. Recurring elements and navigation

- **Fixed positions, fixed sizes.** The header tag (breadcrumb box), headline, breadcrumbs, logos, source, page number, and other recurring elements sit in exactly the same place and formatting on every slide, and the header box is the **same size** on every slide regardless of how long its label is. Never recreate them approximately; copy them. *Why:* when recurring elements shift or resize from slide to slide, the deck flickers as you click through; consistency is what makes it look polished. *(Checked.)*
- **Left-align header text,** starting at the same x as the headline beneath it (or close to the left edge). Never centre it in its box. *Why:* centred labels of different lengths start at different points on every slide.
- **No chevrons or arrows on header boxes** unless they really are a breadcrumb sequence (several steps shown in a row). *Why:* an arrow implies a flow; if there isn't one, it's misleading decoration. A plain rectangle or a tab with a straight *slanted* end is fine; a pointed end is not.
- **Page numbers** on every slide except title, closing, section-divider, and full-bleed question-to-the-room slides; dynamic, correct, and in the deck's established position. *(Checked.)*
- **The page number stays low in its corner.** Normally the source/note box and the page number sit on the **same baseline**, close to the bottom margin (not floating mid-way up). But if something sits below the notes (a bottom tracker, a breadcrumb), the page number moves down with it, to the lowest baseline, rather than staying level with the notes. Keep it in that same low position on every slide, including slides without notes or tracker. *Why:* a page number floating halfway up the corner just takes up space and makes the corner look dead; its job is to be findable, not aligned. *(Checked.)*
- **Logos and motifs in a corner sit the same distance from both edges they're near.** A logo in the top-right corner has the same gap above it as to its right (`slidekit.corner_xy()`), unless it deliberately bleeds off an edge (like a bookmark hanging from the top). *Why:* unequal gaps make the logo look as if it has slipped. *(Checked.)*
- **Navigation trackers must be introduced and must run through.** A tracker (a row of stage or section names with the current one highlighted) only works if the audience has already met that list: on the agenda, the contents slide, the executive summary, or a framework slide. Show it on every slide in its span, not on a handful. For a workshop, the agenda is usually the right list to track. *Why:* an unexplained tracker is noise, and one that appears and disappears makes the audience wonder what they've missed.
- **Trackers are small and stay out of the headline's way.** Put them bottom-centre (with the notes above them) or top-centre, at 10–12 pt, with the current item bold in the brand colour and the rest muted. Give each item a fixed-width slot so nothing shifts as the bold item moves. Top-centre only if it clears the headline comfortably. Ask the user where they want navigation (intake Q9). *Why:* navigation should be findable, not compete with the message.

## 5. Space, alignment, and composition

**Vertical placement**
- **Distribute the content group vertically between the headline and the notes.** Treat the slide's content as one group (chart and panel, columns, tiles, a timeline plus its decision bar). Place the group so the space between the bottom of the headline text and the top of the group equals the space between the bottom of the group and the top of the notes, or the bottom of the slide (or the footer, such as a bottom tracker) if there are no notes. Measure visible content, not empty box edges. *Why:* equal space above and below gives every chart or block the same relationship to the things around it; a group pushed up under the headline looks crowded at the top and unfinished at the bottom. *(Checked.)*
- **Leave real breathing room: at least ~0.6–0.8 in above and below the content group.** Content should never crowd the headline or the slide edges. Measure from the bottom of the headline text to the top of the visible content (a chart title, the first row), and from the bottom of the visible content to the top of the notes. If the group is too tall to leave that much, make the group smaller (a shorter chart, smaller grid, fewer lines, a size step down for secondary text) rather than squeezing the gaps. This is the single most repeated piece of feedback: when in doubt, the content should sit lower and have more space around it, not be pushed up under the headline. `slidekit.place_balanced` warns when a block is too tall; `check_deck.py` warns below 0.6 in.
- **Clear space between a header and what it heads.** A chart title or panel header needs visible space before the chart or points start (about 0.25–0.35 in between the title text and the top of the plotted data or first point), so the header reads as a label, not as part of the content. Move the chart and points down, not the header; take the space from below the group.

**Consistency across slides**
- **Elements that do the same job sit in the same place on every slide.** Chart titles, charts, panel headers, column headers, and decision or discussion bars should appear at the same height and x from slide to slide, so nothing jumps as the viewer clicks through. Get this by giving the same kind of slide the same group: the same chart height, the same panel width, the same bar height and gap above it. Distributed the same way, they then land in the same place. When two slides share a recurring element (a decision bar on slide 10 and a discussion bar on slide 11), stretch or shrink the rest of each group so the bar lands in the same spot. A slide with a two-line headline will sit about half a line lower than its one-line neighbours; that small shift is acceptable, and shortening the headline removes it. *Why:* the viewer learns where each kind of thing lives; when it moves, they have to search for it again. *(Checked.)*
- **Same kind of slide, same positions, even when the notes differ.** Compute the position once on the most constrained slide of the set (the one with the longest headline or notes) and reuse it for the others, so chart titles, charts, and panel headers don't shift between slides. Don't re-balance each slide on its own if that moves shared elements.
- **Consistent spacing values:** inner margins (between the headline or notes and the content) are consistent across slides, and the gaps between distinct blocks (chart title to chart, chart to panel, content to footer) use the same few spacing values throughout the deck.

**Horizontal placement and balance**
- **Margins and gutters:** consistent outer margins, gutters, and vertical gaps. Left and right outer margins look approximately equal unless the deck establishes otherwise. Don't waste width at the edges while sections in the middle are crowded; redistribute space to separate sections clearly. *(Checked.)*
- **Distribute horizontally too.** A row of separate, similar items (agenda stops, cards, icons, standalone boxes) is spaced so the gaps between the items equal the gaps between the outer items and the slide edges. In PowerPoint, place a shape just off each edge of the slide, select it with the items, and use Distribute Horizontally; `slidekit.distribute_x()` does the same. Items that must line up with margin-aligned content above or below them are distributed between the margins instead. *Why:* equal gaps make a row read as one evenly weighted set, centred on the slide.
- **Balance:** left and right visual weight reasonably balanced; parallel columns or sections share top and bottom boundaries unless there's a deliberate reason not to.

**Alignment**
- **Visual, not mathematical, alignment:** text tied to a chart mark, row, card, or object aligns visually with that object (e.g., category labels vertically centred on their bars, not just sharing a textbox edge).
- **Labels on shapes are centred on the shape.** A centre-aligned label on a bar, band, or box uses the shape's own text frame or a text box with exactly the shape's geometry, so it is centred both ways. *(Checked.)*
- **Callouts:** short recommendation or callout text is vertically centred in its box unless the deck clearly uses another convention.
- **Top-align headers that sit side by side.** If one column header wraps to two lines and its neighbour doesn't, align their tops, not their centres. *Why:* top alignment gives a clean shared line the eye can follow; centring makes the shorter header float.
- **Centre column headers over what they label,** and centre their text in the box. Wrap long headers rather than letting them overhang.
- **Pair alignment.** A hero number beside a label-and-note pair is vertically centred on the pair's *visible text*, not on the boxes. Measure the actual text heights.
- **Dividers span exactly what they divide:** a vertical line between two columns starts at the top of the first text and ends at the bottom of the last text in those columns. Keep it thin (≈0.5–0.75 pt) and in a brand colour.

**Containers and parallel elements**
- **Equivalent-element consistency:** identify every class of repeated element on the slide (cards, metric values, labels, recommendation boxes, bars, category labels, section headers, column headers). Within each class, every instance uses the same font family, size, weight, colour, horizontal and vertical alignment, internal padding, dimensions where appropriate, and border/fill/corner treatment. Any unexplained difference is an error. When parallel elements contain text of different lengths (tiles, cards, columns), size the shared layout to the longest one so every instance keeps the same internal positions. *(Checked.)*
- **Container geometry:** similar containers share corner style, bevel, border, fill, padding, and (where appropriate) dimensions.
- **Internal padding:** balanced inside coloured or outlined containers; content is never pressed to one side or stranded away from another.
- **No overflow:** no text overflowing, clipped, colliding, or spilling out of its container. Fix by shortening text, resizing the box, moving the chart, or reducing *chart-label* size (§2), never by breaking the layout. *(Checked.)*

## 6. Shapes, corners, emphasis, and effects

**Corners: one style per deck** *(checked)*
- **Choose square or rounded corners for the whole deck, and don't mix them.** Rounded corners suit a brand with circular shapes (a round logo, an arc motif); otherwise square is the safer default. *Why:* mixing straight and curved corners looks accidental, as if boxes came from different decks.
- **Keep rounding subtle:** a corner radius of about 0.1–0.15 in, never more than about 0.3 in, whatever the box size. Very rounded boxes look like app buttons, not a professional deck.
- **Rounded header box plus content box: round only the outside corners.** Round the header's top corners and the content box's bottom corners, and let the two meet on a straight, parallel edge. Never stack a fully rounded header on a fully rounded content box. The same idea works sideways: a row label rounded on its left joined to a content box rounded on its right.
- **Square boxes *inside* a rounded container are fine** (cells in a grid, content blocks inside a rounded panel), because they don't sit side by side with rounded boxes of the same kind.
- `slidekit.box()` applies the theme's corner style everywhere; `corners="top"` / `"bottom"` build header/content pairs.

**Muting and emphasis**
- **Fade each element along its own hue.** When something is greyed out (inactive tracker rows, past steps), every element fades towards white in its own colour: a teal box becomes a much paler teal, teal text a pale teal, and black or near-black text a mid-light grey. Never switch a coloured element to grey, or grey text to a tint of the brand colour. *Why:* each element stays recognisably itself, just quieter; mixing hues while fading creates a second palette that clashes.
- **Emphasise part of a set without resizing it.** To show where the focus is in a sequence of identical markers (agenda stops, steps, timeline points), keep the markers the same size and use a grouping device instead: bold the times and labels, draw a pale outline box around the focus items, or add a bracket beneath them with a short label ("Main focus: 40 minutes"). *Why:* enlarged markers look like a different kind of thing, or like a mistake, rather than "more important".
- **Bold the primary layer only.** When bold marks the main items of a graphic (event labels above a timeline), keep secondary information (the months beneath it, axis labels) regular. *Why:* bolding everything flattens the hierarchy, so nothing stands out.
- **Badges are opaque.** A badge or circle placed over a box edge or corner must fully hide what's underneath: a white-filled outer ring with the coloured inner disc on top. *Why:* a corner showing through looks unfinished.

**Effects**
- **No shadows** on boxes, text, lines, or arrows, including shadows inherited from the theme. *Why:* they look dated and add visual noise. The one exception: images clearly brought in from elsewhere (screenshots, newspaper clippings, designs), where a soft shadow makes them look pasted in, which helps the audience see they're external. *(Checked.)*
- **No other decorative effects** (glows, bevels, 3D, gradients) unless the brand or template uses them deliberately.

## 7. Lines, arrows, markers, and timelines

- **Arrowheads are small and open** (line arrowheads, not filled triangles), and arrows **never** have shadows.
- **One arrow looks like one arrow.** If an arrow changes colour where it crosses a background (two-toned), build it from straight segments at exactly the same angle, or from curves that join into one continuous arc. Never join a curve to a straight segment.
- **Short, thin connectors** from labels to points stop just short of the marker.
- **Markers between items sit exactly halfway.** Arrows or triangles that point from one item to the next are centred in the gap between them, with equal space before and after, measured from the visible text or shape edges.
- **Marker sizes are modest:** large dots and chunky arrowheads look old-fashioned.
- **Connector words ("or", "vs", "and") between two things** sit in a single solid circle or a small box, centred vertically on the things on either side (the middle of the two columns, not level with their headers). Keep the double-ring badge style for numbers, dates, and steps. *Why:* the word joins the two sides as wholes; a badge styled like a step number makes it look like item 1 of a sequence.
- **Timeline labels stay on one level; wrap long ones** before resorting to staggered heights. Stagger only when labels would otherwise overlap significantly. *Why:* mixed heights add visual noise and suggest a hierarchy that isn't there.
- **Future-pointing timeline ends:** end the track with a chevron-shaped tip on the bar itself, or a very small arrowhead barely taller than the bar. Large standalone arrowheads look dated.

## 8. Tables

- Use a table only when the reader needs exact lookup; otherwise summarize and visualize the key comparison. The key message should be graspable without searching the table.
- Sort by the column most relevant to the slide's message.
- Declutter: minimal shading, gridlines, and title differentiation. Fields in a table behave like a true table (aligned columns, consistent row heights).
- A key for a table (a rating scale, symbols) follows the chart-key rule: small, italic, grey, just above the top-right corner; better still, label the highlighted item directly (`charts.md` §5b).

## 9. Numbers and wording

- **Sensible rounding.** In prose, prefer natural quantities ("about 19,000 vehicles per day"). Use K/M/B only where the chart context makes them obvious; in charts, state the scale in the axis title (`charts.md`).
- **Durations in natural units:** "2 years, 8 months" or "2 yrs, 8 mo", never "2.7 years"; only halves may be fractional, written "1½ years".
- **Every derived number is correct and unambiguous.** Sums, shares, growth rates, durations, and totals are re-computed in code (`qa.md`, accuracy pass); a duration label states the range it covers ("10:45–11:25, 40 minutes"), and a highlight box covers exactly the items its label counts. *(Partly checked.)*
- **Mark the status of numbers:** estimates ("~", "30+", footnoted), illustrative examples ("Illustrative" banner, "[DUMMY DATA]"), and work in progress ("WIP"). Put assumptions on an assumptions slide when figures depend on them.

## 10. Naming

- **Header labels (breadcrumbs, topic tags) need enough context to stand alone,** in 1–3 words: "Market growth" or "Locations", not "Market"; "Store financials", not "Economics"; "Strategic options", not "Options". Name each slide's specific topic; two slides may share a label if they cover the same topic. *Why:* a one-word label like "Market" could mean size, growth, competitors or customers, so it doesn't tell the reader where they are.
- **Executive-summary labels in a findings or recommendation deck are mini findings, not topics:** "Fast-growing market", "Profitable stores", "Pilot is the best route", "£2.4M to approve", not "Market", "Economics", "Options", "Decision". Each says something, like a short headline. Don't reuse them as header tags; tags name topics, summary labels state findings. (Workshop agendas and contents can use short topic words, because they describe what the room will do, not what was found.) *Why:* a summary of topics tells the reader what was looked at; a summary of findings tells them what to think.
- **An executive summary repeated as a tracker is a "Contents" slide.** Label the first, full version "Executive summary"; label the repeats (with rows greyed out) "Contents". *Why:* once it's navigation, calling it a summary misdescribes it.

## 11. Editability and reveals

- Use native text boxes, shapes, lines, tables, and charts. Keep logically separate content in separate objects.
- Never flatten a slide or major component into an image.
- **Alt text.** Every chart has alt text saying what it shows and its main finding (`kit.alt(chart, "...")`; slidekit otherwise writes a plain description from the chart's data). Real photos carry a short description; icons and pure decoration are marked decorative (slidekit does this for icons). *Why:* decks are read with screen readers, and public-sector clients often require it. *(Checked.)*
- Avoid unnecessary grouping, custom geometry, and fragile effects; prefer simple, robust construction the user can edit by hand.
- Progressive reveals follow the animation decision from intake: builds as successive slides when the deck isn't animated (or will be shared as a file); native fade-in on click when it is, dimming earlier items only when the newest one takes priority (`visual-identity.md` §6).
