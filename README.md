# Slidewell

A Claude skill for building PowerPoint decks the way a strong strategy consultant does: **story first, then design, then relentless checking.**

Slidewell exists to prevent the two ways AI-built decks usually go wrong:

1. **Well-formatted slides with the wrong or muddled story.** Slidewell asks a few questions, then proposes the storyline (every headline, in order, and what will prove each one) for you to approve before it builds anything.
2. **Slides that look roughly fine but are full of small inconsistencies, and get worse with every revision.** Slidewell builds from one theme spec, follows a detailed finishing standard, and runs an automated checker, a visual review, and an arithmetic check before it delivers anything. On revisions, it keeps a log of everything you've asked for and flags any change it made to a slide you didn't ask it to touch.

Created by **Ben Marshall**.

## What it makes

- **Findings, recommendation, options-appraisal, and update decks** (steering-group packs, board papers, project updates)
- **Workshop decks**: objectives, agenda, exercises, capture slides, and a share version afterwards
- **Styled templates**: example slides of each kind you'll need, in your brand, with placeholder text to fill in yourself
- **Three alternative versions** of a key slide
- **Revisions** of existing decks, without breaking what's already right

Output is a native, fully editable .pptx: real text boxes, real bullets, native charts.

## Install

**Claude apps (claude.ai, desktop):** download `slidewell.zip`, then upload it in Customize → Skills. Code execution must be turned on, because the skill builds and checks decks with Python.

**Claude Code:** copy the `slidewell` folder into your skills folder (for example `~/.claude/skills/slidewell/`).

**Requirements** (in the environment Claude runs in): Python 3 with `python-pptx`, `lxml`, and `Pillow`; LibreOffice and `pdftoppm` (Poppler) for rendering slide images. Claude's own code environment usually has these or can install them.

## Try it

- "Make a 10-slide recommendation deck for our board from the attached analysis."
- "I'm running a 90-minute workshop with our volunteer leads. Build the deck."
- "Give me a template in our brand for our monthly steering-group update."
- "Give me three options for this executive summary slide."
- "Tidy up slide 4. Don't touch anything else."

Slidewell will ask a short round of questions first (outcome, audience, deck type, brand, and how hands-on you want to be). You can review the storyline and the look before it builds, or let it build straight through. Before delivering, it reviews its own slides in up to three rounds and tells you what it found and fixed.

## How it's organised

`SKILL.md` is what Claude reads first. It points to each reference file at the step where it's needed.

| File | Answers |
|---|---|
| `SKILL.md` | How does Claude work? Modes (new deck, revise, options, share version, template), the storyline gate, build, QA, delivery |
| `references/intake.md` | What do we ask the user, and which slides does each deck type need? |
| `references/storytelling.md` | What should the deck say? Principles: pyramid, grouping and order, SCQA, deck types, headlines |
| `references/storyline-examples.md` | What does that look like in practice? Worked examples, weak patterns, the storyline proposal format |
| `references/visual-identity.md` | What does the deck look like overall? Brand, palette, motifs, type, shape vocabulary, animation |
| `references/layout-patterns.md` | Which layout does each slide use, and how is it built? |
| `references/charts.md` | Everything about charts: whether, which, how, and the finishing details |
| `references/finishing-standard.md` | How is every slide finished? Each rule with its reason, and whether the checker tests it |
| `references/review-brief.md` | What the internal reviewer checks in each review round, starting with the author's most frequent feedback |
| `references/qa.md` | How is the work checked before delivery? Working method, QA loop, accuracy pass, checklists |
| `scripts/check_deck.py` | Automated checker for a .pptx (fonts, sizes, wrapping, spacing, alignment, colour, corners, positions across slides, arithmetic…) |
| `scripts/render.py` | Renders slide images and a contact sheet for visual review |
| `scripts/compare_versions.py` | Lists every change between two versions of a deck and flags unexpected ones |
| `scripts/slidekit.py` | Build helpers that apply the finishing standard automatically |
| `examples/` | Nine fictional decks and the build scripts that made them (see below) |

## Examples

Each deck in `examples/` comes with the script that built it.

| Deck | Type | Shows |
|---|---|---|
| `kelpwise-pitch.pptx` | Recommendation (investment pitch) | Slide-to-slide flow where each headline answers the last; wave-edge motif |
| `purbeck-social-impact-club.pptx` | Short case for action | A lean seven-slide proposal ending in a decision |
| `museum-turnaround.pptx` | Findings and recommendations | Three MECE sections, a small section tracker, a recurring hero image on dividers |
| `library-recommendation.pptx` | Recommendation | Charts with takeaways, show-the-math, a pilot plan and decision bar |
| `market-entry-recommendation.pptx` | Recommendation with options | Forecast line chart, scatter with a target zone, waterfall, labelled RAG options |
| `consulting-proposal.pptx` | Proposal (pitching for work) | Co-branded lockup, phased offer, sprint plan, pricing options, team, case study, logo wall, photo placeholders with overlays |
| `volunteer-workshop.pptx` | Workshop | Objectives, timed agenda tracked through the deck, exercises, capture slide in share-version style |
| `fundraising-workshop.pptx` | Workshop | A four-step method animated on click, then drafting the plan together |
| `animation-options.pptx` | Demo | The same slide revealed as builds, as fade-ins, and as fade-ins that dim earlier items |

All organisations, people, places, and figures in the examples are invented. Photos are placeholder frames labelled with the picture each one needs.

## Using the checker on your own decks

```bash
python scripts/check_deck.py my-deck.pptx            # errors and warnings, slide by slide
python scripts/check_deck.py my-deck.pptx --numbers  # every number on every slide, for an accuracy check
python scripts/render.py my-deck.pptx                # slide images + contact sheet
python scripts/compare_versions.py v1.pptx v2.pptx --focus 4   # what changed, and what changed that shouldn't have
```

## Licence

MIT. See `LICENSE`.
