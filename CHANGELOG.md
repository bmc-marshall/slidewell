# Changelog

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
