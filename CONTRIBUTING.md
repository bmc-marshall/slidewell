# Contributing to Slidewell

Thanks for helping. The most useful contributions are:

- **A deck that came out wrong.** Open an issue with a screenshot of the slide, what you asked for, and the output of `python slidewell/scripts/check_deck.py your-deck.pptx`. Remove anything confidential first.
- **A new checker rule.** Add it to `slidewell/scripts/check_deck.py`, add the rule (with its reason) to `references/finishing-standard.md`, and add a test in `tests/` with a small deck that breaks it.
- **A new layout pattern.** Add it to `references/layout-patterns.md` and use it in an example deck, so it's rebuilt and checked on every change.

## Running the tests

```bash
pip install -r slidewell/requirements.txt pytest
pytest -q
```

The tests rebuild every example deck and require zero checker errors. On Linux, install `fonts-liberation fonts-crosextra-carlito fonts-crosextra-caladea` so text is measured with the same widths as Arial, Calibri and Cambria.

## Style

The instructions are written for Claude to follow, in plain English: say what to do and why, give one worked example, and keep each rule in one place. Examples use only invented organisations, people and figures.

## Releases

Bump the version in `.claude-plugin/plugin.json` and `pyproject.toml`, add a line to `CHANGELOG.md`, and publish a release on GitHub with a new tag like `v1.1.0`. The workflow attaches `slidewell.zip` to the release.
