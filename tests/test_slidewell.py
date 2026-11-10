"""Slidewell tests: every example deck rebuilds and passes the checker, the checker catches known faults,
template mode works, and the renderer degrades cleanly when LibreOffice is missing.

Run: pip install -r slidewell/requirements.txt pytest && pytest -q
"""
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "slidewell"
SCRIPTS = SKILL / "scripts"
EXAMPLES = SKILL / "examples"
sys.path.insert(0, str(SCRIPTS))

BUILDS = sorted(EXAMPLES.glob("build_*.py")) + sorted((ROOT / "docs" / "showcase").glob("build_*.py"))


def check(pptx):
    out = subprocess.run([sys.executable, str(SCRIPTS / "check_deck.py"), str(pptx), "--json"],
                         capture_output=True, text=True)
    return json.loads(out.stdout)


@pytest.mark.parametrize("build", BUILDS, ids=[b.stem for b in BUILDS])
def test_example_builds_and_passes_checker(build, tmp_path):
    """Each example rebuilds from its script and has zero checker errors."""
    r = subprocess.run([sys.executable, str(build)], capture_output=True, text=True, cwd=tmp_path)
    assert r.returncode == 0, r.stderr[-2000:]
    out = Path(r.stdout.strip().splitlines()[-1])
    assert out.exists() and out.suffix == ".pptx"
    errors = [f for f in check(out) if f["level"] == "error"]
    assert not errors, errors
    props = Presentation(str(out)).core_properties
    assert "canny" not in (props.last_modified_by + props.author).lower(), "python-pptx's default author leaked into the deck"


def test_skill_frontmatter():
    text = (SKILL / "SKILL.md").read_text()
    assert text.startswith("---\n")
    front = text.split("---")[1]
    name = next(l for l in front.splitlines() if l.startswith("name:")).split(":", 1)[1].strip()
    desc = next(l for l in front.splitlines() if l.startswith("description:")).split(":", 1)[1].strip()
    assert name == "slidewell"
    assert 100 < len(desc) <= 1024
    assert "<" not in desc and ">" not in desc


def test_referenced_files_exist():
    """Every references/ and scripts/ path mentioned in SKILL.md exists."""
    import re
    text = (SKILL / "SKILL.md").read_text()
    for path in set(re.findall(r"`((?:references|scripts)/[\w.-]+)`", text)):
        assert (SKILL / path).exists(), path


def _bad_deck(path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    for _ in range(3):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        tb = s.shapes.add_textbox(Inches(0.6), Inches(0.6), Inches(12), Inches(1))
        tb.text_frame.text = "A headline"
        tb.text_frame.paragraphs[0].runs[0].font.size = Pt(28)
    s = prs.slides[1]
    body = s.shapes.add_textbox(Inches(0.6), Inches(2), Inches(6), Inches(2))
    body.text_frame.text = "Body text that is far too small to read across a meeting room, and long enough to count as a block."
    body.text_frame.paragraphs[0].runs[0].font.size = Pt(9)
    cd = CategoryChartData(); cd.categories = ["A", "B", "C"]; cd.add_series("Share", (50, 30, 20))
    s.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(7), Inches(2), Inches(5), Inches(4), cd)
    prs.save(path)


def test_checker_catches_known_faults(tmp_path):
    p = tmp_path / "bad.pptx"
    _bad_deck(p)
    rules = {(f["level"], f["rule"]) for f in check(p)}
    assert ("error", "charts") in rules          # pie chart
    assert ("warning", "alt-text") in rules       # chart without alt text
    assert ("warning", "source") in rules         # data without a source line
    assert any(r == "font-size" for _, r in rules)


def test_template_mode(tmp_path):
    from slidekit import Kit
    template = EXAMPLES / "templates" / "brackwater-template.pptx"
    if not template.exists():
        subprocess.run([sys.executable, str(template.with_name("make_brackwater_template.py"))], check=True)
    kit = Kit(template=str(template))
    assert kit.T["display"] == "Cambria" and kit.T["body"] == "Calibri"
    assert len(kit.prs.slides) == 0
    s = kit.new_slide("Title only")
    kit.frame(s, 1, None, "A headline that goes into the template's title placeholder", "Source: test")
    title = kit.placeholder(s, "title")
    assert title is not None and "headline" in title.text_frame.text
    assert any(ph.placeholder_format.type == 13 for ph in s.placeholders), "slide number field copied from layout"
    s2 = kit.new_slide("Title and content")
    kit.fill(s2, "body", None, bullets=[("Point", 0), ("Detail", 1)])
    body = kit.move(kit.placeholder(s2, "body"), y=3.0, h=2.0)
    assert body.width > 0, "moving a placeholder must not collapse its width"
    out = tmp_path / "t.pptx"
    kit.save(str(out))
    assert not [f for f in check(out) if f["level"] == "error"]


def test_icons_need_no_svg_tools(monkeypatch, tmp_path):
    """Icons are recoloured from bundled PNGs, so no rsvg-convert or cairosvg is needed."""
    from slidekit import Kit
    monkeypatch.setenv("PATH", "/nonexistent")
    monkeypatch.setenv("TMPDIR", str(tmp_path))
    import tempfile
    monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))
    kit = Kit()
    s = kit.new_slide()
    pic = kit.icon(s, "leaf", 1, 1, 0.6)
    assert pic.name == "Icon leaf"


def test_renderer_reports_missing_tools(tmp_path):
    env = dict(os.environ, PATH="/nonexistent")
    deck = next(EXAMPLES.glob("*.pptx"))
    r = subprocess.run([sys.executable, str(SCRIPTS / "render.py"), str(deck), "--out", str(tmp_path)],
                       capture_output=True, text=True, env=env)
    assert r.returncode == 3
    assert "VISUAL REVIEW UNAVAILABLE" in r.stderr
