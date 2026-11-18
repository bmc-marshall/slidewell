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


def test_combo_chart_keeps_both_axes(tmp_path):
    """combo_chart puts the line on its own visible right-hand axis and passes the checker."""
    from slidekit import Kit
    kit = Kit()
    s = kit.new_slide()
    kit.frame(s, 1, None, "Fast-casual's share of spending has doubled since 2021", "Source: test data")
    kit.chart_header(s, 0.6, 1.5, 8, "Discretionary income and fast-casual share, 2018–2024")
    gf = kit.combo_chart(s, 0.6, 2.0, 8, 4.2, ["2018", "2019", "2020", "2021"],
                         [("Income ($)", (31000, 32000, 33000, 30000))], ("Fast-casual share", (0.033, 0.034, 0.027, 0.023)),
                         "Income ($)", "Share of spending", bar_max=60000, line_min=-0.04, line_max=0.05)
    C_ = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
    plot = gf.chart._chartSpace.find(f".//{C_}plotArea")
    line_ax = [a.get("val") for a in plot.find(f"{C_}lineChart").findall(f"{C_}axId")]
    sec = [v for v in plot.findall(f"{C_}valAx") if v.find(f"{C_}axId").get("val") in line_ax]
    assert len(sec) == 1 and sec[0].find(f"{C_}delete").get("val") == "0", "secondary axis must stay visible"
    assert sec[0].find(f"{C_}title") is not None
    assert len(gf.chart.plots) == 2 and len(gf.chart.plots[1].series) == 1
    out = tmp_path / "combo.pptx"
    kit.save(str(out))
    errors = [f for f in check(out) if f["level"] == "error"]
    assert not errors, errors


def test_fill_copies_headline_formatting(tmp_path):
    """fill(like=...) gives a new headline the run formatting of an existing one, not the layout's default."""
    from slidekit import Kit
    template = EXAMPLES / "templates" / "brackwater-template.pptx"
    if not template.exists():
        subprocess.run([sys.executable, str(template.with_name("make_brackwater_template.py"))], check=True)
    kit = Kit(template=str(template))
    s1 = kit.new_slide("Title only")
    t1 = kit.fill(s1, "title", "An existing headline set at 28 pt bold")
    r = t1.text_frame.paragraphs[0].runs[0]; r.font.size = Pt(28); r.font.bold = True
    s2 = kit.new_slide("Title only")
    t2 = kit.fill(s2, "title", "A new headline\vover two lines", like=t1)
    runs = [r for p_ in t2.text_frame.paragraphs for r in p_.runs]
    assert len(runs) == 2 and all(r.font.size == Pt(28) and r.font.bold for r in runs)


def test_legend_text_centred_on_swatch():
    """kit.legend: each swatch and its text box share a centre line, and the text is anchored middle."""
    from slidekit import Kit
    from pptx.enum.text import MSO_ANCHOR
    kit = Kit()
    s = kit.new_slide()
    shapes = kit.legend(s, 9.0, [("Tourists", kit.T["main"], 2.4), ("Students", kit.T["grey"], 3.6)])
    for sw, tx in zip(shapes[0::2], shapes[1::2]):
        assert abs((sw.top + sw.height / 2) - (tx.top + tx.height / 2)) < 2000   # EMU; well under a point
        assert tx.text_frame.vertical_anchor == MSO_ANCHOR.MIDDLE


def test_checker_catches_label_formats_and_chart_plus_table(tmp_path):
    """'$31.0K'-style labels are flagged, and a chart plus a table of numbers on one slide is a redundancy warning."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    cd = CategoryChartData(); cd.categories = ["2018", "2019", "2020"]; cd.add_series("Income", (31000, 32000, 33000))
    ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.6), Inches(1.5), Inches(7), Inches(4), cd).chart
    ch.plots[0].has_data_labels = True; dl = ch.plots[0].data_labels
    dl.show_value = True; dl.number_format = '$#,##0.0,"K"'; dl.number_format_is_linked = False
    t = s.shapes.add_table(4, 2, Inches(8), Inches(1.5), Inches(4), Inches(2)).table
    for r, (a, b) in enumerate([("Year", "Income"), ("2018", "31,000"), ("2019", "32,000"), ("2020", "33,000")]):
        t.cell(r, 0).text, t.cell(r, 1).text = a, b
    p = tmp_path / "labels.pptx"; prs.save(p)
    found = {(f["level"], f["rule"]) for f in check(p)}
    assert ("error", "number-format") in found       # a decimal on values of 10 or more
    assert ("warning", "number-format") in found     # "K" suffix in labels
    assert ("warning", "redundancy") in found        # chart and table of numbers together


def test_skill_folder_has_no_nested_archives():
    """claude.ai rejects a skill zip that contains another zip, so nothing inside slidewell/ may be an archive."""
    bad = [p for p in SKILL.rglob("*") if p.suffix.lower() in (".zip", ".skill", ".tar", ".gz", ".7z", ".rar")]
    assert not bad, bad


def test_icons_load_from_folder(tmp_path):
    from slidekit import Kit, icon_names
    names = icon_names()
    assert len(names) == 114 and "leaf" in names
    kit = Kit(); s = kit.new_slide()
    assert kit.icon(s, "leaf", 1, 1, 0.6).name == "Icon leaf"


def test_combo_chart_title_faces_axis_and_columns_labelled(tmp_path):
    """v1.2.0: the right-hand axis title reads top-to-bottom (faces its axis) and the columns carry value labels."""
    from slidekit import Kit
    kit = Kit()
    s = kit.new_slide()
    kit.frame(s, 1, None, "Revenue has grown every year, but margin has slipped", "Source: test data")
    gf = kit.combo_chart(s, 0.6, 2.0, 8, 4.2, ["2022", "2023", "2024"], [("Revenue (£M)", (3.1, 3.4, 3.9))],
                         ("Margin", (0.21, 0.19, 0.17)), "Revenue (£M)", "Margin", bar_max=8, line_min=-0.25, line_max=0.25)
    C_ = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
    A_ = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    right = [v for v in gf.chart._chartSpace.iter(f"{C_}valAx") if v.find(f"{C_}axPos").get("val") == "r"][0]
    assert right.find(f".//{C_}title//{A_}bodyPr").get("rot") == "5400000"
    assert gf.chart.plots[0].has_data_labels and gf.chart.plots[0].data_labels.number_format == "#,##0.0"
    out = tmp_path / "combo.pptx"; kit.save(str(out))
    assert not [f for f in check(out) if "faces away" in f["message"]]


def test_checker_numbers_series_colour_and_divider(tmp_path):
    """v1.2.0 checks: a digit under ten in running text, bold text in another series' colour, an off-centre divider."""
    from slidekit import Kit, rgb
    kit = Kit()
    s = kit.new_slide()
    kit.frame(s, 1, None, "Referrals have grown faster than income since 2022", "Source: test data")
    cd = CategoryChartData(); cd.categories = ["2022", "2023", "2024"]
    cd.add_series("Referrals", (100, 130, 165)); cd.add_series("Income", (100, 110, 124))
    ch = s.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(0.6), Inches(2), Inches(7), Inches(4), cd).chart
    for ser, col in zip(ch.plots[0].series, ("E8664A", "0E5E6F")):
        ser.format.line.color.rgb = rgb(col); kit.no_markers(ser)
    kit.side_points(s, 8.2, 4.5, 2.1, 5.9, [("Referrals up 65%", "since 2022, 9 of them from schools", rgb("0E5E6F")),
                                             ("Income up 24%", "over the same period", rgb("0E5E6F"))], 16)
    kit.text(s, 0.6, 6.9, 1.2, 0.3, "Tidewater", 13, bold=True, smart=False)
    kit.line(s, 2.0, 6.95, 2.0, 7.15)
    kit.text(s, 2.05, 6.9, 1.2, 0.3, "Northlight", 13, bold=True, smart=False)
    out = tmp_path / "faults.pptx"; kit.save(str(out))
    rules = [f["rule"] for f in check(out)]
    assert "numbers" in rules and "series-colour" in rules and "divider" in rules
    msgs = [f["message"] for f in check(out) if f["rule"] == "series-colour"]
    assert len(msgs) == 1 and "Referrals up 65%" in msgs[0]          # the Income lead-in is in Income's colour: fine


def test_lockup_divider_is_midway(tmp_path):
    """kit.lockup places the divider midway between the measured words, in left-aligned wrapping boxes."""
    from slidekit import Kit
    kit = Kit()
    s = kit.new_slide()
    kit.frame(s, 1, None, "A slide with a co-branded footer lockup", None)
    kit.lockup(s, 0.6, 7.0, [("Tidewater", 13, True, kit.T["main"], None), ("Northlight", 13, True, kit.T["ink"], None)])
    boxes = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text in ("Tidewater", "Northlight")]
    assert all(sh.text_frame.word_wrap for sh in boxes)
    out = tmp_path / "lockup.pptx"; kit.save(str(out))
    assert not [f for f in check(out) if f["rule"] == "divider"]


def test_read_and_keep_format_options(tmp_path):
    """v1.4.0: --read lowers body-text floors; --keep-format reports a client deck's own sizes and fonts as warnings."""
    from slidekit import Kit
    kit = Kit()
    s = kit.new_slide()
    kit.frame(s, 1, None, "Residents most want later opening on weekdays", "Source: test data")
    body = ("Most residents who answered the survey said they would visit more often if the building stayed open "
            "until eight on weekdays, especially parents with school-age children.")
    kit.text(s, 0.6, 2.2, 8.0, 1.6, body, 12, smart=False)
    out = tmp_path / "small.pptx"; kit.save(str(out))

    def run(*flags):
        r = subprocess.run([sys.executable, str(SCRIPTS / "check_deck.py"), str(out), "--json", *flags],
                           capture_output=True, text=True)
        return [(f["level"], f["rule"]) for f in json.loads(r.stdout)]

    assert ("error", "font-size") in run()                       # 12pt body fails the presented-deck floor
    assert not any(r == "font-size" and l == "error" for l, r in run("--read"))   # but passes for a read deck
    kept = run("--keep-format")
    assert ("warning", "font-size") in kept and ("error", "font-size") not in kept
