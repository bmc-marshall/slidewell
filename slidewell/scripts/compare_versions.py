#!/usr/bin/env python3
"""Compare two versions of a deck and list everything that changed.

Usage:
    python compare_versions.py old.pptx new.pptx [--focus 12,14] [--json]

Use it after every revision. Changes on slides you were asked to edit are
listed as expected; changes anywhere else are flagged UNEXPECTED. These are the
regressions that undo earlier fixes. Small position/size changes (≥0.01 in) and
formatting changes (font, size, fill, line, corners, padding) are all reported.

Exit code 2 if there are unexpected changes.
"""
import argparse
import json
import sys

from pptx import Presentation
from pptx.util import Emu

TOL = 9144  # 0.01 inch in EMU


def inch(v):
    return round(Emu(v).inches, 2)


def fill_sig(shape):
    try:
        f = shape.fill
        if f.type is None:
            return None
        if f.type == 1:  # solid
            c = f.fore_color
            return f"solid:{c.rgb}" if c.type == 1 else f"solid:{c.theme_color}"
        return str(f.type)
    except Exception:
        return None


def line_sig(shape):
    try:
        ln = shape.line
        w = ln.width
        try:
            col = str(ln.color.rgb) if ln.color.type == 1 else str(ln.color.theme_color)
        except Exception:
            col = None
        return f"{col}/{inch(w) if w else 0}"
    except Exception:
        return None


def geom_sig(shape):
    el = shape._element
    prst = el.xpath(".//a:prstGeom")
    if not prst:
        return None
    adj = [g.get("fmla") for g in prst[0].xpath(".//a:gd")]
    return f"{prst[0].get('prst')}{adj or ''}"


def text_sig(shape):
    if not getattr(shape, "has_text_frame", False) or not shape.has_text_frame:
        return None, None
    tf = shape.text_frame
    fonts = set()
    for p in tf.paragraphs:
        for r in p.runs:
            if r.text.strip():
                fonts.add((r.font.name, r.font.size.pt if r.font.size else None,
                           r.font.bold, r.font.italic, str(p.alignment)))
    bodyPr = tf._txBody.bodyPr
    pad = tuple(bodyPr.get(k) for k in ("lIns", "tIns", "rIns", "bIns")) + (bodyPr.get("anchor"),)
    # ignore invisible differences: non-breaking vs ordinary hyphens, and manual line breaks vs spaces
    norm = tf.text.replace("\u2011", "-").replace("\x0b", " ").replace("\v", " ")
    return " ".join(norm.split()), (tuple(sorted(map(str, fonts))), pad)


def walk(shapes, prefix=""):
    for sh in shapes:
        key = f"{prefix}{sh.shape_id}:{sh.name}"
        yield key, sh
        if sh.shape_type == 6:  # group
            yield from walk(sh.shapes, prefix=key + "/")


def snapshot(slide):
    snap = {}
    for key, sh in walk(slide.shapes):
        text, fmt = text_sig(sh)
        snap[key] = {
            "name": sh.name,
            "box": (sh.left or 0, sh.top or 0, sh.width or 0, sh.height or 0),
            "text": text,
            "fmt": fmt,
            "fill": fill_sig(sh),
            "line": line_sig(sh),
            "geom": geom_sig(sh),
        }
    return snap


def diff_slide(a, b):
    changes = []
    for k in a.keys() - b.keys():
        changes.append(f"removed '{a[k]['name']}'")
    for k in b.keys() - a.keys():
        changes.append(f"added '{b[k]['name']}'")
    for k in a.keys() & b.keys():
        x, y = a[k], b[k]
        nm = x["name"]
        bx, by = x["box"], y["box"]
        if abs(bx[0] - by[0]) > TOL or abs(bx[1] - by[1]) > TOL:
            changes.append(f"moved '{nm}' ({inch(bx[0])},{inch(bx[1])}) -> ({inch(by[0])},{inch(by[1])}) in")
        if abs(bx[2] - by[2]) > TOL or abs(bx[3] - by[3]) > TOL:
            changes.append(f"resized '{nm}' {inch(bx[2])}x{inch(bx[3])} -> {inch(by[2])}x{inch(by[3])} in")
        if x["text"] != y["text"]:
            changes.append(f"text changed in '{nm}': {short(x['text'])!r} -> {short(y['text'])!r}")
        for f, label in (("fmt", "text formatting/padding"), ("fill", "fill"), ("line", "line"), ("geom", "shape/corners")):
            if x[f] != y[f]:
                changes.append(f"{label} changed in '{nm}': {x[f]} -> {y[f]}")
    return changes


def short(t, n=50):
    if t is None:
        return None
    t = " ".join(t.split())
    return t if len(t) <= n else t[: n - 1] + "…"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("old")
    ap.add_argument("new")
    ap.add_argument("--focus", default="", help="slides you were asked to change, e.g. 12,14 (numbers in the NEW deck)")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()
    focus = {int(x) for x in args.focus.split(",") if x.strip()}

    old, new = Presentation(args.old), Presentation(args.new)
    so, sn = list(old.slides), list(new.slides)
    old_ids = {s.slide_id: i for i, s in enumerate(so, 1)}

    report = {"slides": [], "unexpected": 0}
    matched_old = set()
    for j, s in enumerate(sn, 1):
        i = old_ids.get(s.slide_id)
        if i is None:
            report["slides"].append({"slide": j, "status": "new slide", "expected": j in focus, "changes": []})
            if j not in focus:
                report["unexpected"] += 1
            continue
        matched_old.add(i)
        ch = diff_slide(snapshot(so[i - 1]), snapshot(s))
        moved = f" (was slide {i})" if i != j else ""
        if ch or moved:
            exp = j in focus
            report["slides"].append({"slide": j, "status": "changed" + moved, "expected": exp, "changes": ch})
            if ch and not exp:
                report["unexpected"] += 1
    for i in set(range(1, len(so) + 1)) - matched_old:
        report["slides"].append({"slide": i, "status": "deleted (old numbering)", "expected": False, "changes": []})
        report["unexpected"] += 1

    if args.json:
        print(json.dumps(report, indent=2))
    else:
        if not report["slides"]:
            print("No differences.")
        for r in report["slides"]:
            tag = "expected" if r["expected"] else "UNEXPECTED"
            print(f"Slide {r['slide']}: {r['status']} [{tag}]")
            for c in r["changes"]:
                print(f"   - {c}")
        print(f"\n{report['unexpected']} slide(s) with unexpected changes.")
    sys.exit(2 if report["unexpected"] else 0)


if __name__ == "__main__":
    main()
