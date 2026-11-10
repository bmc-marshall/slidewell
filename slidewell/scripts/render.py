#!/usr/bin/env python3
"""Render a .pptx to slide images plus one contact sheet, for visual QA.

Usage:
    python render.py deck.pptx [--out DIR] [--dpi 110] [--slides 3,5-7] [--cols 4]

Writes DIR/slide-NN.png for each slide and DIR/contact-sheet.png showing all
rendered slides with their numbers. Prints the paths.

Rendering uses LibreOffice. Fonts missing from this machine are substituted,
which changes text widths: if text looks clipped or crowded, check the text
box size against its content before "fixing" it (see SKILL.md, Theme > Fonts).
"""
import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile

from PIL import Image, ImageDraw, ImageFont


def parse_slides(spec):
    if not spec:
        return None
    out = set()
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-")
            out.update(range(int(a), int(b) + 1))
        elif part:
            out.add(int(part))
    return out


def find_tools():
    """Return (soffice, pdftoppm) paths, or exit with code 3 and a plain message if either is missing."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    for mac in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",):
        if not soffice and os.path.exists(mac):
            soffice = mac
    pdftoppm = shutil.which("pdftoppm")
    missing = [n for n, v in (("LibreOffice (soffice)", soffice), ("Poppler (pdftoppm)", pdftoppm)) if not v]
    if missing:
        print("VISUAL REVIEW UNAVAILABLE: " + " and ".join(missing) + " not found.\n"
              "Run check_deck.py as usual, deliver, and tell the user the visual review was skipped.\n"
              "To enable it: macOS `brew install --cask libreoffice && brew install poppler`; "
              "Ubuntu/Debian `sudo apt install libreoffice poppler-utils`; "
              "Windows: install LibreOffice and Poppler and add both to PATH.", file=sys.stderr)
        sys.exit(3)
    return soffice, pdftoppm


SOFFICE = "soffice"


def to_pdf(pptx, workdir):
    profile = tempfile.mkdtemp(prefix="lo-profile-")
    cmd = [
        SOFFICE, f"-env:UserInstallation=file://{profile}", "--headless",
        "--convert-to", "pdf", "--outdir", workdir, pptx,
    ]
    try:
        subprocess.run(cmd, check=True, timeout=300,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    pdf = os.path.join(workdir, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
    if not os.path.exists(pdf):
        sys.exit("Rendering failed: LibreOffice produced no PDF.")
    return pdf


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx")
    ap.add_argument("--out", default=None, help="output folder (default: <deck>_render)")
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--slides", default=None, help="e.g. 3,5-7 (default: all)")
    ap.add_argument("--cols", type=int, default=4, help="contact sheet columns")
    args = ap.parse_args()

    global SOFFICE
    SOFFICE, pdftoppm = find_tools()
    pptx = os.path.abspath(args.pptx)
    out = args.out or os.path.splitext(pptx)[0] + "_render"
    os.makedirs(out, exist_ok=True)
    for old in glob.glob(os.path.join(out, "slide-*.png")):
        os.remove(old)

    with tempfile.TemporaryDirectory() as tmp:
        pdf = to_pdf(pptx, tmp)
        subprocess.run([pdftoppm, "-r", str(args.dpi), "-png", pdf, os.path.join(tmp, "p")], check=True)
        pages = sorted(glob.glob(os.path.join(tmp, "p-*.png")),
                       key=lambda p: int(p.rsplit("-", 1)[1].split(".")[0]))
        wanted = parse_slides(args.slides)
        paths = []
        for i, p in enumerate(pages, 1):
            if wanted and i not in wanted:
                continue
            dest = os.path.join(out, f"slide-{i:02d}.png")
            shutil.move(p, dest)
            paths.append((i, dest))

    if not paths:
        sys.exit("No slides rendered.")

    # Contact sheet
    thumbs = []
    for i, p in paths:
        im = Image.open(p).convert("RGB")
        im.thumbnail((480, 480))
        thumbs.append((i, im))
    tw = max(t.width for _, t in thumbs)
    th = max(t.height for _, t in thumbs)
    cols = min(args.cols, len(thumbs))
    rows = (len(thumbs) + cols - 1) // cols
    pad, label_h = 12, 22
    sheet = Image.new("RGB", (cols * (tw + pad) + pad, rows * (th + label_h + pad) + pad), "white")
    draw = ImageDraw.Draw(sheet)
    try:
        font = ImageFont.truetype("DejaVuSans.ttf", 15)
    except OSError:
        font = ImageFont.load_default()
    for k, (i, t) in enumerate(thumbs):
        x = pad + (k % cols) * (tw + pad)
        y = pad + (k // cols) * (th + label_h + pad)
        draw.text((x, y), f"Slide {i}", fill="black", font=font)
        sheet.paste(t, (x, y + label_h))
        draw.rectangle([x - 1, y + label_h - 1, x + t.width, y + label_h + t.height], outline="#999999")
    sheet_path = os.path.join(out, "contact-sheet.png")
    sheet.save(sheet_path)

    for i, p in paths:
        print(p)
    print(sheet_path)


if __name__ == "__main__":
    main()
