"""Builds docs/images/social-preview.png (1280x640) and docs/images/hero.png (2x) from five showcase slides.
Run after rendering the slides: python slidewell/scripts/render.py docs/showcase/showcase.pptx --out RENDER --dpi 200
then: python tools/make_advert.py RENDER"""
import subprocess
import sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFilter, ImageFont

ROOT = Path(__file__).resolve().parents[1]
SRC = Path(sys.argv[1])
LAYOUT = (14, 18, 7, 11, 19)                 # big, small top, small bottom, bottom-left, bottom-right
PITCH = ["Consulting-grade", "PowerPoint decks", "with Claude"]
BENEFITS = ["Storyline first.", "Varied designs, quickly.", "Every slide checked."]


def font(style, size):
    return ImageFont.truetype(subprocess.check_output(["fc-match", "-f", "%{file}", "Inter:" + style]).decode(), size)


def build(W, H, k, path):
    out = Image.new("RGB", (W, H), (18, 58, 51)); d = ImageDraw.Draw(out)
    d.text((int(52 * k), int(128 * k)), "Slidewell", fill=(255, 255, 255), font=font("bold", int(66 * k)))
    d.rectangle([int(54 * k), int(218 * k), int(132 * k), int(224 * k)], fill=(111, 211, 168))
    y = int(248 * k)
    for line in PITCH:
        d.text((int(52 * k), y), line, fill=(225, 240, 233), font=font("semibold", int(31 * k))); y += int(42 * k)
    y += int(20 * k)
    for line in BENEFITS:
        d.text((int(52 * k), y), line, fill=(156, 216, 192), font=font("regular", int(21 * k))); y += int(30 * k)
    d.text((int(52 * k), H - int(58 * k)), "github.com/bmc-marshall/slidewell", fill=(111, 211, 168), font=font("medium", int(18 * k)))

    def put(n, x, y, w):
        im = Image.open(SRC / f"slide-{n:02d}.png").convert("RGB"); h = int(im.height * w / im.width)
        im = im.resize((w, h), Image.LANCZOS)
        sh = Image.new("RGBA", (w + int(24 * k), h + int(24 * k)), (0, 0, 0, 0))
        ImageDraw.Draw(sh).rectangle([int(12 * k), int(16 * k), w + int(12 * k), h + int(16 * k)], fill=(0, 0, 0, 110))
        sh = sh.filter(ImageFilter.GaussianBlur(int(8 * k)))
        out.paste(sh, (x - int(12 * k), y - int(12 * k)), sh); out.paste(im, (x, y))
        ImageDraw.Draw(out).rectangle([x, y, x + w - 1, y + h - 1], outline=(70, 110, 100), width=max(1, int(k)))

    g = int(14 * k); X0 = int(455 * k); R = W - int(36 * k)
    sm = int(250 * k); bg = R - X0 - sm - g; hb = int(bg * 9 / 16); hs = int(sm * 9 / 16)
    bw = (R - X0 - g) // 2; hb2 = int(bw * 9 / 16); y0 = (H - (hb + g + hb2)) // 2
    big, s1, s2, bl, br = LAYOUT
    put(big, X0, y0, bg); put(s1, X0 + bg + g, y0, sm); put(s2, X0 + bg + g, y0 + hb - hs, sm)
    put(bl, X0, y0 + hb + g, bw); put(br, X0 + bw + g, y0 + hb + g, bw)
    out.save(path, optimize=True)


build(1280, 640, 1, ROOT / "docs/images/social-preview.png")
build(2560, 1280, 2, ROOT / "docs/images/hero.png")
print("done")
