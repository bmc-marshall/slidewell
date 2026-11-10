"""Find a font file for measuring text, on Linux, macOS or Windows.

Text has to be measured with the real font to wrap and check it correctly. This uses fontconfig (fc-match) when
it's installed, and otherwise searches the standard font folders of each operating system by the fonts' own names.
Office fonts that aren't installed fall back to their metric-compatible stand-ins (same widths), e.g.
Calibri -> Carlito, Cambria -> Caladea, Arial -> Liberation Sans.
"""
import os
import shutil
import subprocess
import sys
from functools import lru_cache

STAND_INS = {
    "arial": ["Liberation Sans", "Arimo", "Helvetica", "Helvetica Neue"],
    "helvetica": ["Arial", "Liberation Sans"],
    "calibri": ["Carlito"],
    "cambria": ["Caladea"],
    "times new roman": ["Liberation Serif", "Tinos", "Times"],
    "courier new": ["Liberation Mono", "Cousine", "Courier"],
}
FALLBACKS = ["Arial", "Liberation Sans", "DejaVu Sans", "Helvetica", "Segoe UI"]


def _font_dirs():
    home = os.path.expanduser("~")
    if sys.platform == "darwin":
        return ["/System/Library/Fonts", "/System/Library/Fonts/Supplemental", "/Library/Fonts", f"{home}/Library/Fonts"]
    if os.name == "nt":
        windir = os.environ.get("WINDIR", r"C:\Windows")
        local = os.environ.get("LOCALAPPDATA", "")
        return [os.path.join(windir, "Fonts"), os.path.join(local, "Microsoft", "Windows", "Fonts")]
    return ["/usr/share/fonts", "/usr/local/share/fonts", f"{home}/.local/share/fonts", f"{home}/.fonts"]


@lru_cache(maxsize=1)
def _index():
    """{(family_lower, is_bold): path} for every font file in the standard folders."""
    from PIL import ImageFont
    idx = {}
    for d in _font_dirs():
        for root, _, files in os.walk(d):
            for f in files:
                if not f.lower().endswith((".ttf", ".otf", ".ttc")):
                    continue
                p = os.path.join(root, f)
                try:
                    fam, style = ImageFont.truetype(p, 10).getname()
                except Exception:
                    continue
                style = (style or "").lower()
                if "italic" in style or "oblique" in style:
                    continue
                bold = "bold" in style or "black" in style or "heavy" in style
                key = (fam.lower(), bold)
                # prefer the plain Regular / Bold face over Light, Medium, Condensed and so on
                plain = style in ("regular", "bold", "roman", "book", "normal")
                if key not in idx or plain:
                    idx[key] = p
    return idx


def _fc_match(family, bold):
    if not shutil.which("fc-match"):
        return None, None
    try:
        out = subprocess.run(["fc-match", "-f", "%{file}|%{family}", family + (":bold" if bold else "")],
                             capture_output=True, text=True, timeout=10).stdout
        path, fam = out.split("|", 1)
        return path, fam.split(",")[0]
    except Exception:
        return None, None


@lru_cache(maxsize=None)
def find_font(family, bold=False):
    """(path, family actually used). Exact family first, then a metric-compatible stand-in, then a common sans."""
    path, fam = _fc_match(family, bold)
    if path:
        return path, fam
    idx = _index()
    for name in [family] + STAND_INS.get(family.lower(), []) + FALLBACKS:
        for b in ([bold, not bold] if bold else [False]):
            p = idx.get((name.lower(), b))
            if p:
                return p, name
    return None, None


def font_path(family, bold=False):
    """A path PIL can load. Raises a clear error if no usable font exists at all."""
    p, _ = find_font(family, bold)
    if not p:
        raise RuntimeError(f"No font found for {family!r} (and no fallback). Install it, or Liberation Sans / DejaVu Sans.")
    return p
