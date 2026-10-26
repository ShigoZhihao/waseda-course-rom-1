"""Crop the open-access paper figures for the answer slides into figures/papers/*.png.

The PDFs are not in the repository. Download them from the DOI pages (listed in
figures/papers/README.md) and pass their paths.
Crop boxes are pixels of a 200 dpi render (origin top-left). Pages are 0-based PDF pages.
Captions are left out of the crops; the slides carry their own citation line.
Usage: python scripts/extract_paper_figs.py --nicolet N.pdf --alligne A.pdf --favrel F.pdf
"""
import argparse
from pathlib import Path
import fitz
import numpy as np
from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "figures" / "papers"
DPI = 200

# name: (paper key, 0-based page, crop box x0, y0, x1, y1 in 200 dpi pixels)
FIGS = {
    "nicolet2025_fig3": ("nicolet", 4, (292, 486, 1294, 871)),
    "nicolet2025_fig5": ("nicolet", 6, (532, 308, 1120, 539)),
    "nicolet2025_fig7": ("nicolet", 7, (231, 936, 1439, 1140)),
    "alligne2016_fig4": ("alligne", 6, (343, 278, 1292, 773)),
    "favrel2017_fig5": ("favrel", 4, (253, 807, 1416, 1178)),
}


def main():
    ap = argparse.ArgumentParser()
    for key in ("nicolet", "alligne", "favrel"):
        ap.add_argument(f"--{key}", required=True, type=Path)
    args = ap.parse_args()
    OUT.mkdir(parents=True, exist_ok=True)
    for name, (key, page, (x0, y0, x1, y1)) in FIGS.items():
        pix = fitz.open(getattr(args, key))[page].get_pixmap(dpi=DPI)
        img = np.frombuffer(pix.samples, np.uint8).reshape(pix.h, pix.w, pix.n)[:, :, :3]
        Image.fromarray(img[y0:y1, x0:x1]).save(OUT / f"{name}.png")
        print(f"{name}.png  {x1 - x0}x{y1 - y0}")


if __name__ == "__main__":
    main()
