"""Crop figures from the thesis PDF (Liu 2022) into figures/thesis/*.png.

Re-runnable: crop boxes are constants (PDF points, 72 per inch, origin top-left).
Pages are 1-based PDF pages (PDF page = printed page + 4).
Figures in this PDF are stacks of thin embedded image strips, so no single embedded
image covers a whole figure; every figure is cropped from a 220 dpi render.
Captions are transcribed by hand from the PDF, because text extraction garbles the math.
Usage: python scripts/extract_thesis_figs.py
Outputs: figures/thesis/T*.png, figures/thesis/README.md, and figures/thesis/_work/
(contact sheet and the 400 dpi zoom of Fig. 3.12 (c); git-ignored).
"""
from pathlib import Path
import fitz
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent.parent
PDF = ROOT / "docs" / "Honbun-9040.pdf"
OUT = ROOT / "figures" / "thesis"
WORK = OUT / "_work"
DPI = 220
S = DPI / 72.0

# id: (figure, pdf_page, crop box x0,y0,x1,y1 in points, caption box in points, slides)
FIGS = {
 "T0":   ("Fig. 4.1",  92, (92, 363, 486, 716),  (184, 720, 397, 732), "8"),
 "T1a":  ("Fig. 4.12", 112, (66, 348, 512, 656), (70, 665, 510, 709),  "19"),
 "T1b":  ("Fig. 4.14", 115, (120, 102, 493, 677), (87, 683, 528, 735), "15, 19"),
 "T2a":  ("Fig. 4.15", 117, (83, 102, 529, 249), (87, 252, 528, 295),  "—"),
 "T2b":  ("Fig. 4.16", 117, (93, 295, 520, 705), (87, 706, 525, 743),  "15"),
 "T3":   ("Fig. 3.12",  52, (66, 102, 512, 252), (70, 255, 510, 303),  "19"),
 "T3b":  ("Fig. 3.23",  65, (83, 336, 529, 640), (87, 644, 528, 687),  "—"),
 "T4a":  ("Fig. 3.31",  78, (66, 287, 512, 506), (81, 516, 444, 540),  "35"),
 "T4b":  ("Fig. 3.33",  81, (83, 102, 529, 247), (87, 249, 528, 281),  "35"),
 "T5a_1":("Fig. 3.2",   38, (66, 102, 512, 291), (70, 294, 408, 319),  "—"),
 "T5a_2":("Fig. 3.3",   38, (66, 323, 512, 480), (70, 482, 501, 506),  "32"),
 "T5b":  ("Fig. 3.8",   45, (83, 300, 529, 621), (87, 642, 528, 677),  "32"),
 "T6":   ("Fig. 3.1",   36, (78, 102, 500, 531), (70, 534, 510, 570),  "32"),
 "T7":   ("Fig. 5.20", 150, (93, 102, 484, 696), (70, 699, 510, 733),  "4, 6, 12, 25"),
 "T8":   ("Fig. 3.9",   47, (145, 102, 468, 297), (87, 300, 528, 331), "19"),
}
CAPTIONS = {
    "T0": "Figure 4.1 Reshaping procedures of data matrices",
    "T1a": "Figure 4.12 TKE contribution of modes and the convergence of modes from POD analysis by two kinds of "
           "data sets: (a) (c), the axial velocity w and the transversal velocity v; (b) (d), the synthesis (in-plane) "
           "velocity V_yz. Sw = 1.1, θ = 8°, configuration 1, Re = 5.25 × 10⁴.",
    "T1b": "Figure 4.14 POD modes of the synthesis (in-plane) velocity fluctuation Ṽ_yz with their Power spectral "
           "density of time coefficient. Dominant frequency components are notated. Sw = 1.1, θ = 8°, "
           "configuration 1, Re = 5.25 × 10⁴.",
    "T2a": "Figure 4.15 The relationship between DMD modes amplitude or modes' growth rate and frequency. The frequency "
           "components being focused are marked in circle. Sw = 1.1, θ = 8°, configuration 1, Re = 5.25 × 10⁴.",
    "T2b": "Figure 4.16 DMD spectrum with the spatial modes in certain frequency being focused. Sw = 1.1, θ = 8°, "
           "configuration 1, Re = 5.25 × 10⁴.",
    "T3": "Figure 3.12 Spatial distribution of the instantaneous vortex centers and corresponding probability density "
          "function of the vortex center radius at the plane z/D₁ = 3. Sw = 1.1, Re = 5.25 × 10⁴, θ = 2°, "
          "configuration 1.",
    "T3b": "Figure 3.23 The evolution of vortex trajectory radius r_v at pipe configuration 1 and 2 with diffuser "
           "angle θ = 0°, 2°, 4°",
    "T4a": "Figure 3.31 Simplified theoretical convective pressure fluctuation model",
    "T4b": "Figure 3.33 Theoretical curve C_p*(r*) with all experimental cases data. (b): enlarged view",
    "T5a_1": "Figure 3.2 Overview of the air draft-tube test apparatus",
    "T5a_2": "Figure 3.3 Layout of unsteady pressure taps and dimensions of the upstream pipe and the test section",
    "T5b": "Figure 3.8 PIV measurements setup, position 1: tangential velocity measurements, position 2: axial "
           "velocity measurements",
    "T6": "Figure 3.1 Simplified representation of the helical precessing vortex filament with vortex core limit lines "
          "in the conical diffuser",
    "T7": "Figure 5.20 Snapshots of the vortex rope periodically switching between the single helix mode and double "
          "helix mode starts from two different instants t₁ and t₂. The vortex rope is visualized by iso-surfaces of "
          "pressure. Blue: -400 Pa, light blue: -200 Pa, transparent light blue: -80 Pa.",
    "T8": "Figure 3.9 Synchronized camera output signal (red) and the pressure signal (black). The precession period "
          "of PVC is shown as T_PVC, Sw = 1.1, Re = 5.25 × 10⁴, θ = 4°, configuration 1.",
}

# r* read from Fig. 3.12 (c) on the 400 dpi zoom (_work/T3_pdf_zoom.png), for GIF 14a.
R_STAR = r"""**r\* = 0.13**（r_v / R_sec。図から読んだ値は 0.128 ± 0.004）

- 読んだ図：T3 = Fig. 3.12 (c)「Probability Density Function of the vortex trajectory radius」、PDF p.52（印刷 p.48）。条件は z/D₁ = 3, Sw = 1.1, θ = 2°。
- 方法：panel (c) を 400 dpi で画像化し、画素を測った。
  - 横軸の目盛り：0 → 画素 x = 119、0.1 → 262、0.2 → 405、0.3 → 549、0.4 → 692（0.1 あたり 143.3 画素）。
  - 赤い確率密度曲線の最上点：画素 x = 297〜307（中央 302）。点線「Mean value r̄_v」も x = 302。
  - r\* = (302 − 119) / 1433 = 0.128。頂点が平らな幅 ±5 画素 → ±0.004。
- 確率密度のピークと平均値の点線が同じ位置にあるので、ピーク値と平均値のどちらを使っても 0.13。
- 読み取れないときの既定値 0.15 は使っていない。
"""

# Probability-density plot of Fig. 3.12 (p.52), for the 400 dpi zoom. Points.
ZOOM_BOX = (379, 102, 512, 229)  # panel (c), probability density of r_v


def crop(doc, page_no, box, dpi=DPI):
    page = doc[page_no - 1]
    pix = page.get_pixmap(dpi=dpi, clip=fitz.Rect(*box))
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    WORK.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(PDF))
    rows = []
    for fid, (fig, pg, box, cbox, slides) in FIGS.items():
        im = crop(doc, pg, box)
        im.save(OUT / f"{fid}.png")
        px = tuple(round(v * S) for v in box)
        rows.append((fid, fig, pg, box, px, im.size, CAPTIONS[fid], slides))
    write_readme(rows)
    contact_sheet()
    crop(doc, 52, ZOOM_BOX, dpi=400).save(WORK / "T3_pdf_zoom.png")


def write_readme(rows):
    L = []
    L.append("# 博論からの切り出し図（figures/thesis）\n")
    L.append("出典：劉志豪, 早稲田大学博士論文, 2022年2月「管路内旋回流れによる渦心振れ回り現象に関する研究」（`docs/Honbun-9040.pdf`）からの切り出し。")
    L.append("**再配布不可。公開リポジトリへ push しない。**\n")
    L.append("スライド上の出典表記：「劉 (2022) 博士論文 Fig. x.x」\n")
    L.append("生成：`python scripts/extract_thesis_figs.py`（切り出し枠はスクリプト内の定数）。印刷ページ = PDFページ - 4。")
    L.append("方法はいずれも「220 dpi で描画して切り出し」。埋め込み画像は細い帯状の画像を縦に並べたもので、1枚で図全体を覆うものは無かったため。")
    L.append("切り出し枠の px は 220 dpi のページ全体画像（左上原点）での座標、pt は PDF ポイント（1 pt = 1/72 inch）。")
    L.append("T5a は Fig. 3.2 と Fig. 3.3 の間にキャプション（Fig. 3.2 の説明文）が入るため、T5a_1（Fig. 3.2）と T5a_2（Fig. 3.3）の2枚に分けた。")
    L.append("slides 欄はデッキのページ番号（図を表示する、または図番号を本文で引用するページ）。「—」は今のデッキで使わない図。\n")
    L.append("| ID | Figure | PDF page | printed page | method | crop box px @220dpi (x0,y0,x1,y1) | crop box pt | output size px | caption (transcribed from the PDF) | slides |")
    L.append("|---|---|---|---|---|---|---|---|---|---|")
    for fid, fig, pg, box, px, size, cap, slides in rows:
        L.append(f"| {fid} | {fig} | {pg} | {pg-4} | render 220 dpi | {px[0]},{px[1]},{px[2]},{px[3]} | "
                 f"{box[0]},{box[1]},{box[2]},{box[3]} | {size[0]}x{size[1]} | {cap.replace('|','/')} | {slides} |")
    L.append("\n## r* (for GIF 14a, p.19)\n")
    L.append(R_STAR)
    (OUT / "README.md").write_text("\n".join(L), encoding="utf-8")


def contact_sheet():
    from PIL import ImageFont
    names = list(FIGS)
    cols, W = 3, 2400
    cell_w = W // cols
    try:
        font = ImageFont.truetype("arial.ttf", 28)
    except OSError:
        font = ImageFont.load_default()
    tiles = []
    for n in names:
        im = Image.open(OUT / f"{n}.png").convert("RGB")
        sc = (cell_w - 20) / im.width
        im = im.resize((cell_w - 20, max(1, int(im.height * sc))))
        tiles.append((n, im))
    rows = [tiles[i:i + cols] for i in range(0, len(tiles), cols)]
    heights = [max(t[1].height for t in r) + 44 for r in rows]
    sheet = Image.new("RGB", (W, sum(heights)), "white")
    d = ImageDraw.Draw(sheet)
    y = 0
    for r, h in zip(rows, heights):
        for i, (n, im) in enumerate(r):
            x = i * cell_w + 10
            d.text((x, y + 4), n, fill="red", font=font)
            sheet.paste(im, (x, y + 40))
            d.rectangle([x - 1, y + 39, x + im.width, y + 40 + im.height], outline="gray")
        y += h
    sheet.save(WORK / "thesis_contact.png")


if __name__ == "__main__":
    main()
