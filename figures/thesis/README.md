# 博論からの切り出し図（figures/thesis）

出典：劉志豪, 早稲田大学博士論文, 2022年2月「管路内旋回流れによる渦心振れ回り現象に関する研究」（`docs/Honbun-9040.pdf`）からの切り出し。
**再配布不可。公開リポジトリへ push しない。**

スライド上の出典表記：「劉 (2022) 博士論文 Fig. x.x」

生成：`python scripts/extract_thesis_figs.py`（切り出し枠はスクリプト内の定数）。印刷ページ = PDFページ - 4。
方法はいずれも「220 dpi で描画して切り出し」。埋め込み画像は細い帯状の画像を縦に並べたもので、1枚で図全体を覆うものは無かったため。
切り出し枠の px は 220 dpi のページ全体画像（左上原点）での座標、pt は PDF ポイント（1 pt = 1/72 inch）。
T5a は Fig. 3.2 と Fig. 3.3 の間にキャプション（Fig. 3.2 の説明文）が入るため、T5a_1（Fig. 3.2）と T5a_2（Fig. 3.3）の2枚に分けた。

| ID | Figure | PDF page | printed page | method | crop box px @220dpi (x0,y0,x1,y1) | crop box pt | output size px | caption (transcribed from the PDF) | slides |
|---|---|---|---|---|---|---|---|---|---|
| T0 | Fig. 4.1 | 92 | 88 | render 220 dpi | 281,1109,1485,2188 | 92,363,486,716 | 1204x1079 | Figure 4.1 Reshaping procedures of data matrices | 6 |
| T1a | Fig. 4.12 | 112 | 108 | render 220 dpi | 202,1063,1564,2004 | 66,348,512,656 | 1364x942 | Figure 4.12 TKE contribution of modes and the convergence of modes from POD analysis by two kinds of data sets: (a) (c), the axial velocity w and the transversal velocity v; (b) (d), the synthesis (in-plane) velocity V_yz. Sw = 1.1, θ = 8°, configuration 1, Re = 5.25 × 10⁴. | 14 右上 |
| T1b | Fig. 4.14 | 115 | 111 | render 220 dpi | 367,312,1506,2069 | 120,102,493,677 | 1141x1758 | Figure 4.14 POD modes of the synthesis (in-plane) velocity fluctuation Ṽ_yz with their Power spectral density of time coefficient. Dominant frequency components are notated. Sw = 1.1, θ = 8°, configuration 1, Re = 5.25 × 10⁴. | 14 右下, 12 |
| T2a | Fig. 4.15 | 117 | 113 | render 220 dpi | 254,312,1616,761 | 83,102,529,249 | 1364x450 | Figure 4.15 The relationship between DMD modes amplitude or modes' growth rate and frequency. The frequency components being focused are marked in circle. Sw = 1.1, θ = 8°, configuration 1, Re = 5.25 × 10⁴. | 12 |
| T2b | Fig. 4.16 | 117 | 113 | render 220 dpi | 284,901,1589,2154 | 93,295,520,705 | 1305x1254 | Figure 4.16 DMD spectrum with the spatial modes in certain frequency being focused. Sw = 1.1, θ = 8°, configuration 1, Re = 5.25 × 10⁴. | 12, 11 |
| T3 | Fig. 3.12 | 52 | 48 | render 220 dpi | 202,312,1564,770 | 66,102,512,252 | 1364x459 | Figure 3.12 Spatial distribution of the instantaneous vortex centers and corresponding probability density function of the vortex center radius at the plane z/D₁ = 3. Sw = 1.1, Re = 5.25 × 10⁴, θ = 2°, configuration 1. | 14a, 10b |
| T3b | Fig. 3.23 | 65 | 61 | render 220 dpi | 254,1027,1616,1956 | 83,336,529,640 | 1364x930 | Figure 3.23 The evolution of vortex trajectory radius r_v at pipe configuration 1 and 2 with diffuser angle θ = 0°, 2°, 4° | 14a |
| T4a | Fig. 3.31 | 78 | 74 | render 220 dpi | 202,877,1564,1546 | 66,287,512,506 | 1364x671 | Figure 3.31 Simplified theoretical convective pressure fluctuation model | 25 (A) |
| T4b | Fig. 3.33 | 81 | 77 | render 220 dpi | 254,312,1616,755 | 83,102,529,247 | 1364x444 | Figure 3.33 Theoretical curve C_p*(r*) with all experimental cases data. (b): enlarged view | 25 (A) |
| T5a_1 | Fig. 3.2 | 38 | 34 | render 220 dpi | 202,312,1564,889 | 66,102,512,291 | 1364x579 | Figure 3.2 Overview of the air draft-tube test apparatus | 24 (A) |
| T5a_2 | Fig. 3.3 | 38 | 34 | render 220 dpi | 202,987,1564,1467 | 66,323,512,480 | 1364x481 | Figure 3.3 Layout of unsteady pressure taps and dimensions of the upstream pipe and the test section | 24 (A) |
| T5b | Fig. 3.8 | 45 | 41 | render 220 dpi | 254,917,1616,1897 | 83,300,529,621 | 1364x982 | Figure 3.8 PIV measurements setup, position 1: tangential velocity measurements, position 2: axial velocity measurements | 24 (A) |
| T6 | Fig. 3.1 | 36 | 32 | render 220 dpi | 238,312,1528,1622 | 78,102,500,531 | 1290x1312 | Figure 3.1 Simplified representation of the helical precessing vortex filament with vortex core limit lines in the conical diffuser | 24 (A), 14a |
| T7 | Fig. 5.20 | 150 | 146 | render 220 dpi | 284,312,1479,2127 | 93,102,484,696 | 1195x1816 | Figure 5.20 Snapshots of the vortex rope periodically switching between the single helix mode and double helix mode starts from two different instants t₁ and t₂. The vortex rope is visualized by iso-surfaces of pressure. Blue: -400 Pa, light blue: -200 Pa, transparent light blue: -80 Pa. | 10, 19 |
| T8 | Fig. 3.9 | 47 | 43 | render 220 dpi | 443,312,1430,907 | 145,102,468,297 | 987x597 | Figure 3.9 Synchronized camera output signal (red) and the pressure signal (black). The precession period of PVC is shown as T_PVC, Sw = 1.1, Re = 5.25 × 10⁴, θ = 4°, configuration 1. | 14, 14a |

## r* (for GIF 14a)

**r\* = 0.13**（r_v / R_sec。図から読んだ値は 0.128 ± 0.004）

- 読んだ図：T3 = Fig. 3.12 (c)「Probability Density Function of the vortex trajectory radius」、PDF p.52（印刷 p.48）。条件は z/D₁ = 3, Sw = 1.1, θ = 2°。
- 方法：panel (c) を 400 dpi で画像化し、画素を測った。
  - 横軸の目盛り：0 → 画素 x = 119、0.1 → 262、0.2 → 405、0.3 → 549、0.4 → 692（0.1 あたり 143.3 画素）。
  - 赤い確率密度曲線の最上点：画素 x = 297〜307（中央 302）。点線「Mean value r̄_v」も x = 302。
  - r\* = (302 − 119) / 1433 = 0.128。頂点が平らな幅 ±5 画素 → ±0.004。
- 確率密度のピークと平均値の点線が同じ位置にあるので、ピーク値と平均値のどちらを使っても 0.13。
- 読み取れないときの既定値 0.15 は使っていない。
