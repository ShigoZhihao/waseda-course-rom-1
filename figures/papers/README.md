# 論文の図（figures/papers）

思考実験の答え合わせ（p.34、p.36、p.37、p.39）で引用する、オープンアクセス論文の図。画像はリポジトリの対象外。
次の手順で取得。

## Zhang et al. 2026（p.39）

出典：L. Zhang, M. Ma, Y. Li, L. Kong, L. Xu, Z. Huang, B. Wang, "Reconstruction and Prediction of Three-Dimensional
Transient Flow Field in a Draft Tube of Francis Turbine Using Sparse Sensors and a Proper Orthogonal
Decomposition-Long Short-Term Memory Network", *Energies* 19(10):2300, 2026, DOI 10.3390/en19102300。
ライセンス：CC BY（出典の表示で転載可）。

| ファイル | 図 | 取得元 |
|---|---|---|
| `zhang2026_fig2.png` | Figure 2. Locations of the sensors | `https://mdpi-res.com/energies/energies-19-02300/article_deploy/html/images/energies-19-02300-g002.png` |
| `zhang2026_fig17.png` | Figure 17. True and predicted temporal coefficients of the first three POD modes under PL condition | `https://mdpi-res.com/energies/energies-19-02300/article_deploy/html/images/energies-19-02300-g017.png` |

```
curl -L -o figures/papers/zhang2026_fig2.png  https://mdpi-res.com/energies/energies-19-02300/article_deploy/html/images/energies-19-02300-g002.png
curl -L -o figures/papers/zhang2026_fig17.png https://mdpi-res.com/energies/energies-19-02300/article_deploy/html/images/energies-19-02300-g017.png
```

## 1D 水撃モデルと共振の図（p.34、p.36、p.37）

3本の論文の PDF を DOI のページから取得し、`scripts/extract_paper_figs.py` で切り出し。
図番号の下の説明文（キャプション）は切り出しに含めず、スライドの出典行で表示。

```
python scripts/extract_paper_figs.py --nicolet nicolet.pdf --alligne alligne.pdf --favrel favrel.pdf
```

| ファイル | 図 | 論文 | スライド |
|---|---|---|---|
| `nicolet2025_fig3.png` | Fig. 3：長さ dx の管要素と、T 字形の RLC 等価回路 | Nicolet 2025 | p.36 |
| `nicolet2025_fig7.png` | Fig. 7：貯水池・水圧管・水車・ドラフトチューブ（励振源つき）・放水路の系統 | Nicolet 2025 | p.36 |
| `nicolet2025_fig5.png` | Fig. 5：キャビテーションを含むドラフトチューブの等価回路（L_DT と C_DT） | Nicolet 2025 | p.37 |
| `alligne2016_fig4.png` | Fig. 4：ドラフトチューブ圧力の伝達関数 \|G\| と f/n（PL1、PL2） | Alligné 2016 | p.37 |
| `favrel2017_fig5.png` | Fig. 5：振れ回り周波数と固有振動数、圧力スペクトルの振幅（流量係数に対して） | Favrel 2017 | p.34 |

| 論文 | DOI | ライセンス |
|---|---|---|
| C. Nicolet, C. Landry, A. Béguin, S. Alligné, "Analytical approach for Francis turbine part load resonance risk assessment", *IOP Conf. Ser.: Earth Environ. Sci.* 1483:012036, 2025 | 10.1088/1755-1315/1483/1/012036 | CC BY 4.0 |
| S. Alligné, C. Nicolet, A. Béguin, C. Landry, J. Gomes, F. Avellan, "Hydroelectric System Response to Part Load Vortex Rope Excitation", *IOP Conf. Ser.: Earth Environ. Sci.* 49:052002, 2016 | 10.1088/1755-1315/49/5/052002 | CC BY 3.0 |
| A. Favrel, A. Müller, C. Landry, J. Gomes, K. Yamamoto, F. Avellan, "Dynamics of the precessing vortex rope and its interaction with the system at Francis turbines part load operating conditions", *J. Phys.: Conf. Ser.* 813:012023, 2017 | 10.1088/1742-6596/813/1/012023 | CC BY 3.0 |

切り出し範囲（200 dpi で描いたページの画素。ページ番号は 0 始まり）は `scripts/extract_paper_figs.py` の `FIGS` に記載。
