# 流体機械演習B 第4回 Reduced Order Modeling（ROM）：流れ場の低次元化と予測 — 演習

POD（固有直交分解）と DMD（動的モード分解）を、円柱後流のデータで試す演習。

## Colab で開く

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/ShigoZhihao/waseda-course-rom-1/blob/main/notebooks/01_pod_dmd.ipynb)

<https://colab.research.google.com/github/ShigoZhihao/waseda-course-rom-1/blob/main/notebooks/01_pod_dmd.ipynb>

- Colab の無料枠（GPU なし）で動作。追加の `pip install` は不要。
- 最初のセルで、このリポジトリの clone と `rom_utils.py` の読み込み。
- 全セルの実行時間は 2 分以内。

## 中身

| パス | 内容 |
|---|---|
| `notebooks/01_pod_dmd.ipynb` | 演習の本体（セル 0〜15） |
| `rom_utils.py` | `pod()`, `pod_snapshots()`, `exact_dmd()`, `gappy_fill()`, `degrade_to_piv()`, 描画関数 |
| `data/cylinder_vort_f32.npz` | データ①：円柱後流 DNS（渦度 `vort`、速度 `u`, `v`、座標 `x`, `y`、時間刻み `dt`） |
| `data/cylinder_pivlike_f32.npz` | データ②：①を PIV 風に劣化させたもの（`u`, `v` に NaN、`mask`、生成条件 `meta`） |
| `scripts/` | データとアニメーションと図を作るスクリプト |
| `figures/static/` | notebook が出力する図 |
| `figures/anim/` | スライド用アニメーション 12 本（GIF・MP4・最初のフレームの PNG）。`python scripts/make_animations.py` で作り直し可能（約 3 分） |
| `scripts/build_deck.py` | 講義スライド `deck/rom-session1.pptx` を作るスクリプト。`--pdf` で PDF も書き出し（Windows の PowerPoint が必要） |

スライドと PDF は博士論文の図を含むため、リポジトリの対象外（下の「博士論文の図について」）。
作り直しに必要なもの：`docs/template.pptx`、`figures/thesis/` の図、`figures/papers/` の図（取得手順は `figures/papers/README.md`）。

変数名は教科書のコード（[dynamicslab/databook_python](https://github.com/dynamicslab/databook_python)）に準拠：`X`, `U`, `S`, `VT`, `Atilde`, `Phi`, `lam`。

## データの出典

- 教科書：S. L. Brunton and J. N. Kutz, *Data-Driven Science and Engineering*, 2nd ed., Cambridge University Press, 2022（[databookuw.com](http://databookuw.com)）。
- 円柱後流（Re = 100）のデータは、教科書と同じ IBPM（immersed boundary projection method）の計算結果。
  速度 `UALL`, `VALL` と渦度 `VORTALL` を含む `CYLINDER_ALL.mat` の入手元は、DMD の本の配布データ
  [dmdbook.com/DATA.zip](http://dmdbook.com/DATA.zip)（`DATA/FLUIDS/`）。
  - J. N. Kutz, S. L. Brunton, B. W. Brunton, J. L. Proctor, *Dynamic Mode Decomposition*, SIAM, 2016.
  - 計算コード：[github.com/cwrowley/ibpm](https://github.com/cwrowley/ibpm)
- 格子と時間刻みの根拠は、同じ zip の `ibpm_scripts/cylinder05.sh`（`-nx 450 -ny 200 -length 9 -xoffset -1 -yoffset -2 -dt .02 -tecplot 10`）と
  `cylinder.geom`（直径 1 の円柱を原点に配置）。
  - スナップショットの間隔は 0.02 × 10 = 0.2（円柱直径 D と一様流速 U で無次元化した時間）。
  - D = U = 1 のため、周波数がそのままストローハル数 St = fD/U。

## データ①の作り方

```
python scripts/prepare_cylinder.py --stride 2
```

- `CYLINDER_ALL.mat` の 449×199 格子・151 スナップショットを、2 点に 1 点に間引いて 100×225 格子に変換（Δx = 0.04 D、円柱直径に 25 点）。
- float32 で保存（36 MB）。全格子のままでは 142 MB となり、GitHub の 1 ファイル上限 100 MB を超えるため。
- 配列の形は `(ny, nx, m)`。`X = vort.reshape(ny*nx, m)` で、列が 1 時刻のスナップショット行列。

## データ②（PIV 風の劣化データ）の作り方

```
python scripts/make_pivlike.py --stride 4 --noise 0.05 --dropout 0.05 --blobs 3 --tstride 1 --seed 0
```

①に次の操作を加えて作成。真値（①を同じ格子に間引いたもの）が手元にあるため、壊れ方を数値で比較可能。

| 操作 | 引数 | 内容 |
|---|---|---|
| 格子の間引き | `--stride 4` | 元の DNS 格子で 4 点に 1 点（50×113 格子）。PIV の判定窓を模擬 |
| ノイズ | `--noise 0.05` | ガウスノイズ。標準偏差は速度揺らぎの RMS の 5%（成分ごと） |
| ランダム欠損 | `--dropout 0.05` | 各時刻で 5% の点を NaN に置換 |
| 固まった欠損 | `--blobs 3` | 各時刻で楕円 3 個の領域を NaN に置換（レーザーの影・反射を模擬。位置は時刻ごとに乱数） |
| 物体マスク | — | 円柱の内側を NaN にし、`mask` に保存 |
| 時間の間引き | `--tstride 1` | 1 なら間引きなし。2 なら 2 点に 1 点 |

生成条件は npz の `meta`（JSON 文字列）に保存。notebook のセル10 で表示。

## 博士論文の図について

`figures/thesis/` の図は、博士論文（劉志豪, 早稲田大学, 2022）からの切り出し。**再配布なし。**
画像ファイルは `.gitignore` でリポジトリの対象外とし、スライドの中でだけ引用。
