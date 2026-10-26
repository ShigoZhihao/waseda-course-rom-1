"""Build data set 2 (data/cylinder_pivlike_f32.npz): PIV-like degraded copy of data set 1.

Operations: grid subsampling, Gaussian noise, random dropouts,
ellipse-shaped gaps (laser shadows / reflections) and the cylinder body mask.
--stride is relative to the 449 x 199 DNS grid, so --stride 4 gives about 112 x 50 points
even if data set 1 was itself subsampled.
The generation parameters are saved in the npz as `meta` (a JSON string).

Usage: python scripts/make_pivlike.py --stride 4 --noise 0.05 --dropout 0.05 --blobs 3 --tstride 1 --seed 0
"""
import argparse
import json
import os
import sys

import numpy as np

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from rom_utils import degrade_to_piv  # noqa: E402


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--stride", type=int, default=4)
    p.add_argument("--noise", type=float, default=0.05)
    p.add_argument("--dropout", type=float, default=0.05)
    p.add_argument("--blobs", type=int, default=3)
    p.add_argument("--tstride", type=int, default=1)
    p.add_argument("--seed", type=int, default=0)
    p.add_argument("--src", default=os.path.join(ROOT, "data", "cylinder_vort_f32.npz"))
    p.add_argument("--out", default=os.path.join(ROOT, "data", "cylinder_pivlike_f32.npz"))
    a = p.parse_args()

    d = np.load(a.src)
    base = int(d["grid_stride"])
    if a.stride % base:
        sys.exit(f"--stride {a.stride} is not a multiple of the data set 1 stride {base}")
    s = a.stride // base

    xc, yc, rad = d["body"]
    X, Y = np.meshgrid(d["x"], d["y"])
    body = (X - xc) ** 2 + (Y - yc) ** 2 <= rad**2

    u2, v2, mask, _ = degrade_to_piv(d["u"], d["v"], stride=s, noise=a.noise, dropout=a.dropout,
                                     blobs=a.blobs, tstride=a.tstride, seed=a.seed, body=body)
    x2, y2 = d["x"][::s], d["y"][::s]
    dt2 = np.float32(d["dt"] * a.tstride)

    ny, nx, m = u2.shape
    gap = np.isnan(u2)
    stats = {
        "grid": [ny, nx],
        "points_ratio_to_dns": ny * nx / (449 * 199),
        "missing_fraction_all": float(gap.mean()),
        "missing_fraction_outside_body": float(gap[~mask].mean()),
        "body_fraction": float(mask.mean()),
    }
    meta = {"args": {k: getattr(a, k) for k in ("stride", "noise", "dropout", "blobs", "tstride", "seed")},
            "noise_definition": "std = noise * RMS of the velocity fluctuation (per component)",
            "blobs_definition": "per snapshot, `blobs` ellipses at random positions are NaN",
            "truth": "data set 1 u, v on the same grid: u[::s, ::s, ::tstride]",
            "stats": stats}
    np.savez_compressed(a.out, u=u2, v=v2, x=x2, y=y2, dt=dt2, mask=mask, meta=np.array(json.dumps(meta)))
    print(json.dumps(meta, indent=1))
    print("u", u2.shape, "v", v2.shape)
    print(f"wrote {a.out}  {os.path.getsize(a.out) / 1e6:.1f} MB")


if __name__ == "__main__":
    main()
