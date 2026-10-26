"""Build data set 1 (data/cylinder_vort_f32.npz) from the IBPM cylinder wake data.

Source: DATA/FLUIDS/CYLINDER_ALL.mat (UALL, VALL, VORTALL) in http://dmdbook.com/DATA.zip
(Kutz, Brunton, Brunton, Proctor, Dynamic Mode Decomposition, SIAM 2016; the same IBPM run
as the Brunton & Kutz textbook data). Re = 100, cylinder diameter D = 1 at the origin.

Grid and time step come from DATA/FLUIDS/ibpm_scripts/cylinder05.sh in the same zip:
-nx 450 -ny 200 -length 9 -xoffset -1 -yoffset -2 -dt .02 -tecplot 10.
The 449 x 199 interior nodes give x = -1 + 0.02 (i + 1), y = -2 + 0.02 (j + 1).
One snapshot every 10 steps gives dt = 0.2 (units of D / U_inf), so frequency = St.

Each column of the .mat arrays is one snapshot stored with y running fastest
(textbook code: np.reshape(col, (449, 199)).T is the (ny, nx) image).

Usage: python scripts/prepare_cylinder.py [--stride 2] [--keep-zip]
The default stride 2 gives a 100 x 225 grid (36 MB). Stride 1 gives 142 MB, over the
100 MB per-file limit of GitHub.
"""
import argparse
import io
import os
import shutil
import urllib.request
import zipfile

import numpy as np
import scipy.io as sio

URL = "http://dmdbook.com/DATA.zip"
MEMBER = "DATA/FLUIDS/CYLINDER_ALL.mat"
NX, NY = 449, 199
DX = 0.02
DT_SOLVER, SNAP_EVERY = 0.02, 10
N_SNAP = 151  # the .mat file has one extra test snapshot (UEXTRA etc.); keep 151

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def fetch(cache_dir):
    os.makedirs(cache_dir, exist_ok=True)
    zpath = os.path.join(cache_dir, "DMD_DATA.zip")
    if not os.path.exists(zpath):
        print("downloading", URL)
        with urllib.request.urlopen(URL) as r, open(zpath, "wb") as f:
            shutil.copyfileobj(r, f, length=1 << 20)
    with zipfile.ZipFile(zpath) as z:
        return sio.loadmat(io.BytesIO(z.read(MEMBER))), zpath


def to_fields(A):
    """(449*199, m) columns -> (ny, nx, m) float32 images."""
    m = A.shape[1]
    return np.ascontiguousarray(A.reshape(NX, NY, m).transpose(1, 0, 2)).astype(np.float32)


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--stride", type=int, default=2, help="keep every stride-th grid point")
    p.add_argument("--out", default=os.path.join(ROOT, "data", "cylinder_vort_f32.npz"))
    p.add_argument("--keep-zip", action="store_true")
    a = p.parse_args()

    cache = os.path.join(ROOT, "data", "_download")
    mat, zpath = fetch(cache)
    x = -1.0 + DX * (np.arange(NX) + 1)
    y = -2.0 + DX * (np.arange(NY) + 1)
    ix = np.arange(NX)[:: a.stride]
    jy = np.arange(NY)[:: a.stride]

    out = {}
    for key, name in [("VORTALL", "vort"), ("UALL", "u"), ("VALL", "v")]:
        f = to_fields(mat[key][:, :N_SNAP])
        out[name] = f[np.ix_(jy, ix)]
    out["x"] = x[ix].astype(np.float32)
    out["y"] = y[jy].astype(np.float32)
    out["dt"] = np.float32(DT_SOLVER * SNAP_EVERY)
    out["grid_stride"] = np.int32(a.stride)  # stride relative to the 449 x 199 DNS grid
    out["body"] = np.array([0.0, 0.0, 0.5], np.float32)  # cylinder centre x, y and radius
    out["meta"] = np.array(
        f"source={URL}:{MEMBER}; Re=100; D=1; dt=0.2 (solver dt 0.02 x 10); "
        f"stride={a.stride}; snapshots={N_SNAP}; layout=(ny, nx, m) float32"
    )
    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    np.savez_compressed(a.out, **out)
    for k in ("vort", "u", "v"):
        print(k, out[k].shape, out[k].dtype)
    print("x", out["x"][[0, -1]], "y", out["y"][[0, -1]], "dt", out["dt"])
    print(f"wrote {a.out}  {os.path.getsize(a.out) / 1e6:.1f} MB")
    if not a.keep_zip:
        os.remove(zpath)


if __name__ == "__main__":
    main()
