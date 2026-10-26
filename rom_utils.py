"""POD / DMD helpers for the ROM session 1 exercise.

Variable names follow the textbook code (Brunton & Kutz, dynamicslab/databook_python):
X (n x m, one column per snapshot), U, S, VT, Atilde, Phi, lam.
Only numpy, scipy and matplotlib are used.
"""
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap

# Theme colours of the slide template
DEEP_BLUE = "#000028"
CARD = "#1B2A4E"
ACC1 = "#009999"
ACC2 = "#00D7A0"
ACC3 = "#00BEDC"
ACC4 = "#0087BE"
ACC5 = "#00557C"
WARM = "#FFB04A"
LIGHT = "#F3F3F0"
THEME = [ACC4, WARM, ACC1, ACC2, ACC3, ACC5, DEEP_BLUE, LIGHT]

# Sequential map for magnitudes, diverging map for signed fields (negative = ACC4, positive = WARM)
CMAP_SEQ = LinearSegmentedColormap.from_list("rom_seq", [DEEP_BLUE, ACC4, ACC2, LIGHT])
CMAP_DIV = LinearSegmentedColormap.from_list("rom_div", [ACC4, LIGHT, WARM])


# ---------------------------------------------------------------- data

def load_cylinder(path):
    """Load data set 1. Fields are (ny, nx, m) float32 arrays."""
    d = np.load(path, allow_pickle=False)
    return {k: d[k] for k in d.files}


def to_X(field):
    """(ny, nx, m) -> X (n x m). One column = one snapshot."""
    ny, nx, m = field.shape
    return field.reshape(ny * nx, m)


def to_field(col, ny, nx):
    """One column of X -> (ny, nx) image."""
    return np.asarray(col).reshape(ny, nx)


# ---------------------------------------------------------------- POD

def pod(X, subtract_mean=True):
    """Thin SVD of the snapshot matrix. Returns U, S, VT, xbar."""
    X = np.asarray(X, dtype=np.float64)
    xbar = X.mean(axis=1, keepdims=True) if subtract_mean else np.zeros((X.shape[0], 1))
    U, S, VT = np.linalg.svd(X - xbar, full_matrices=False)
    return U, S, VT, xbar


def pod_snapshots(X):
    """Method of snapshots (Sirovich 1987): eigen-decompose the small m x m matrix X^T X.

    U = X V S^-1. Use it when n >> m. X must already have its mean removed if you want fluctuation modes.
    """
    X = np.asarray(X, dtype=np.float64)
    XtX = X.T @ X
    evals, V = np.linalg.eigh(XtX)
    idx = np.argsort(evals)[::-1]
    evals, V = evals[idx], V[:, idx]
    keep = evals > evals[0] * 1e-12
    S = np.sqrt(evals[keep])
    V = V[:, keep]
    U = X @ V / S
    return U, S, V.T


def energy(S):
    """Cumulative energy ratio sum(sigma_i^2) / sum(sigma^2)."""
    e = S**2
    return np.cumsum(e) / e.sum()


def rank_for_energy(S, level):
    """Smallest r whose cumulative energy reaches `level` (e.g. 0.95)."""
    return int(np.searchsorted(energy(S), level) + 1)


# ---------------------------------------------------------------- DMD

def exact_dmd(X, dt, r):
    """Exact DMD (Tu et al. 2014) of the snapshot sequence X (n x m).

    Returns lam (r,), Phi (n x r), b (r,), freq_hz (r,), growth (r,).
    freq_hz = arg(lam) / (2 pi dt), growth = ln|lam| / dt.
    """
    X = np.asarray(X, dtype=np.float64)
    X1, X2 = X[:, :-1], X[:, 1:]
    U, S, VT = np.linalg.svd(X1, full_matrices=False)
    U, S, V = U[:, :r], S[:r], VT[:r, :].T
    Atilde = U.T @ X2 @ V / S
    lam, W = np.linalg.eig(Atilde)
    Phi = X2 @ V / S @ W
    b = np.linalg.lstsq(Phi, X1[:, 0], rcond=None)[0]
    freq_hz = np.angle(lam) / (2 * np.pi * dt)
    growth = np.log(np.abs(lam)) / dt
    return lam, Phi, b, freq_hz, growth


def dmd_reconstruct(Phi, lam, b, k):
    """x_k = sum_i phi_i b_i lam_i^k for the time indices k."""
    k = np.atleast_1d(k)
    return np.real(Phi @ (b[:, None] * lam[:, None] ** k[None, :]))


# ---------------------------------------------------------------- gappy POD

def gappy_fill(X_nan, U_r, n_iter=1, callback=None):
    """Fill the NaN entries of X_nan (n x m) with gappy POD (Everson & Sirovich 1995).

    U_r: an int r (start from the temporal-mean fill) or an initial basis (n x r).
    Each iteration: remove the mean, fit r POD coefficients on the observed entries of each
    snapshot by least squares, put the reconstruction into the gaps, then recompute the basis.
    Rows that are NaN at every time (for example the body mask) stay NaN.
    callback(it, X_filled, S) is called after each iteration if given.
    """
    X = np.array(X_nan, dtype=np.float64)
    gap = np.isnan(X)
    dead = gap.all(axis=1)
    live = ~dead
    Xl, gl = X[live], gap[live]
    xbar = np.nanmean(Xl, axis=1, keepdims=True)
    Xl = np.where(gl, xbar, Xl)
    if np.isscalar(U_r):
        r = int(U_r)
        basis = None
    else:
        basis = np.asarray(U_r)[live]
        r = basis.shape[1]
    S = None
    for it in range(n_iter):
        xbar = Xl.mean(axis=1, keepdims=True)
        if basis is None or it > 0:
            Uf, S, _ = np.linalg.svd(Xl - xbar, full_matrices=False)
            basis = Uf[:, :r]
        for k in range(Xl.shape[1]):
            obs = ~gl[:, k]
            a = np.linalg.lstsq(basis[obs], Xl[obs, k] - xbar[obs, 0], rcond=None)[0]
            xk = xbar[:, 0] + basis @ a
            Xl[gl[:, k], k] = xk[gl[:, k]]
        if callback is not None:
            S = np.linalg.svd(Xl - Xl.mean(axis=1, keepdims=True), compute_uv=False)
            out = np.full_like(X, np.nan)
            out[live] = Xl
            callback(it, out, S)
    out = np.full_like(X, np.nan)
    out[live] = Xl
    return out


# ---------------------------------------------------------------- PIV-like degradation

def degrade_to_piv(u, v, stride=4, noise=0.05, dropout=0.05, blobs=3, tstride=1, seed=0,
                   body=None):
    """Turn clean DNS velocity fields into PIV-like data.

    u, v: (ny, nx, m) clean fields. body: optional (ny, nx) bool mask of the solid body.
    Steps: subsample the grid by `stride` and time by `tstride`; add Gaussian noise with
    std = noise * RMS of the velocity fluctuation; set `dropout` of the vectors to NaN at
    random; set `blobs` ellipses per snapshot to NaN (laser shadows / reflections, placed
    at random in each snapshot); set the body to NaN.
    Returns u2, v2 (with NaN), mask (ny2, nx2 body mask), truth = (u_true, v_true) on the same grid.
    """
    rng = np.random.default_rng(seed)
    ut = np.asarray(u[::stride, ::stride, ::tstride], dtype=np.float64)
    vt = np.asarray(v[::stride, ::stride, ::tstride], dtype=np.float64)
    ny, nx, m = ut.shape
    mask = np.zeros((ny, nx), bool) if body is None else np.asarray(body)[::stride, ::stride].copy()

    def rms_fluct(f):
        return np.sqrt(np.mean((f - f.mean(axis=2, keepdims=True)) ** 2))

    u2 = ut + noise * rms_fluct(ut) * rng.standard_normal(ut.shape)
    v2 = vt + noise * rms_fluct(vt) * rng.standard_normal(vt.shape)

    gap = rng.random((ny, nx, m)) < dropout
    jj, ii = np.mgrid[0:ny, 0:nx]
    for k in range(m):
        for _ in range(blobs):
            cx, cy = rng.uniform(0.1 * nx, 0.95 * nx), rng.uniform(0.1 * ny, 0.9 * ny)
            ax, ay = rng.uniform(0.03, 0.07) * nx, rng.uniform(0.05, 0.12) * ny
            gap[:, :, k] |= ((ii - cx) / ax) ** 2 + ((jj - cy) / ay) ** 2 <= 1.0
    gap |= mask[:, :, None]
    u2[gap] = np.nan
    v2[gap] = np.nan
    return u2.astype(np.float32), v2.astype(np.float32), mask, (ut.astype(np.float32), vt.astype(np.float32))


# ---------------------------------------------------------------- plots

def plot_field(ax, f, x=None, y=None, vlim=None, cmap=CMAP_DIV, body=None):
    """Signed field on (x, y) with a symmetric colour range."""
    f = np.asarray(f, dtype=float)
    if vlim is None:
        vlim = np.nanpercentile(np.abs(f), 99)
    ext = None if x is None else [x[0], x[-1], y[0], y[-1]]
    im = ax.imshow(f, origin="lower", extent=ext, cmap=cmap, vmin=-vlim, vmax=vlim,
                   interpolation="bilinear", aspect="equal")
    if body is not None and x is not None:
        ax.add_patch(plt.Circle(body[:2], body[2], color="0.35", zorder=3))
    ax.set_xticks([])
    ax.set_yticks([])
    return im


def plot_sigma(ax, S, label=None, color=ACC4, marker="o"):
    """Singular values on a log axis."""
    ax.semilogy(np.arange(1, len(S) + 1), S, marker=marker, ms=3, lw=1, color=color, label=label)
    ax.set_xlabel("mode index i")
    ax.set_ylabel(r"$\sigma_i$")
    ax.grid(True, which="both", alpha=0.3)


def plot_modes(fig, U, ny, nx, k=4, x=None, y=None, body=None, titles=None):
    """First k spatial modes as images in one row."""
    axs = fig.subplots(1, k)
    for i in range(k):
        plot_field(axs[i], to_field(U[:, i], ny, nx), x, y, body=body)
        axs[i].set_title(titles[i] if titles else f"mode {i + 1}")
    return axs


def plot_phase_plane(ax, a1, a2, color=ACC4):
    """a1-a2 phase plane. A mode pair gives a circle."""
    ax.plot(a1, a2, "-", color=color, lw=1.2)
    ax.plot(a1[0], a2[0], "o", color=WARM)
    ax.set_aspect("equal")
    ax.set_xlabel(r"$a_1(t)$")
    ax.set_ylabel(r"$a_2(t)$")
    ax.grid(True, alpha=0.3)


def plot_dmd_eigs(ax, lam, color=ACC4, label=None, marker="o", circle=True):
    """DMD eigenvalues on the complex plane with the unit circle."""
    if circle:
        t = np.linspace(0, 2 * np.pi, 400)
        ax.plot(np.cos(t), np.sin(t), "--", color="0.5", lw=1)
    ax.plot(np.real(lam), np.imag(lam), marker, color=color, ms=6, mfc="none", label=label)
    ax.set_aspect("equal")
    ax.set_xlabel(r"Re($\lambda$)")
    ax.set_ylabel(r"Im($\lambda$)")
    ax.grid(True, alpha=0.3)
