"""Make the lecture animations in figures/anim/ (GIF, MP4 and a first-frame PNG each).

Every animation uses the functions of rom_utils.py and the data in data/, so the output is
reproducible. Rules for every animation:
  - 960 x 540 px, 3 to 8 s, 10 to 15 fps, loops, GIF <= 2 MB (fewer colours if larger)
  - colours from the slide theme (rom_utils.DEEP_BLUE ... LIGHT, CMAP_SEQ, CMAP_DIV)
  - the first frame already shows the conclusion, because PowerPoint shows the first
    frame in the edit view and in the PDF export

Usage:
  python scripts/make_animations.py                 # all, in priority order
  python scripts/make_animations.py --only 07 11    # names that start with 07 or 11
  python scripts/make_animations.py --check-only    # check the files that exist

In a notebook (plays the animation inline, as in Colab):
  import make_animations as ma
  ma.show("07_modes_and_coeffs")
"""
import argparse
import functools
import io
import json
import os
import sys
import time

import numpy as np
import matplotlib
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.animation import FuncAnimation, FFMpegWriter, PillowWriter
from PIL import Image

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from rom_utils import (dmd_reconstruct, exact_dmd, gappy_fill, load_cylinder, plot_field, pod,  # noqa: E402
                       rank_for_energy, to_field, to_X)
from rom_utils import (ACC1, ACC2, ACC3, ACC4, ACC5, CARD, DEEP_BLUE, LIGHT, WARM,  # noqa: E402,F401
                       CMAP_DIV, CMAP_SEQ)

OUT = os.path.join(ROOT, "figures", "anim")
W, H, DPI = 960, 540, 100
GIF_MAX_MB = 2.0
GIF_COLORS = (128, 64, 32, 16)      # tried in this order until the GIF is small enough
SECONDS = (3.0, 8.0)
FPS = (10, 15)


# ---------------------------------------------------------------- style

def _font_family():
    """Arial like the slides, Meiryo UI for Japanese; only fonts that are installed."""
    have = {f.name for f in font_manager.fontManager.ttflist}
    want = ["Arial", "Liberation Sans", "DejaVu Sans", "Meiryo UI", "Meiryo", "Yu Gothic",
            "Noto Sans CJK JP", "IPAexGothic"]
    return [f for f in want if f in have] or ["DejaVu Sans"]


def _has_japanese_font():
    """Colab has no Japanese font; text that needs one falls back to English there."""
    have = {f.name for f in font_manager.fontManager.ttflist}
    return any(f in have for f in ("Meiryo UI", "Meiryo", "Yu Gothic", "Noto Sans CJK JP", "IPAexGothic"))


# Applied only while an animation is built and drawn, so a notebook keeps its own style.
STYLE = {
    "font.family": _font_family(),
    "font.size": 13,
    "axes.titlesize": 14,
    "axes.labelsize": 13,
    "figure.facecolor": "white",
    "axes.facecolor": "white",
    "savefig.facecolor": "white",
    "text.color": DEEP_BLUE,
    "axes.labelcolor": DEEP_BLUE,
    "axes.edgecolor": DEEP_BLUE,
    "xtick.color": DEEP_BLUE,
    "ytick.color": DEEP_BLUE,
    "axes.prop_cycle": matplotlib.cycler(color=[ACC4, WARM, ACC1, ACC2, ACC3, ACC5]),
    "mathtext.fontset": "dejavusans",
    "axes.unicode_minus": False,
}


def new_fig():
    """A 960 x 540 px figure. Do not use bbox_inches='tight': the pixel size must stay fixed."""
    return plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI)


# ---------------------------------------------------------------- data (loaded once)

@functools.lru_cache(maxsize=None)
def cyl():
    """Data set 1 with its POD. X is the vorticity snapshot matrix (n x m)."""
    d = load_cylinder(os.path.join(ROOT, "data", "cylinder_vort_f32.npz"))
    vort = d["vort"].astype(np.float64)
    ny, nx, m = vort.shape
    X = to_X(vort)
    U, S, VT, xbar = pod(X, subtract_mean=True)
    return dict(d=d, vort=vort, u=d["u"].astype(np.float64), v=d["v"].astype(np.float64),
                x=d["x"], y=d["y"], dt=round(float(d["dt"]), 6), ny=ny, nx=nx, m=m,
                body=(*d["body"][:2], d["body"][2]), X=X, U=U, S=S, VT=VT, xbar=xbar)


@functools.lru_cache(maxsize=None)
def piv():
    """Data set 2 and its truth (data set 1 on the same grid and time step)."""
    d2 = np.load(os.path.join(ROOT, "data", "cylinder_pivlike_f32.npz"))
    meta = json.loads(str(d2["meta"]))
    c = cyl()
    s = int(meta["args"]["stride"]) // int(c["d"]["grid_stride"])
    ts = int(meta["args"]["tstride"])
    return dict(u=d2["u"].astype(np.float64), v=d2["v"].astype(np.float64), x=d2["x"], y=d2["y"],
                dt=round(float(d2["dt"]), 6), mask=d2["mask"], meta=meta,
                u_true=c["u"][::s, ::s, ::ts], v_true=c["v"][::s, ::s, ::ts])


# ---------------------------------------------------------------- registry

ANIMS = {}      # name -> (priority, function)
PRIORITY = {"must": 0, "recommended": 1, "optional": 2}


def register(name, priority):
    """Decorator. The function returns dict(fig=..., update=..., frames=int, fps=int).

    update(i) draws frame i (0 <= i < frames). update(0) must show the conclusion.
    """
    def deco(fn):
        ANIMS[name] = (PRIORITY[priority], fn)
        return fn
    return deco


# ---------------------------------------------------------------- writers

class QuantizedPillowWriter(PillowWriter):
    """PillowWriter that uses one shared palette for all frames and keeps the GIF small.

    A shared palette stops colour flicker between frames. The number of colours goes down
    (GIF_COLORS) until the file is at most GIF_MAX_MB.
    """

    def finish(self):
        frames = [f.convert("RGB") for f in self._frames]
        step = max(1, len(frames) // 12)
        sample = frames[::step]
        strip = Image.new("RGB", (frames[0].width, frames[0].height * len(sample)))
        for i, f in enumerate(sample):
            strip.paste(f, (0, i * f.height))
        # GIF stores frame times in 10 ms steps; round the cumulative time so the total
        # length matches the MP4 (for 12 fps: 80, 80, 90 ms, ...).
        edges = np.round(np.arange(len(frames) + 1) * 100 / self.fps) * 10
        duration = [int(v) for v in np.diff(edges)]
        for n_col in GIF_COLORS:
            pal = strip.quantize(colors=n_col, method=Image.Quantize.MEDIANCUT)
            q = [f.quantize(palette=pal, dither=Image.Dither.NONE) for f in frames]
            buf = io.BytesIO()
            q[0].save(buf, format="GIF", save_all=True, append_images=q[1:], duration=duration,
                      loop=0, optimize=True, disposal=1)
            if buf.tell() <= GIF_MAX_MB * 1e6:
                break
        self.colors_used = n_col
        with open(self.outfile, "wb") as fh:
            fh.write(buf.getvalue())


def _ffmpeg_path():
    try:
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except ImportError:
        return "ffmpeg"


plt.rcParams["animation.ffmpeg_path"] = _ffmpeg_path()


def render(name):
    """Write <name>.gif, <name>.mp4 and <name>.png (first frame) into figures/anim/."""
    os.makedirs(OUT, exist_ok=True)
    base = os.path.join(OUT, name)
    with plt.rc_context(STYLE):
        made = ANIMS[name][1]()
        fig, update, n, fps = made["fig"], made["update"], made["frames"], made["fps"]
        update(0)
        fig.savefig(base + ".png", dpi=DPI)

        anim = FuncAnimation(fig, update, frames=n, interval=1000 / fps, blit=False)
        gif = QuantizedPillowWriter(fps=fps)
        anim.save(base + ".gif", writer=gif, dpi=DPI)
        mp4 = FFMpegWriter(fps=fps, codec="libx264", bitrate=-1,
                           extra_args=["-pix_fmt", "yuv420p", "-crf", "20"])
        anim.save(base + ".mp4", writer=mp4, dpi=DPI)
    plt.close(fig)
    return gif.colors_used


def jshtml(name, dpi=72):
    """The animation as HTML with play controls (matplotlib to_jshtml), for Jupyter and Colab.

    dpi 72 gives 691 x 389 px frames; JPEG frames keep the HTML at about 40 % of PNG.
    """
    with plt.rc_context({**STYLE, "animation.frame_format": "jpeg"}):
        made = ANIMS[name][1]()
        fig = made["fig"]
        fig.set_dpi(dpi)
        anim = FuncAnimation(fig, made["update"], frames=made["frames"],
                             interval=1000 / made["fps"], blit=False)
        html = anim.to_jshtml(fps=made["fps"], default_mode="loop")
    plt.close(fig)
    return html


def show(name, dpi=72):
    """Play one animation inline in a notebook. The same GIF is in figures/anim/."""
    from IPython.display import HTML, display
    display(HTML(jshtml(name, dpi)))


# ---------------------------------------------------------------- check

def check(name):
    """Measure the three files of one animation against the rules."""
    base = os.path.join(OUT, name)
    r = {"name": name}
    with Image.open(base + ".gif") as im:
        r["size_px"] = list(im.size)
        durs = []
        for k in range(im.n_frames):
            im.seek(k)
            durs.append(im.info.get("duration", 0))
        # Pillow stores identical neighbouring frames once, with their durations added,
        # so the GIF can have fewer frames than the MP4 while its length is the same.
        r["gif_frames"] = len(durs)
        r["seconds"] = round(sum(durs) / 1000, 2)
        r["loop"] = im.info.get("loop", None) == 0
    r["gif_mb"] = round(os.path.getsize(base + ".gif") / 1e6, 2)
    r["mp4_mb"] = round(os.path.getsize(base + ".mp4") / 1e6, 2)
    try:
        import imageio_ffmpeg
        nf, secs = imageio_ffmpeg.count_frames_and_secs(base + ".mp4")
        r["mp4_frames"], r["mp4_seconds"] = nf, round(secs, 2)
        r["fps"] = round(nf / secs, 1)
    except Exception as e:  # noqa: BLE001
        r["mp4_frames"], r["mp4_seconds"], r["fps"] = None, str(e), 0
    with Image.open(base + ".png") as im:
        r["png_px"] = list(im.size)
    problems = []
    if r["size_px"] != [W, H] or r["png_px"] != [W, H]:
        problems.append("size")
    if not SECONDS[0] <= r["seconds"] <= SECONDS[1]:
        problems.append("length")
    if not FPS[0] <= r["fps"] <= FPS[1]:
        problems.append("fps")
    if r["gif_mb"] > GIF_MAX_MB:
        problems.append("gif>2MB")
    if not r["loop"]:
        problems.append("loop")
    if not isinstance(r["mp4_seconds"], float) or abs(r["mp4_seconds"] - r["seconds"]) > 0.15:
        problems.append("gif and mp4 length differ")
    r["ok"] = not problems
    r["problems"] = problems
    return r


# ---------------------------------------------------------------- animations
# One function per animation. main() renders them in priority order: must, recommended, optional.
def _loop_window(a, lo=55, hi=65):
    """Start index and length (in snapshots) of the window whose end best matches its start."""
    m = a.shape[1]
    best = (np.inf, 0, lo)
    for L in range(lo, hi + 1):
        for k0 in range(0, m - L):
            d = np.linalg.norm(a[:, k0 + L] - a[:, k0])
            if d < best[0]:
                best = (d, k0, L)
    return best[1], best[2]


@register("07_modes_and_coeffs", "must")
def anim_07_modes_and_coeffs():
    """Vorticity evolves; the spatial modes u1, u2 stay fixed; only a1(t), a2(t) move.

    The pair (a1, a2) traces a circle, so one mode pair describes one shedding cycle.
    """
    c = cyl()
    ny, nx, x, y, body, dt = c["ny"], c["nx"], c["x"], c["y"], c["body"], c["dt"]
    U, S, VT = c["U"], c["S"], c["VT"]
    a = S[:2, None] * VT[:2]
    k0, L = _loop_window(a)
    ks = np.arange(k0, k0 + L)
    kc = np.arange(k0, k0 + L + 1)
    t = ks * dt
    tc = kc * dt
    vlim = np.percentile(np.abs(c["vort"][:, :, ks]), 99)
    amax = 1.15 * np.abs(a[:, kc]).max()

    fig = new_fig()
    fig.text(0.02, 0.945, "Fixed spatial modes, moving time coefficients", fontsize=15,
             fontweight="bold", va="center")

    axf = fig.add_axes([0.02, 0.40, 0.56, 0.46])
    im = plot_field(axf, to_field(c["vort"][:, :, ks[0]], ny, nx),
                    x, y, vlim=vlim, cmap=CMAP_DIV)
    axf.add_patch(plt.Circle(body[:2], body[2], color=CARD, zorder=3))
    axf.set_title("vorticity of the flow", fontsize=13, loc="left")
    ttxt = axf.text(0.98, 0.06, "", transform=axf.transAxes, ha="right", fontsize=12,
                    bbox=dict(fc="white", ec="none", alpha=0.8, pad=1.5))

    for j, (ax_x, lab) in enumerate([(0.62, "mode 1 (fixed)"), (0.81, "mode 2 (fixed)")]):
        axm = fig.add_axes([ax_x, 0.56, 0.17, 0.26])
        um = to_field(U[:, j], ny, nx)
        plot_field(axm, um, x, y, vlim=np.percentile(np.abs(um), 99), cmap=CMAP_DIV)
        axm.add_patch(plt.Circle(body[:2], body[2], color=CARD, zorder=3))
        axm.set_title(lab, fontsize=12)

    fig.text(0.03, 0.29, r"$x(t)\ \approx\ \bar{x}$", fontsize=15, va="center")
    fig.text(0.07, 0.21, r"$+\ a_1(t)\,u_1\ +\ a_2(t)\,u_2$", fontsize=15, va="center")
    fig.text(0.03, 0.12, "u fixed, a(t) moves", fontsize=12, va="center")

    axp = fig.add_axes([0.36, 0.10, 0.20, 0.26])
    axp.set_aspect("equal")
    axp.set_xlim(-amax, amax)
    axp.set_ylim(-amax, amax)
    axp.grid(True, alpha=0.3)
    axp.set_xlabel(r"$a_1$", fontsize=12, labelpad=1)
    axp.set_ylabel(r"$a_2$", fontsize=12, labelpad=1)
    axp.tick_params(labelsize=10)
    axp.set_xticks([])
    axp.set_yticks([])
    axp.set_title("one mode pair = circle", fontsize=12)
    axp.plot(a[0, kc], a[1, kc], color=ACC5, alpha=0.3, lw=2)
    trail, = axp.plot([], [], color=ACC4, lw=2.2)
    dotp, = axp.plot([], [], "o", color=WARM, ms=8, mec=DEEP_BLUE)

    dots, trails = [], []
    for j, (yy, col, nm) in enumerate([(0.33, ACC4, r"$a_1(t)$"), (0.10, WARM, r"$a_2(t)$")]):
        axt = fig.add_axes([0.68, yy, 0.30, 0.17])
        axt.plot(tc, a[j, kc], color=col, alpha=0.3, lw=2)
        tr, = axt.plot([], [], color=col, lw=2.2)
        dt_, = axt.plot([], [], "o", color=col, ms=8, mec=DEEP_BLUE)
        axt.set_xlim(tc[0], tc[-1])
        axt.set_ylim(-amax, amax)
        axt.set_ylabel(nm, fontsize=12, labelpad=1)
        axt.set_yticks([0])
        axt.tick_params(labelsize=10)
        axt.grid(True, alpha=0.3)
        if j == 0:
            axt.set_xticklabels([])
        else:
            axt.set_xlabel("time t", fontsize=12, labelpad=1)
        trails.append(tr)
        dots.append(dt_)

    def update(i):
        im.set_data(to_field(c["vort"][:, :, ks[i]], ny, nx))
        ttxt.set_text(f"t = {t[i]:.1f}")
        trail.set_data(a[0, ks[:i + 1]], a[1, ks[:i + 1]])
        dotp.set_data([a[0, ks[i]]], [a[1, ks[i]]])
        for j in range(2):
            trails[j].set_data(t[:i + 1], a[j, ks[:i + 1]])
            dots[j].set_data([t[i]], [a[j, ks[i]]])

    return dict(fig=fig, update=update, frames=L, fps=10)


@register("08_U_from_snapshots", "recommended")
def anim_08_U_from_snapshots():
    """Mode 1 is built by adding the snapshots x~_k, each multiplied by its weight v_k1.

    The running sum looks more and more like u1 from the SVD; the correlation reaches 1.00.
    The loop starts on the finished sum, so the first frame shows the result.
    """
    c = cyl()
    ny, nx, m, x, y, body = c["ny"], c["nx"], c["m"], c["x"], c["y"], c["body"]
    U, S, VT, xbar = c["U"], c["S"], c["VT"], c["xbar"]
    Xt = c["X"] - xbar
    w = VT[0]
    u1 = U[:, 0]
    cum = np.cumsum(Xt * w[None, :], axis=1) / S[0]
    corr = (cum * u1[:, None]).sum(0) / (np.linalg.norm(cum, axis=0) * np.linalg.norm(u1) + 1e-30)
    vs = np.percentile(np.abs(Xt), 99)
    vu = np.percentile(np.abs(u1), 99)
    vc = np.percentile(np.abs(cum), 99)

    step = 3
    kk = np.minimum(np.arange(2, m + step, step), m - 1)
    kk = np.unique(kk)
    hold = 10
    seq = np.concatenate([np.full(hold, m - 1), kk])
    t = np.arange(m)

    fig = new_fig()
    fig.text(0.02, 0.93, "Mode 1 is a weighted sum of snapshots", fontsize=16,
             fontweight="bold", va="center")

    ax1 = fig.add_axes([0.02, 0.50, 0.30, 0.30])
    im1 = plot_field(ax1, to_field(Xt[:, 0], ny, nx), x, y, vlim=vs, cmap=CMAP_DIV)
    ax1.add_patch(plt.Circle(body[:2], body[2], color=CARD, zorder=3))
    tt1 = ax1.set_title("", fontsize=13)

    ax2 = fig.add_axes([0.40, 0.50, 0.25, 0.30])
    ax2.plot(t, w, color=ACC4, alpha=0.3, lw=2)
    tr2, = ax2.plot([], [], color=ACC4, lw=2.2)
    dot2, = ax2.plot([], [], "o", color=WARM, ms=8, mec=DEEP_BLUE)
    ax2.set_xlim(0, m - 1)
    ax2.set_ylim(-1.15 * np.abs(w).max(), 1.15 * np.abs(w).max())
    ax2.set_title(r"weight $v_{k1}$", fontsize=13)
    ax2.set_xlabel("snapshot k", fontsize=12, labelpad=1)
    ax2.set_yticks([0])
    ax2.tick_params(labelsize=11)
    ax2.grid(True, alpha=0.3)

    ax3 = fig.add_axes([0.69, 0.50, 0.30, 0.30])
    im3 = plot_field(ax3, to_field(cum[:, 0], ny, nx), x, y, vlim=vc, cmap=CMAP_DIV)
    ax3.add_patch(plt.Circle(body[:2], body[2], color=CARD, zorder=3))
    ax3.set_title("running sum", fontsize=13)

    for xx in (0.345, 0.665):
        fig.text(xx, 0.65, r"$\times$" if xx < 0.5 else r"$\rightarrow$", fontsize=18,
                 ha="center", va="center")

    fig.text(0.02, 0.30, r"$u_1\ =\ X\,v_1/\sigma_1\ =\ \frac{1}{\sigma_1}\sum_k v_{k1}\,\tilde{x}_k$",
             fontsize=17, va="center")
    fig.text(0.02, 0.15, r"$\tilde{x}_k$ = snapshot minus the mean", fontsize=13, va="center")

    ax4 = fig.add_axes([0.58, 0.08, 0.18, 0.24])
    plot_field(ax4, to_field(u1, nx=nx, ny=ny), x, y, vlim=vu, cmap=CMAP_DIV)
    ax4.add_patch(plt.Circle(body[:2], body[2], color=CARD, zorder=3))
    ax4.set_title("target: $u_1$ from SVD", fontsize=12)

    ctxt = fig.text(0.87, 0.20, "", fontsize=17, ha="center", va="center", fontweight="bold")

    def update(i):
        k = int(seq[i])
        im1.set_data(to_field(Xt[:, k], ny, nx))
        tt1.set_text(rf"snapshot $\tilde{{x}}_k$,  k = {k + 1}")
        tr2.set_data(t[:k + 1], w[:k + 1])
        dot2.set_data([k], [w[k]])
        im3.set_data(to_field(cum[:, k], ny, nx))
        ctxt.set_text(f"corr = {corr[k]:.2f}")

    return dict(fig=fig, update=update, frames=len(seq), fps=10)


@register("06_build_X", "optional")
def anim_06_build_X():
    """Each snapshot (a vorticity image) is flattened and becomes one new column of X.

    X is tall and thin: n grid points by m snapshots.
    """
    c = cyl()
    ny, nx, x, y, body, dt = c["ny"], c["nx"], c["x"], c["y"], c["body"], c["dt"]
    X = c["X"]
    n = X.shape[0]
    mc = 30
    Xc = X[:, :mc]
    vlim = np.percentile(np.abs(Xc), 95)
    norm = np.clip((Xc / vlim + 1) / 2, 0, 1)
    base = CMAP_DIV(norm)
    hold = 10
    seq = np.concatenate([np.arange(mc), np.full(hold, mc - 1)])

    fig = new_fig()
    axf = fig.add_axes([0.03, 0.34, 0.48, 0.44])
    imf = plot_field(axf, to_field(Xc[:, 0], ny, nx), x, y, vlim=vlim, cmap=CMAP_DIV)
    axf.add_patch(plt.Circle(body[:2], body[2], color=CARD, zorder=3))
    axf.set_title("one snapshot", fontsize=18)
    ttxt = fig.text(0.27, 0.27, "", fontsize=16, ha="center", va="center")
    fig.text(0.27, 0.14, "1 column = 1 snapshot", fontsize=18, ha="center", va="center",
             fontweight="bold")

    axx = fig.add_axes([0.74, 0.17, 0.22, 0.64])
    imx = axx.imshow(base, aspect="auto", interpolation="nearest", origin="upper")
    axx.set_xticks([])
    axx.set_yticks([])
    axx.set_title("matrix X", fontsize=18)
    hl = plt.Rectangle((-0.5, -0.5), 1, n, fill=False, ec=WARM, lw=1.5)
    axx.add_patch(hl)
    fig.text(0.705, 0.49, f"n = {n:,} grid points", fontsize=16, rotation=90,
             ha="center", va="center")
    fig.text(0.85, 0.095, f"first {mc} snapshots", fontsize=16, ha="center", va="center")
    fig.add_artist(matplotlib.patches.FancyArrowPatch(
        (0.74, 0.145), (0.96, 0.145), transform=fig.transFigure, arrowstyle="<->",
        mutation_scale=14, color=DEEP_BLUE, lw=1.5))
    fig.add_artist(matplotlib.patches.FancyArrowPatch(
        (0.53, 0.56), (0.675, 0.56), transform=fig.transFigure, arrowstyle="-|>",
        mutation_scale=22, color=WARM, lw=3))

    def update(i):
        k = int(seq[i])
        imf.set_data(to_field(Xc[:, k], ny, nx))
        ttxt.set_text(f"t = {k * dt:.1f}")
        arr = base.copy()
        arr[:, k + 1:, 3] = 0.4
        imx.set_data(arr)
        hl.set_x(k - 0.5)

    return dict(fig=fig, update=update, frames=len(seq), fps=10)


@register("10b_rotating_pattern", "must")
def anim_10b_rotating_pattern():
    """A rotating pattern is the sum of two fixed modes with weights cos(wt) and sin(wt).

    Left: cos(theta - wt) on an annulus. Middle: the fixed modes cos(theta), sin(theta).
    Right: a needle at angle wt; its two projections are the weights. One mode pair is enough.
    """
    n, nfr, phi0 = 140, 72, 0.9
    g = np.linspace(-1, 1, n)
    gx, gy = np.meshgrid(g, g)
    rr, th = np.hypot(gx, gy), np.arctan2(gy, gx)
    ring = (rr > 0.3) & (rr < 1.0)
    cmap = CMAP_DIV.copy()
    cmap.set_bad("white")

    def img(f):
        return np.ma.masked_where(~ring, f)

    ph = phi0 + 2 * np.pi * np.arange(nfr) / nfr

    fig = new_fig()
    fig.text(0.5, 0.93, r"$\cos(\theta-\omega t)=\cos\theta\,\cos\omega t+\sin\theta\,\sin\omega t$",
             ha="center", va="center", fontsize=20)
    axl = fig.add_axes([0.02, 0.12, 0.38, 0.68])
    axm1 = fig.add_axes([0.45, 0.50, 0.14, 0.25])
    axm2 = fig.add_axes([0.45, 0.10, 0.14, 0.25])
    axd = fig.add_axes([0.68, 0.17, 0.30, 0.53])

    kw = dict(origin="lower", extent=[-1, 1, -1, 1], cmap=cmap, vmin=-1, vmax=1,
              interpolation="bilinear")
    iml = axl.imshow(img(np.cos(th - ph[0])), **kw)
    axm1.imshow(img(np.cos(th)), **kw)
    axm2.imshow(img(np.sin(th)), **kw)
    for ax in (axl, axm1, axm2):
        ax.set_xticks([])
        ax.set_yticks([])
        for s in ax.spines.values():
            s.set_visible(False)
    axl.set_title("rotating pattern", fontsize=18)
    axm1.set_title(r"$\cos\theta$", fontsize=18, color=ACC4)
    axm2.set_title(r"$\sin\theta$", fontsize=18, color=WARM)

    t = np.linspace(0, 2 * np.pi, 300)
    axd.plot(np.cos(t), np.sin(t), color=ACC5, alpha=0.3, lw=2)
    axd.axhline(0, color=ACC5, alpha=0.4, lw=1)
    axd.axvline(0, color=ACC5, alpha=0.4, lw=1)
    needle, = axd.plot([0, np.cos(ph[0])], [0, np.sin(ph[0])], color=DEEP_BLUE, lw=3)
    tip, = axd.plot([np.cos(ph[0])], [np.sin(ph[0])], "o", color=DEEP_BLUE, ms=9)
    vl, = axd.plot([np.cos(ph[0])] * 2, [0, np.sin(ph[0])], "--", color=ACC4, lw=1.5)
    hl, = axd.plot([0, np.cos(ph[0])], [np.sin(ph[0])] * 2, "--", color=WARM, lw=1.5)
    dc, = axd.plot([np.cos(ph[0])], [0], "o", color=ACC4, ms=11)
    ds, = axd.plot([0], [np.sin(ph[0])], "o", color=WARM, ms=11)
    axd.set_xlim(-1.3, 1.3)
    axd.set_ylim(-1.3, 1.3)
    axd.set_aspect("equal")
    axd.axis("off")
    axd.set_title("weights", fontsize=18)
    tc = fig.text(0.68, 0.10, "", fontsize=18, color=ACC4, ha="left", va="center")
    ts = fig.text(0.68, 0.04, "", fontsize=18, color=WARM, ha="left", va="center")

    def update(i):
        p = ph[i]
        c, s = np.cos(p), np.sin(p)
        iml.set_data(img(np.cos(th - p)))
        needle.set_data([0, c], [0, s])
        tip.set_data([c], [s])
        vl.set_data([c, c], [0, s])
        hl.set_data([0, c], [s, s])
        dc.set_data([c], [0])
        ds.set_data([0], [s])
        tc.set_text(rf"$\cos\omega t={c:+.2f}$")
        ts.set_text(rf"$\sin\omega t={s:+.2f}$")

    return dict(fig=fig, update=update, frames=nfr, fps=12)


@register("10a_rank_r_recon", "recommended")
def anim_10a_rank_r_recon():
    """Rank-r reconstruction of one snapshot and its error against r.

    The truncated SVD error equals sqrt(sum_{i>r} sigma_i^2 / sum sigma_i^2), so it is the best
    rank-r approximation (Eckart-Young). The error falls quickly, so a few modes are enough.
    """
    c = cyl()
    U, S, VT, xbar = c["U"], c["S"], c["VT"], c["xbar"]
    ny, nx, x, y, body = c["ny"], c["nx"], c["x"], c["y"], c["body"]
    k0, rmax, hold, tail = 60, 20, 3, 9
    Xt = c["X"] - xbar
    nrm = np.linalg.norm(Xt)
    rs = np.arange(1, rmax + 1)
    err = np.array([np.linalg.norm(Xt - (U[:, :r] * S[:r]) @ VT[:r]) / nrm for r in rs])
    s2 = S ** 2
    ey = np.array([np.sqrt(s2[r:].sum() / s2.sum()) for r in rs])
    truth = c["X"][:, k0]
    recs = [xbar[:, 0] + U[:, :r] @ (S[:r] * VT[:r, k0]) for r in rs]
    vlim = np.percentile(np.abs(truth), 99.5)
    seq = [r - 1 for r in rs for _ in range(hold)] + [rmax - 1] * tail

    fig = new_fig()
    axt = fig.add_axes([0.02, 0.55, 0.50, 0.33])
    axr = fig.add_axes([0.02, 0.09, 0.50, 0.33])
    axe = fig.add_axes([0.63, 0.17, 0.35, 0.70])
    ext = [x[0], x[-1], y[0], y[-1]]
    kw = dict(origin="lower", extent=ext, cmap=CMAP_DIV, vmin=-vlim, vmax=vlim,
              interpolation="bilinear", aspect="equal")
    axt.imshow(to_field(truth, ny, nx), **kw)
    imr = axr.imshow(to_field(recs[0], ny, nx), **kw)
    for ax in (axt, axr):
        ax.add_patch(plt.Circle(body[:2], body[2], color=CARD, zorder=3))
        ax.set_xticks([])
        ax.set_yticks([])
    axt.set_title("truth", fontsize=18)
    ttl = axr.set_title("", fontsize=18)

    axe.semilogy(rs, err, "-o", color=ACC4, alpha=0.25, lw=2, ms=5)
    ln, = axe.semilogy(rs[:1], err[:1], "-o", color=ACC4, lw=2.5, ms=6,
                       label=r"$\|\tilde X-\tilde X_r\|_F/\|\tilde X\|_F$")
    ly, = axe.semilogy(rs[:1], ey[:1], "o", mfc="none", mec=WARM, mew=2, ms=11,
                       label=r"$\sqrt{\sum_{i>r}\sigma_i^2/\sum\sigma_i^2}$")
    cur, = axe.semilogy([1], [err[0]], "o", color=DEEP_BLUE, ms=12, zorder=5)
    axe.set_xlim(0.5, rmax + 0.5)
    axe.set_ylim(err.min() * 0.6, 1.4)
    axe.set_xticks([1, 5, 10, 15, 20])
    axe.set_xlabel("rank r", fontsize=16)
    axe.set_ylabel("relative error", fontsize=16)
    axe.tick_params(labelsize=16)
    axe.grid(True, which="both", alpha=0.3)
    axe.legend(loc="upper right", fontsize=16, frameon=False)
    axe.set_title("best rank-r error", fontsize=18)

    def update(f):
        j = seq[f]
        r = j + 1
        imr.set_data(to_field(recs[j], ny, nx))
        ttl.set_text(f"rank r = {r}  (error of X: {err[j]:.0%})")
        ln.set_data(rs[:r], err[:r])
        ly.set_data(rs[:r], ey[:r])
        cur.set_data([r], [err[j]])

    return dict(fig=fig, update=update, frames=len(seq), fps=12)


@register("10c_translating_bump", "optional")
def anim_10c_translating_bump():
    """A Gaussian bump moves across the domain; its singular values keep growing in number.

    A stationary bump has one non-zero singular value. A translating one needs many modes,
    so a linear basis is not efficient for non-periodic translation. The loop starts on the
    final state, so the first frame shows the result.
    """
    N, M, hold = 200, 40, 12
    xg = np.linspace(0, 10, N)
    w = 0.5
    cs = np.linspace(1.5, 8.5, M)
    A = np.exp(-(xg[:, None] - cs[None, :]) ** 2 / (2 * w ** 2))
    sig = []
    for j in range(1, M + 1):
        s = np.linalg.svd(A[:, :j], compute_uv=False)
        sig.append(np.maximum(s / s[0], 1e-16))
    thr = 1e-2
    nsig = [int((s > thr).sum()) for s in sig]
    imax = 30
    idx = np.arange(1, imax + 1)
    fin = sig[-1][:imax]
    seq = [M - 1] * hold + list(range(M))
    nf = len(seq)

    fig = new_fig()
    axt = fig.add_axes([0.11, 0.60, 0.86, 0.27])
    axb = fig.add_axes([0.11, 0.15, 0.86, 0.32])
    ghosts = [axt.plot(xg, A[:, j], color=ACC4, lw=1.5, alpha=0.08)[0] for j in range(M)]
    bump, = axt.plot(xg, A[:, 0], color=ACC4, lw=3)
    axt.set_xlim(0, 10)
    axt.set_ylim(0, 1.15)
    axt.set_yticks([])
    axt.tick_params(labelsize=16)
    axt.set_xlabel("x", fontsize=16, labelpad=0)
    axt.set_title("moving bump, snapshots so far", fontsize=18)

    axb.semilogy(idx, fin, "-o", color=ACC4, alpha=0.3, lw=2, ms=5)
    axb.semilogy([1], [1], "o", color=WARM, ms=13, label="stationary bump")
    ln, = axb.semilogy([1], sig[0][:1], "-o", color=ACC4, lw=2.5, ms=6, label="moving bump")
    axb.set_xlim(0, imax + 0.5)
    axb.set_ylim(1e-8, 3)
    axb.set_yticks([1e-6, 1e-3, 1])
    axb.set_xlabel("index i", fontsize=16)
    axb.set_ylabel(r"$\sigma_i/\sigma_1$", fontsize=16)
    axb.tick_params(labelsize=16)
    axb.grid(True, which="both", alpha=0.3)
    axb.legend(loc="lower left", fontsize=16, frameon=False)
    txt = axb.text(0.98, 0.95, "", transform=axb.transAxes, ha="right", va="top", fontsize=16)

    def update(f):
        j = seq[f]
        for k, g in enumerate(ghosts):
            g.set_alpha(0.3 if k < j else 0.08)
        bump.set_data(xg, A[:, j])
        s = sig[j][:imax]
        ln.set_data(idx[:len(s)], s)
        txt.set_text(rf"moving: {nsig[j]} above 1%" "\n" "stationary: 1")

    return dict(fig=fig, update=update, frames=nf, fps=12)


def _unit_circle(ax, lw=1):
    t = np.linspace(0, 2 * np.pi, 400)
    ax.plot(np.cos(t), np.sin(t), "--", color=ACC5, lw=lw)


@register("11_dmd_rollout", "must")
def anim_11_dmd_rollout():
    """DMD fitted on the first 100 snapshots and rolled forward as a linear model.

    Left: true vorticity. Middle: x_k = sum phi_i b_i lam_i^k. Right: the 21 eigenvalues.
    The bar below shows the fit window, the forecast and the relative error at each k.
    """
    c = cyl()
    X, dt, ny, nx, m = c["X"], c["dt"], c["ny"], c["nx"], c["m"]
    n_fit, r = 100, 21
    lam, Phi, b, _, _ = exact_dmd(X[:, :n_fit], dt, r)
    Xh = dmd_reconstruct(Phi, lam, b, np.arange(m))
    err = 100 * np.linalg.norm(Xh - X, axis=0) / np.linalg.norm(X, axis=0)
    mean = c["xbar"][:, 0]
    vlim = np.percentile(np.abs(X - mean[:, None]), 99.5)

    ks = np.arange(0, m, 2)
    hold = 12
    frames = len(ks) + hold
    fps = 13

    fig = new_fig()
    fig.text(0.5, 0.94, r"DMD is a linear model:  $x_k\approx\sum_i \phi_i\, b_i\, \lambda_i^{\,k}$",
             ha="center", va="center", fontsize=18)
    ax_t = fig.add_axes([0.02, 0.45, 0.35, 0.36])
    ax_d = fig.add_axes([0.38, 0.45, 0.35, 0.36])
    ax_e = fig.add_axes([0.79, 0.45, 0.20, 0.20 * 960 / 540])
    ax_b = fig.add_axes([0.08, 0.12, 0.88, 0.17])

    ims = []
    for ax, title in ((ax_t, "True vorticity"), (ax_d, "DMD rollout")):
        im = plot_field(ax, to_field(X[:, 0] - mean, ny, nx), c["x"], c["y"], vlim=vlim, body=None)
        ax.add_patch(plt.Circle(c["body"][:2], c["body"][2], color=CARD, zorder=3))
        ax.set_title(title, fontsize=14, pad=4)
        ims.append(im)
    t_k = fig.text(0.195, 0.455, "", ha="center", va="center", fontsize=13)
    t_d = fig.text(0.555, 0.455, "", ha="center", va="center", fontsize=13)

    _unit_circle(ax_e)
    ax_e.plot(lam.real, lam.imag, "o", color=ACC4, ms=6, mfc="none", mew=1.5)
    ax_e.set_xlim(-1.3, 1.3)
    ax_e.set_ylim(-1.3, 1.3)
    ax_e.set_aspect("equal")
    ax_e.set_xticks([-1, 0, 1])
    ax_e.set_yticks([-1, 0, 1])
    ax_e.tick_params(labelsize=12)
    ax_e.set_title(r"eigenvalues $\lambda_i$", fontsize=14, pad=4)
    ax_e.grid(True, alpha=0.3)

    ymax = 1.3 * err.max()
    ax_b.axvspan(0, n_fit - 1, color=ACC4, alpha=0.15, lw=0)
    ax_b.axvspan(n_fit - 1, m - 1, color=WARM, alpha=0.25, lw=0)
    ax_b.plot(np.arange(m), err, color=ACC5, lw=1.5, alpha=0.3)
    (ln,) = ax_b.plot([], [], color=ACC5, lw=2.5)
    (pt,) = ax_b.plot([], [], "o", color=ACC5, ms=8)
    vl = ax_b.axvline(0, color=DEEP_BLUE, lw=1.5)
    ax_b.set_xlim(0, m - 1)
    ax_b.set_ylim(0, ymax)
    ax_b.text(n_fit / 2, ymax * 1.03, f"fit window (k < {n_fit})", ha="center", va="bottom", fontsize=13)
    ax_b.text((n_fit + m) / 2, ymax * 1.03, "forecast", ha="center", va="bottom", fontsize=13)
    ax_b.set_xlabel("snapshot index k", fontsize=13)
    ax_b.set_ylabel("error [%]", fontsize=13)
    ax_b.tick_params(labelsize=12)
    t_err = ax_b.text(0.5, 0.2, "", transform=ax_b.transAxes, ha="center", va="center", fontsize=14,
                      bbox=dict(fc="white", ec="none", alpha=0.85, pad=2))

    def update(i):
        k = int(ks[min(i, len(ks) - 1)])
        ims[0].set_data(to_field(X[:, k] - mean, ny, nx))
        ims[1].set_data(to_field(Xh[:, k] - mean, ny, nx))
        t_k.set_text(f"k = {k}")
        t_d.set_text("fit" if k < n_fit else "forecast")
        ln.set_data(np.arange(k + 1), err[:k + 1])
        pt.set_data([k], [err[k]])
        vl.set_xdata([k, k])
        t_err.set_text("relative error at k: " + f"{err[k]:.2f} %")

    return dict(fig=fig, update=update, frames=frames, fps=fps)


@register("12_noise_eigs", "recommended")
def anim_12_noise_eigs():
    """DMD eigenvalues of the cylinder data while white noise grows to twice the fluctuation RMS.

    One noise realisation is scaled by the level. r keeps 99 % of the energy, so every
    eigenvalue belongs to the flow. The eigenvalues move inside the unit circle (damped side),
    the higher harmonics most. On the full circle the shift is small, so the right panel
    enlarges |lambda| against the frequency.
    """
    c = cyl()
    X, dt = c["X"], c["dt"]
    r = rank_for_energy(c["S"], 0.99)
    sx = np.sqrt(np.mean((X - c["xbar"]) ** 2))
    E = np.random.default_rng(0).standard_normal(X.shape)
    n_lv, top = 40, 2.0
    levels = np.linspace(0, top, n_lv)
    lams = np.array([exact_dmd(X + lv * sx * E, dt, r)[0] for lv in levels])
    rad, st = np.abs(lams), np.angle(lams) / (2 * np.pi * dt)
    hold = 12
    frames = n_lv + hold
    fps = 12

    fig = new_fig()
    fig.text(0.5, 0.94, "Noise pushes DMD eigenvalues inside the unit circle", ha="center",
             va="center", fontsize=18)
    ax_f = fig.add_axes([0.04, 0.17, 0.30, 0.30 * 960 / 540])
    ax_z = fig.add_axes([0.45, 0.17, 0.52, 0.66])

    _unit_circle(ax_f)
    ax_f.set_xlim(-1.25, 1.25)
    ax_f.set_ylim(-1.25, 1.25)
    ax_f.set_aspect("equal")
    ax_f.set_xticks([-1, 0, 1])
    ax_f.set_yticks([-1, 0, 1])
    ax_f.set_title(r"eigenvalues $\lambda_i$", fontsize=14, pad=4)

    lo = rad.min() - 0.1 * (1 - rad.min())
    smax = 1.08 * np.abs(st).max()
    ax_z.axhline(1.0, ls="--", color=ACC5, lw=1.2)
    ax_z.text(0.98 * smax, 1.0, r"$|\lambda| = 1$", ha="right", va="bottom", fontsize=13, color=ACC5)
    ax_z.axhspan(lo, 1.0, color=WARM, alpha=0.06, lw=0)
    ax_z.text(0.5, 0.05, "inside the circle = damped", transform=ax_z.transAxes, ha="center", fontsize=13)
    ax_z.set_xlim(-smax, smax)
    ax_z.set_ylim(lo, 1.0 + 0.15 * (1 - lo))
    ax_z.set_xlabel(r"frequency $\arg(\lambda)/(2\pi\Delta t)$ = St", fontsize=13)
    ax_z.set_ylabel(r"$|\lambda|$", fontsize=13)
    ax_z.set_title(r"zoom on $|\lambda|$", fontsize=14, pad=4)
    ax_z.ticklabel_format(axis="y", useOffset=False, style="plain")

    sc = []
    for ax, xs, ys in ((ax_f, lams.real, lams.imag), (ax_z, st, rad)):
        ax.tick_params(labelsize=12)
        ax.grid(True, alpha=0.3)
        ax.plot(xs[-1], ys[-1], "o", color=WARM, ms=8, alpha=0.45, mew=0, zorder=2)
        trail = ax.scatter([], [], s=12, color=ACC4, alpha=0.25, zorder=3, lw=0)
        cur = ax.scatter(xs[0], ys[0], s=55, facecolors="none", edgecolors=ACC4, linewidths=1.8, zorder=4)
        sc.append((xs, ys, trail, cur))

    fig.text(0.04, 0.06, f"r = {r} (99 % energy)", fontsize=14)
    t_lv = fig.text(0.97, 0.03, "", ha="right", va="bottom", fontsize=14)

    def update(i):
        j = min(i, n_lv - 1)
        for xs, ys, trail, cur in sc:
            trail.set_offsets(np.c_[xs[:j].ravel(), ys[:j].ravel()] if j else np.empty((0, 2)))
            cur.set_offsets(np.c_[xs[j], ys[j]])
        t_lv.set_text(f"noise = {100 * levels[j]:.0f} % of fluctuation RMS   (orange: {100 * top:.0f} %)")

    return dict(fig=fig, update=update, frames=frames, fps=fps)


@register("14b_gappy_fill", "must")
def anim_14b_gappy_fill():
    """Gappy POD fills the PIV-like gaps iteration by iteration.

    Left: u of one snapshot, gap outlines in CARD. Top right: sigma_i / sigma_1 of the filled
    matrix approaches the truth. Bottom right: the error at the gap entries falls.
    """
    d = piv()
    mask = d["mask"]
    live = ~mask.ravel()

    def stack(fu, fv):
        return np.vstack([fu.reshape(-1, fu.shape[2])[live], fv.reshape(-1, fv.shape[2])[live]])

    Xnan = stack(d["u"], d["v"])
    Xtrue = stack(d["u_true"], d["v_true"])
    G = np.isnan(Xnan)
    mbar = np.nanmean(Xnan, axis=1, keepdims=True)
    states = [np.where(G, mbar, Xnan)]
    nlive = int(live.sum())
    ny2, nx2 = mask.shape
    ref = np.linalg.norm((Xtrue - Xtrue.mean(1, keepdims=True))[G])

    def err(Xf):
        return np.linalg.norm((Xf - Xtrue)[G]) / ref

    def sig(Xf):
        s = np.linalg.svd(Xf - Xf.mean(1, keepdims=True), compute_uv=False)[:-1]
        return s / s[0]

    gappy_fill(Xnan, 10, n_iter=10, callback=lambda it, Xf, S: states.append(Xf.copy()))
    errs = np.array([err(Xf) for Xf in states])
    sigs = [sig(Xf) for Xf in states]
    s_true = sig(Xtrue)
    ns = len(states)
    nk = 30

    k = int(np.argmax(G[:nlive].sum(axis=0)))
    ufull = np.full((ns, ny2 * nx2), np.nan)
    for j, Xf in enumerate(states):
        ufull[j, live] = Xf[:nlive, k] - mbar[:nlive, 0]
    ufull = ufull.reshape(ns, ny2, nx2)
    vlim = np.nanpercentile(np.abs(ufull[-1]), 99)
    gapimg = np.zeros(ny2 * nx2)
    gapimg[live] = G[:nlive, k]
    gapimg = gapimg.reshape(ny2, nx2)
    x, y = d["x"], d["y"]

    fig = new_fig()
    axf = fig.add_axes([0.03, 0.40, 0.50, 0.40])
    axs = fig.add_axes([0.505 + 0.115, 0.58, 0.365, 0.28])
    axe = fig.add_axes([0.505 + 0.115, 0.14, 0.365, 0.28])
    fig.text(0.03, 0.91, "u fluctuation, one snapshot", fontsize=18, ha="left")
    fig.text(0.62, 0.91, "Gappy POD, r = 10", fontsize=18, ha="left")

    ext = [x[0], x[-1], y[0], y[-1]]
    cmap = CMAP_DIV.copy()
    cmap.set_bad(CARD)
    im = axf.imshow(ufull[0], origin="lower", extent=ext, cmap=cmap, vmin=-vlim, vmax=vlim,
                    interpolation="nearest", aspect="equal")
    axf.contour(x, y, gapimg, levels=[0.5], colors=[CARD], linewidths=1.6)
    axf.set_xticks([])
    axf.set_yticks([])
    for sp in axf.spines.values():
        sp.set_visible(False)
    t_it = fig.text(0.03, 0.30, "", fontsize=20, ha="left")
    t_er = fig.text(0.03, 0.20, "", fontsize=20, ha="left")
    fig.text(0.03, 0.10, "dark outline = gaps", fontsize=16, ha="left", color=CARD)

    idx = np.arange(1, nk + 1)
    axs.semilogy(idx, s_true[:nk], "--", color=DEEP_BLUE, lw=2, label="truth")
    lines = [axs.semilogy(idx, s[:nk], "-", lw=2.2, color=ACC4)[0] for s in sigs]
    axs.set_xlabel("mode index i", fontsize=16)
    axs.set_ylabel(r"$\sigma_i/\sigma_1$", fontsize=16)
    axs.tick_params(labelsize=16)
    axs.grid(True, which="both", alpha=0.3)
    axs.legend(fontsize=16, loc="upper right", frameon=False, handlelength=1.6)

    axe.plot(np.arange(ns), errs, "-", color=ACC4, lw=2, alpha=0.3)
    (cur,) = axe.plot([], [], "-o", color=WARM, lw=2.5, ms=7)
    axe.set_xlim(-0.4, ns - 0.6)
    axe.set_ylim(0, errs.max() * 1.1)
    axe.set_xlabel("iteration", fontsize=16)
    axe.set_ylabel("error at gaps", fontsize=16)
    axe.set_xticks([0, 5, 10])
    axe.tick_params(labelsize=16)
    axe.grid(True, alpha=0.3)

    hold = [5] * (ns - 1) + [10]
    seq = [j for j, h in enumerate(hold) for _ in range(h)]

    def update(i):
        j = seq[i]
        im.set_data(ufull[j])
        for q, ln in enumerate(lines):
            if q == j:
                ln.set_color(WARM)
                ln.set_alpha(1.0)
                ln.set_zorder(5)
            elif q == ns - 1:
                ln.set_color(ACC4)
                ln.set_alpha(0.35)
                ln.set_zorder(3)
            elif q < j:
                ln.set_color(ACC4)
                ln.set_alpha(0.18)
                ln.set_zorder(2)
            else:
                ln.set_alpha(0.0)
        cur.set_data(np.arange(j + 1), errs[:j + 1])
        t_it.set_text("mean fill" if j == 0 else f"iteration {j}")
        t_er.set_text(f"relative error at gaps {errs[j]:.2f}")

    return dict(fig=fig, update=update, frames=len(seq), fps=10)


@register("14a_pvc_orbit", "recommended")
def anim_14a_pvc_orbit():
    """A precessing vortex core moves on a small circle; the wall pressure follows its angle.

    Left: Rankine vortex (vorticity and velocity arrows) in a round pipe, its centre on a
    circle of radius r* R; an image vortex outside the wall keeps the flow tangent to it.
    Right: pressure at one wall sensor, cos(theta_s - 2 pi f t).
    r* is read from the thesis figure; the core radius a and the sensor angle are display choices.
    """
    R, rs, a = 1.0, 0.13, 0.26
    th_s = np.deg2rad(40.0)
    nper, per = 2, 30
    N = nper * per
    ph = 2 * np.pi * np.arange(N) / per

    n = 150
    g = np.linspace(-1.1, 1.1, n)
    GX, GY = np.meshgrid(g, g)
    inside = GX ** 2 + GY ** 2 <= R ** 2

    qs = np.linspace(-0.85, 0.85, 12)
    QX, QY = np.meshgrid(qs, qs)
    qin = QX ** 2 + QY ** 2 <= 0.9 ** 2
    qx, qy = QX[qin], QY[qin]

    def qvel(cx, cy):
        dx, dy = qx - cx, qy - cy
        rho2 = np.maximum(dx ** 2 + dy ** 2, 1e-9)
        k = np.where(rho2 < a ** 2, 1.0 / a ** 2, 1.0 / rho2)
        k2 = 1.0 / np.maximum((qx - cx * R ** 2 / rs ** 2) ** 2 + (qy - cy * R ** 2 / rs ** 2) ** 2, 1e-9)
        u = -k * dy + k2 * (qy - cy * R ** 2 / rs ** 2)
        v = k * dx - k2 * (qx - cx * R ** 2 / rs ** 2)
        sp = np.hypot(u, v)
        s = np.minimum(sp, 3.0) / np.maximum(sp, 1e-9)
        return u * s, v * s

    vort = np.empty((N, n, n))
    UV = []
    for i, p in enumerate(ph):
        cx, cy = rs * np.cos(p), rs * np.sin(p)
        d2 = (GX - cx) ** 2 + (GY - cy) ** 2
        w = 0.5 * (1 - np.tanh((np.sqrt(d2) - a) / 0.025))
        vort[i] = np.where(inside, w, np.nan)
        UV.append(qvel(cx, cy))

    fig = new_fig()
    ax = fig.add_axes([0.01, 0.12, 0.48, 0.74])
    ax.set_xlim(-1.45, 1.12)
    ax.set_ylim(-1.08, 1.08)
    ax.set_aspect("equal")
    ax.axis("off")
    fig.text(0.28, 0.91, "Vortex core in a pipe", fontsize=18, ha="center")
    im = ax.imshow(vort[0], origin="lower", extent=[-1.1, 1.1, -1.1, 1.1], cmap=CMAP_DIV,
                   vmin=-1, vmax=1, interpolation="bilinear", zorder=1)
    ax.add_patch(plt.Circle((0, 0), R, fill=False, ec=DEEP_BLUE, lw=3.5, zorder=6))
    t = np.linspace(0, 2 * np.pi, 200)
    ax.plot(rs * np.cos(t), rs * np.sin(t), "--", color=DEEP_BLUE, lw=1.6, alpha=0.3, zorder=4)
    (tail,) = ax.plot([], [], "-", color=DEEP_BLUE, lw=2.5, zorder=5)
    q = ax.quiver(qx, qy, *UV[0], color=ACC5, alpha=0.65, scale=22, width=0.004, zorder=3)
    (dot,) = ax.plot([], [], "o", color=DEEP_BLUE, ms=8, zorder=7)
    ax.plot(R * np.cos(th_s), R * np.sin(th_s), "s", color=ACC1, ms=13, zorder=8, mec=DEEP_BLUE)
    ax.text(1.07 * np.cos(th_s), 1.07 * np.sin(th_s), "sensor", fontsize=16, color=ACC1,
            ha="left", va="bottom", zorder=8)
    ang = np.deg2rad(215)
    ax.annotate(r"$r^*$ = 0.13", xy=(rs * np.cos(ang), rs * np.sin(ang)), xytext=(-1.42, -0.85),
                fontsize=18, ha="left", va="center", zorder=9,
                arrowprops=dict(arrowstyle="-", color=DEEP_BLUE, lw=1.4))

    axp = fig.add_axes([0.62, 0.27, 0.36, 0.46])
    tc = np.linspace(0, nper, 400)
    pc = np.cos(th_s - 2 * np.pi * tc)
    axp.plot(tc, pc, "-", color=ACC4, lw=2.2, alpha=0.3)
    (pm,) = axp.plot([], [], "o", color=ACC1, ms=11, mec=DEEP_BLUE, zorder=5, clip_on=False)
    axp.axhline(0, color=DEEP_BLUE, lw=0.8, alpha=0.3)
    axp.set_xlim(0, nper)
    axp.set_ylim(-1.25, 1.25)
    axp.set_xticks(range(nper + 1))
    axp.set_yticks([])
    axp.set_xlabel("t f  (cycles)", fontsize=16)
    axp.tick_params(labelsize=16)
    axp.grid(True, axis="x", alpha=0.3)
    for s in ("top", "right"):
        axp.spines[s].set_visible(False)
    fig.text(0.80, 0.91, "Wall pressure", fontsize=18, ha="center")
    fig.text(0.80, 0.83, r"$p'(\theta_s,t) \propto \cos(\theta_s - 2\pi f t)$", fontsize=16,
             ha="center")
    note = "合成データ（博論 Fig. 3.12 の r* を模擬）" if _has_japanese_font()         else "synthetic data (r* as in thesis Fig. 3.12)"
    fig.text(0.985, 0.025, note, fontsize=16, ha="right", va="bottom")

    tail_n = 10

    def update(i):
        p = ph[i]
        im.set_data(vort[i])
        q.set_UVC(*UV[i])
        dot.set_data([rs * np.cos(p)], [rs * np.sin(p)])
        pp = p - 2 * np.pi * np.arange(tail_n + 1) / per
        tail.set_data(rs * np.cos(pp), rs * np.sin(pp))
        tt = i / per
        pm.set_data([tt], [np.cos(th_s - 2 * np.pi * tt)])

    return dict(fig=fig, update=update, frames=N, fps=12)


def _deim_points(Xi):
    """DEIM greedy selection. Returns the points and the residual map of every step."""
    n, p = Xi.shape
    pts = [int(np.argmax(np.abs(Xi[:, 0])))]
    res = [Xi[:, 0].copy()]
    for j in range(1, p):
        c = np.linalg.solve(Xi[pts, :j], Xi[pts, j])
        r = Xi[:, j] - Xi[:, :j] @ c
        res.append(r)
        pts.append(int(np.argmax(np.abs(r))))
    return pts, res


@register("17_deim_greedy", "optional")
def anim_17_deim_greedy():
    """DEIM picks its sensor points one by one, always where the residual is largest.

    Left: the residual map and the points chosen so far, the newest one ringed. Right: the
    error of the nonlinear snapshots N interpolated from j points with the first j modes.
    """
    c = cyl()
    ny, nx = c["ny"], c["nx"]
    x, y = c["x"], c["y"]
    xc, yc, rad = c["body"]
    w, p = c["vort"], 20
    N = c["u"] * np.gradient(w, x, axis=1) + c["v"] * np.gradient(w, y, axis=0)
    xx, yy = np.meshgrid(x, y)
    inside = np.hypot(xx - xc, yy - yc) < rad * 1.1
    N[inside, :] = 0.0
    Xi = pod(to_X(N), subtract_mean=False)[0][:, :p]
    keep = ~inside.ravel()
    Xi_m = Xi * keep[:, None]
    pts, res = _deim_points(Xi_m)
    Nx = to_X(N)
    dei = np.array([np.linalg.norm(Nx - Xi_m[:, :j] @ np.linalg.solve(Xi_m[pts[:j], :j], Nx[pts[:j]]))
                    for j in range(1, p + 1)]) / np.linalg.norm(Nx)
    maps = [np.abs(r) * keep for r in res]
    peak = np.array([mp.max() for mp in maps])
    px = xx.ravel()[pts]
    py = yy.ravel()[pts]
    ext = [x[0], x[-1], y[0], y[-1]]

    fig = new_fig()
    ax = fig.add_axes([0.02, 0.17, 0.60, 0.70])
    ax2 = fig.add_axes([0.76, 0.25, 0.22, 0.58])
    # light low end: the slide is printed, so a mostly zero map must stay white
    light_seq = matplotlib.colors.LinearSegmentedColormap.from_list("light_seq", ["#ffffff", ACC3, ACC5])
    im = ax.imshow(np.zeros((ny, nx)), origin="lower", extent=ext, cmap=light_seq, vmin=0, vmax=1,
                   interpolation="bilinear", aspect="equal")
    ax.add_patch(plt.Circle((xc, yc), rad, color=CARD, zorder=3))
    ax.set_xticks([])
    ax.set_yticks([])
    ax.set_title("residual map |r_j| (scaled, square-root colour)", fontsize=14)
    ax.plot(px, py, "o", ms=7, color=WARM, alpha=0.3, mec="none", zorder=4)
    pts_art, = ax.plot([], [], "o", ms=8, color=WARM, mec=DEEP_BLUE, mew=0.8, zorder=5)
    ring, = ax.plot([], [], "o", ms=18, mfc="none", mec=WARM, mew=2.5, zorder=6)

    js = np.arange(1, p + 1)
    ax2.semilogy(js, dei, "-", color=ACC4, lw=2, alpha=0.3)
    curve, = ax2.semilogy([], [], "o-", color=ACC4, lw=2, ms=5)
    ax2.set_xlim(0.5, p + 0.5)
    ax2.set_xticks([1, 5, 10, 15, 20])
    ax2.tick_params(labelsize=12)
    ax2.set_xlabel("points chosen j", fontsize=12)
    ax2.set_ylabel("relative error of N", fontsize=13)
    ax2.set_title("DEIM error with j points", fontsize=14)
    ax2.grid(True, which="both", alpha=0.3)
    fig.text(0.02, 0.06, "next point = argmax |residual|", fontsize=16, color=DEEP_BLUE)
    step = fig.text(0.62, 0.06, "", fontsize=14, color=DEEP_BLUE)

    sub, hold = 3, 12
    frames = p * sub + hold

    def update(i):
        k = min(i // sub, p - 1)
        s = i % sub if i < p * sub else sub - 1
        im.set_data(np.sqrt(to_field(maps[k] / peak[k], ny, nx)))
        shown = k + 1 if s >= 1 else k
        pts_art.set_data(px[:shown], py[:shown])
        if s >= 1:
            ring.set_data([px[k]], [py[k]])
        else:
            ring.set_data([], [])
        nc = k + 1 if s >= 2 else k
        curve.set_data(js[:nc], dei[:nc])
        step.set_text(f"j = {k + 1} of {p}")

    return dict(fig=fig, update=update, frames=frames, fps=12)


@register("19_manifold_vs_subspace", "optional")
def anim_19_manifold_vs_subspace():
    """A curved solution manifold and the best straight line through the origin.

    The view turns once around the scene. The orange segments are the projection errors:
    a linear subspace of dimension 1 cannot follow a curved 1D manifold.
    """
    from matplotlib.lines import Line2D

    a = np.linspace(-2.3, 2.3, 200)
    P = np.stack([np.cos(a), np.sin(a), 0.6 * np.sin(2 * a)], axis=1)
    _, S, VT = np.linalg.svd(P, full_matrices=False)
    d1, d2 = VT[0], VT[1]
    proj = (P @ d1)[:, None] * d1[None, :]
    err = np.linalg.norm(P - proj) / np.linalg.norm(P)
    sel = np.linspace(0, len(a) - 1, 26).astype(int)
    tl = np.linspace(-1.5, 1.5, 2)
    line = tl[:, None] * d1[None, :]
    g = np.linspace(-1.3, 1.3, 8)
    G1, G2 = np.meshgrid(g, g)
    plane = G1[..., None] * d1 + G2[..., None] * d2

    fig = new_fig()
    ax = fig.add_axes([0.0, 0.0, 1.0, 0.92], projection="3d")
    ax.set_axis_off()
    ax.plot_surface(plane[..., 0], plane[..., 1], plane[..., 2], color=ACC3, alpha=0.12,
                    linewidth=0, shade=False)
    ax.plot(line[:, 0], line[:, 1], line[:, 2], "-", color=ACC5, lw=3.5)
    ax.plot(P[:, 0], P[:, 1], P[:, 2], "-", color=ACC4, lw=3.5)
    for k in sel:
        ax.plot([P[k, 0], proj[k, 0]], [P[k, 1], proj[k, 1]], [P[k, 2], proj[k, 2]],
                "-", color=WARM, lw=1.8)
    ax.plot(P[sel, 0], P[sel, 1], P[sel, 2], "o", ms=4, color=ACC4)
    ax.plot([0], [0], [0], "o", ms=6, color=DEEP_BLUE)
    ax.set_xlim(-1.2, 1.2)
    ax.set_ylim(-1.2, 1.2)
    ax.set_zlim(-1.0, 1.0)
    # Fill the frame; legend and text stay inside x = 150-810 px, the crop the slide uses.
    ax.set_box_aspect((1, 1, 0.8), zoom=1.55)
    handles = [Line2D([], [], color=ACC4, lw=3.5, label="nonlinear manifold"),
               Line2D([], [], color=ACC5, lw=3.5, label="linear subspace (best line)"),
               Line2D([], [], color=WARM, lw=2, label="projection error")]
    fig.legend(handles=handles, loc="upper left", bbox_to_anchor=(0.17, 0.98), fontsize=13,
               frameon=False)
    fig.text(0.83, 0.04, f"relative error of the line: {100 * err:.0f} %", fontsize=14,
             ha="right", color=DEEP_BLUE)

    frames, fps = 84, 12

    def update(i):
        ax.view_init(elev=22, azim=-60 + 360.0 * i / frames)

    return dict(fig=fig, update=update, frames=frames, fps=fps)


# ---------------------------------------------------------------- draft tube surrogate (toy model)
# Synthetic data only. A simple physics model plays the role of the experiment and the 3D CFD.
# Inputs: q = Q/Q_BEP and sigma (Thoma number); n_ED and the Froude number are fixed.
# Rotating part: free-vortex wall pressure -K/d^2 (thesis eq. 3-28 to 3-39, no image vortex).
# Synchronous part: elbow source (non-zero without cavitation) x 1D system response H(f; sigma).
SG_Q = (0.62, 0.85)            # vortex rope regime in Q/Q_BEP (Favrel+ 2017; Doerfler 2017)
SG_ZETA = 0.12                 # damping ratio of the 1D system response
SG_NR, SG_NT = 40, 120         # polar grid: radial cells, angular cells
SG_SENS = np.deg2rad([0.0, 90.0, 180.0, 270.0])


def _sg_f0(s):
    """Natural frequency f0/n of the 1D system. Compliance C ~ exp(-b sigma), f0 ~ 1/sqrt(L C)."""
    return 0.12 * np.exp(9.15 * (s - 0.08))


def _sg_H(f, s):
    """Response of the 1D system (one mode, idealised: no lock-in, no hysteresis)."""
    x = f / _sg_f0(s)
    return 1.0 / (1.0 - x ** 2 + 2j * SG_ZETA * x)


def _sg_truth_params(q, s):
    """Skeleton parameters of the toy truth: orbit radius r*, K = Gamma*^2, core radius a, f/n, source S0."""
    rs = 0.06 + 0.22 * (np.clip(0.86 - q, 0, None) / 0.24) ** 0.7
    K = (0.6 + 1.2 * (1.0 / q - 1.0 / 0.86)) ** 2
    a = 0.10 + 0.12 * np.exp(-(s - 0.08) / 0.04)
    f = 0.25 + 0.35 * 0.02 * np.log1p(np.exp((0.78 - q) / 0.02)) + 0.03 * (0.14 - s)
    return rs, K, a, f, 0.6 * rs * K


def _sg_grid():
    rc = (np.arange(SG_NR) + 0.5) / SG_NR
    th = 2 * np.pi * np.arange(SG_NT) / SG_NT
    return np.append(rc, 1.0), th           # the last ring is the wall (r = 1)


def _sg_field(rs, K, a):
    """Rankine-vortex pressure in the frame that turns with the vortex (vortex at angle 0).

    Rows: radial cells, then the wall ring. The wall is outside the core, so the wall value is -K/d^2.
    """
    r, th = _sg_grid()
    d2 = r[:, None] ** 2 + rs ** 2 - 2 * r[:, None] * rs * np.cos(th[None, :])
    return np.where(d2 >= a ** 2, -K / np.maximum(d2, 1e-12), -K / a ** 2 * (2 - d2 / a ** 2))


def _sg_rbf_fit(P, Y):
    """Thin-plate spline with a linear tail. P (m, 2) scaled inputs, Y (m, k) outputs."""
    m = len(P)
    r = np.linalg.norm(P[:, None] - P[None], axis=2)
    Phi = np.where(r > 0, r ** 2 * np.log(np.maximum(r, 1e-300)), 0.0)
    Pp = np.hstack([np.ones((m, 1)), P])
    A = np.block([[Phi, Pp], [Pp.T, np.zeros((3, 3))]])
    return np.linalg.solve(A, np.vstack([Y, np.zeros((3, Y.shape[1]))])), P


def _sg_rbf_eval(model, p):
    c, P = model
    r = np.linalg.norm(P - p, axis=1)
    phi = np.where(r > 0, r ** 2 * np.log(np.maximum(r, 1e-300)), 0.0)
    return np.concatenate([phi, [1.0, *p]]) @ c


@functools.lru_cache(maxsize=None)
def sg_surrogate():
    """Train the surrogate on a 4 x 4 grid of operating points.

    Field: phase-locked field (frame turning with the vortex) -> POD -> RBF on the POD coefficients.
    Scalars: f/n and the source S0 -> RBF. The synchronous part uses the known 1D model H(f; sigma).
    """
    qs = np.linspace(0.63, 0.84, 4)
    ss = np.linspace(0.09, 0.20, 4)
    pts = np.array([(q, s) for s in ss for q in qs])
    lo, hi = np.array([SG_Q[0], 0.08]), np.array([SG_Q[1], 0.20])
    scale = lambda p: (np.asarray(p) - lo) / (hi - lo)          # noqa: E731
    F = []
    sc = []
    for q, s in pts:
        rs, K, a, f, S0 = _sg_truth_params(q, s)
        F.append(_sg_field(rs, K, a).ravel())
        sc.append([f, S0])
    F = np.array(F).T
    mean = F.mean(axis=1, keepdims=True)
    U, S, _ = np.linalg.svd(F - mean, full_matrices=False)
    r = int(np.searchsorted(np.cumsum(S ** 2) / np.sum(S ** 2), 0.9999) + 1)
    coef = U[:, :r].T @ (F - mean)
    model_a = _sg_rbf_fit(scale(pts), coef.T)
    model_s = _sg_rbf_fit(scale(pts), np.array(sc))

    def predict(q, s):
        p = scale((q, s))
        field = (mean[:, 0] + U[:, :r] @ _sg_rbf_eval(model_a, p)).reshape(SG_NR + 1, SG_NT)
        f, S0 = _sg_rbf_eval(model_s, p)
        return field, f, S0

    return dict(pts=pts, predict=predict, r=r)


def _sg_wave(ring, f, sync, tn):
    """Wall pressure at the sensors over the time axis tn = t n. ring: wall values in the turning frame."""
    th = 2 * np.pi * np.arange(SG_NT) / SG_NT
    ph = 2 * np.pi * f * tn
    out = []
    for ts in SG_SENS:
        rot = np.interp(np.mod(ts - ph, 2 * np.pi), th, ring, period=2 * np.pi) - ring.mean()
        out.append(rot + abs(sync) * np.cos(ph + np.angle(sync)))
    return np.array(out)


@register("25_surrogate_map", "recommended")
def anim_25_surrogate_map():
    """A moving operating point on the (Q/Q_BEP, sigma) map; the surrogate predicts the field and the
    wall-pressure waveforms at once. Truth = the toy physics model at the same point.
    """
    ja = _has_japanese_font()
    T = (lambda j, e: j) if ja else (lambda j, e: e)
    sg = sg_surrogate()
    predict = sg["predict"]
    t0 = time.perf_counter()
    for _ in range(200):
        predict(0.73, 0.15)
    ms = (time.perf_counter() - t0) / 200 * 1e3

    N, fps = 90, 12
    ang = 2 * np.pi * np.arange(N) / N + np.deg2rad(128)    # frame 0 sits in the resonance band
    path = np.c_[0.735 + 0.09 * np.cos(ang), 0.145 + 0.05 * np.sin(ang)]
    tn = np.linspace(0, 8, 240)
    frames = []
    for (q, s) in path:
        rs, K, a, f, S0 = _sg_truth_params(q, s)
        Ft = _sg_field(rs, K, a)
        Fr, fr, S0r = predict(q, s)
        wt = _sg_wave(Ft[-1], f, S0 * _sg_H(f, s), tn)
        wr = _sg_wave(Fr[-1], fr, S0r * _sg_H(fr, s), tn)
        pt = Ft[:-1] - Ft[:-1].mean(axis=1, keepdims=True)
        pr = Fr[:-1] - Fr[:-1].mean(axis=1, keepdims=True)
        frames.append(dict(q=q, s=s, f=f, fr=fr, pt=pt, pr=pr, wt=wt, wr=wr,
                           ef=np.linalg.norm(pr - pt) / np.linalg.norm(pt),
                           ew=np.linalg.norm(wr - wt) / np.linalg.norm(wt),
                           res=abs(f / _sg_f0(s) - 1) < 0.1))
    ef = np.array([d["ef"] for d in frames])
    ew = np.array([d["ew"] for d in frames])
    print(f"25_surrogate_map: POD modes {sg['r']}, field error mean {ef.mean():.3f} max {ef.max():.3f}, "
          f"waveform error mean {ew.mean():.3f} max {ew.max():.3f}, predict {ms:.2f} ms")

    fig = new_fig()
    fig.text(0.02, 0.965, T("運転点を動かすと、流れ場と壁面圧力をその場で予測", "Move the operating point: field "
                            "and wall pressure predicted at once"), fontsize=17, va="top")
    # operating map
    axm = fig.add_axes([0.075, 0.37, 0.235, 0.45])
    qg = np.linspace(0.56, 0.90, 200)
    sgr = np.linspace(0.08, 0.21, 200)
    QG, SGm = np.meshgrid(qg, sgr)
    band = np.abs(_sg_truth_params(QG, SGm)[3] / _sg_f0(SGm) - 1) < 0.1
    axm.contourf(QG, SGm, band.astype(float), levels=[0.5, 1.5], colors=[WARM], alpha=0.45)
    for x0, x1 in ((0.56, SG_Q[0]), (SG_Q[1], 0.90)):
        axm.axvspan(x0, x1, color=DEEP_BLUE, alpha=0.12, lw=0)
    axm.text(0.59, 0.205, T("範囲外", "out"), fontsize=12, ha="center", va="top", rotation=90)
    axm.text(0.875, 0.205, T("範囲外", "out"), fontsize=12, ha="center", va="top", rotation=90)
    axm.plot(sg["pts"][:, 0], sg["pts"][:, 1], "o", color=DEEP_BLUE, ms=6, zorder=4)
    axm.plot(path[:, 0], path[:, 1], "-", color=DEEP_BLUE, lw=1, alpha=0.25)
    (trail,) = axm.plot([], [], "-", color=ACC1, lw=2.5, zorder=5)
    (star,) = axm.plot([], [], "*", color=ACC1, ms=19, mec=DEEP_BLUE, zorder=6)
    axm.set_xlim(0.56, 0.90)
    axm.set_ylim(0.08, 0.21)
    axm.set_xticks([0.6, 0.7, 0.8, 0.9])
    axm.set_yticks([0.10, 0.15, 0.20])
    axm.tick_params(labelsize=12)
    axm.set_xlabel(r"$Q/Q_{BEP}$", fontsize=13)
    axm.set_ylabel(T("σ（トーマ数）", "σ (Thoma number)"), fontsize=13)
    from matplotlib.lines import Line2D
    from matplotlib.patches import Patch
    axm.legend(handles=[Line2D([], [], ls="", marker="o", color=DEEP_BLUE, ms=6, label=T("学習 16 点", "training 16")),
                        Line2D([], [], ls="", marker="*", color=ACC1, mec=DEEP_BLUE, ms=13, label=T("予測点", "query")),
                        Patch(color=WARM, alpha=0.45, label=T("共振帯", "resonance"))],
               loc="lower left", bbox_to_anchor=(-0.28, 1.0), ncol=3, fontsize=12, frameon=False,
               handletextpad=0.1, columnspacing=0.5, handlelength=1.0)
    info = fig.text(0.02, 0.235, "", fontsize=12, va="top", linespacing=1.5)
    # fields
    r, th = _sg_grid()
    re = np.linspace(0, 1, SG_NR + 1)
    te = 2 * np.pi * (np.arange(SG_NT + 1) - 0.5) / SG_NT
    XE, YE = re[:, None] * np.cos(te[None]), re[:, None] * np.sin(te[None])
    vmax = 1.0
    meshes, etxt = [], []
    for k, (y0, ttl) in enumerate(((0.47, T("真値（実験・CFD の代役）", "truth (stands in for test/CFD)")),
                                    (0.05, T("サロゲート（POD＋RBF）", "surrogate (POD + RBF)")))):
        ax = fig.add_axes([0.355, y0, 0.205, 0.365])
        ax.set_aspect("equal")
        ax.axis("off")
        ax.set_xlim(-1.08, 1.08)
        ax.set_ylim(-1.08, 1.08)
        mesh = ax.pcolormesh(XE, YE, np.zeros((SG_NR, SG_NT)), cmap=CMAP_DIV, vmin=-vmax, vmax=vmax,
                             shading="flat", rasterized=True)
        ax.add_patch(plt.Circle((0, 0), 1, fill=False, ec=DEEP_BLUE, lw=2.5))
        for j, ts in enumerate(SG_SENS):
            ax.plot(np.cos(ts), np.sin(ts), "s", ms=8, color=[ACC4, ACC1, ACC5, WARM][j], mec=DEEP_BLUE,
                    zorder=5, clip_on=False)
        fig.text(0.4575, y0 + 0.375, ttl, fontsize=12, ha="center")
        meshes.append(mesh)
        if k == 1:
            etxt.append(fig.text(0.4575, 0.012, "", fontsize=12, ha="center"))
    # sensor waveforms
    ymax = max(np.abs(d["wt"]).max() for d in frames) * 1.1
    lines = []
    names = ["S1", "S2", "S3", "S4"]
    cols = [ACC4, ACC1, ACC5, WARM]
    fig.text(0.79, 0.875, T("壁面センサの圧力脈動 Cp'", "wall pressure at the sensors Cp'"), fontsize=13,
             ha="center")
    for j in range(4):
        ax = fig.add_axes([0.62, 0.705 - j * 0.165, 0.36, 0.14])
        ax.axhline(0, color=DEEP_BLUE, lw=0.6, alpha=0.3)
        (lt,) = ax.plot(tn, frames[0]["wt"][j], "-", color=cols[j], lw=4, alpha=0.35)
        (lr,) = ax.plot(tn, frames[0]["wr"][j], "--", color=DEEP_BLUE, lw=1.4)
        lines.append((lt, lr))
        ax.set_xlim(0, 8)
        ax.set_ylim(-ymax, ymax)
        ax.set_yticks([-1, 1])
        ax.tick_params(axis="y", labelsize=11, length=2)
        ax.text(-0.55, 0, names[j], fontsize=12, ha="right", va="center", color=cols[j], fontweight="bold")
        for sp in ("top", "right") + (("bottom",) if j < 3 else ()):
            ax.spines[sp].set_visible(False)
        if j < 3:
            ax.set_xticks([])
        else:
            ax.tick_params(labelsize=12)
            ax.set_xlabel(T("t·n（ランナの回転数）", "t·n (runner turns)"), fontsize=12)
    fig.text(0.60, 0.06, T("太線：真値　破線：サロゲート", "thick: truth   dashed: surrogate"), fontsize=12)
    wtxt = fig.text(0.98, 0.06, "", fontsize=12, ha="right")
    fig.text(0.98, 0.012, T(f"予測1点 {ms:.2f} ms（この PC で測定）", f"{ms:.2f} ms per query (this PC)"),
             fontsize=12, ha="right")
    fig.text(0.02, 0.012, T("合成データの簡易モデル（n_ED・Fr 固定）", "synthetic toy model (n_ED, Fr fixed)"),
             fontsize=12)

    def update(i):
        d = frames[i]
        trail.set_data(path[max(0, i - 12):i + 1, 0], path[max(0, i - 12):i + 1, 1])
        star.set_data([d["q"]], [d["s"]])
        shift = (i * SG_NT // 12) % SG_NT
        meshes[0].set_array(np.roll(d["pt"], shift, axis=1).ravel())
        meshes[1].set_array(np.roll(d["pr"], shift, axis=1).ravel())
        etxt[0].set_text(T(f"場の誤差 {d['ef'] * 100:.1f} %", f"field error {d['ef'] * 100:.1f} %"))
        wtxt.set_text(T(f"波形の誤差 {d['ew'] * 100:.1f} %", f"waveform error {d['ew'] * 100:.1f} %"))
        for j, (lt, lr) in enumerate(lines):
            lt.set_ydata(d["wt"][j])
            lr.set_ydata(d["wr"][j])
        warn = T("共振帯：線形 1D モデルは目安", "resonance band: 1D model is a guide") if d["res"] else ""
        info.set_text(T(f"Q/Q_BEP = {d['q']:.2f}　σ = {d['s']:.3f}\nf/n = {d['fr']:.3f}（真値 {d['f']:.3f}）\n{warn}",
                        f"Q/Q_BEP = {d['q']:.2f}  σ = {d['s']:.3f}\nf/n = {d['fr']:.3f} (truth {d['f']:.3f})\n{warn}"))
        info.set_color("#8A5A00" if d["res"] else DEEP_BLUE)

    return dict(fig=fig, update=update, frames=N, fps=fps)


# ---------------------------------------------------------------- main

def main():
    matplotlib.use("Agg")
    p = argparse.ArgumentParser()
    p.add_argument("--only", nargs="*", default=None, help="name prefixes, for example 07 10b")
    p.add_argument("--check-only", action="store_true")
    a = p.parse_args()
    names = sorted(ANIMS, key=lambda k: (ANIMS[k][0], list(ANIMS).index(k)))
    if a.only:
        names = [n for n in names if any(n.startswith(o) for o in a.only)]
    results = []
    for n in names:
        t = time.time()
        if not a.check_only:
            colors = render(n)
        r = check(n)
        if not a.check_only:
            r["colors"] = colors
            r["render_s"] = round(time.time() - t, 1)
        results.append(r)
        print(json.dumps(r, ensure_ascii=False))
    bad = [r["name"] for r in results if not r["ok"]]
    print(f"{len(results) - len(bad)}/{len(results)} OK" + (f"; problems: {bad}" if bad else ""))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
