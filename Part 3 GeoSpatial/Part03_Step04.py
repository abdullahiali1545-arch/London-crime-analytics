from pathlib import Path
import os
import re
import numpy as np
import pandas as pd
import geopandas as gpd
import matplotlib as mpl
import matplotlib.pyplot as plt
import matplotlib.patheffects as pe
import imageio.v2 as imageio
from PIL import Image  # used to enforce same size as safety net

# ========================
# USER SETTINGS
# ========================
OFFENCE = "VIOLENCE AGAINST THE PERSON"
MEASURE = "Offences"
FPS = 1.75                   # playback speed; GIF uses ~1000/FPS ms per frame
ADD_LABELS = True
FIG_W, FIG_H = 8, 8           # figure size (inches)
DPI = 240                     # per-frame resolution

# ========================
# PATHS (relative to current working dir)
# ========================
WORK_DIR   = Path.cwd()
SHAPEFILE  = WORK_DIR / "London_Borough_Excluding_MHW.shp"
EXCEL_PATH = WORK_DIR / "M1045_MonthlyCrimeDashboard_TNOCrimeData.xlsx"
#SHEET_NAME = "borough_offenceGroup_agg"
FRAMES_DIR = WORK_DIR / "frames_monthly"
FRAMES_DIR.mkdir(exist_ok=True)

# ========================
# HELPERS
# ========================
def safe_name(s: str) -> str:
    s = str(s).strip().replace(' ', '_').replace('–', '-').replace('—', '-')
    return re.sub(r'[\\/:\*\?"<>\|]', '-', s)

# ========================
# LOAD DATA
# ========================
map_df = gpd.read_file(SHAPEFILE)

df = pd.read_excel(EXCEL_PATH)
df["Area name"] = df["Area name"].astype(str).str.strip()
map_df["NAME"]  = map_df["NAME"].astype(str).str.strip()

# Optional: remove non-borough rows
invalid = {"City of London", "Unknown", "Other / NK"}
df = df[~df["Area name"].isin(invalid)]

# Filter to offence + measure
df = df[(df["Offence Group"] == OFFENCE) & (df["Measure"] == MEASURE)].copy()
if df.empty:
    raise ValueError(f"No rows after filtering for Offence='{OFFENCE}' and Measure='{MEASURE}'.")

# Parse time to monthly datetime and normalize to first of month
df["Month_Year"] = pd.to_datetime(df["Month_Year"], errors="coerce")
if df["Month_Year"].isna().all():
    raise ValueError("Could not parse 'Month_Year' to datetime.")
df["Period"] = df["Month_Year"].dt.to_period("M").dt.to_timestamp()

# Numeric value column
value_col = "Count" if "Count" in df.columns else df.select_dtypes("number").columns[-1]

# Aggregate to one value per borough per month (safety)
monthly = (
    df.groupby(["Area name", "Period"], as_index=False)[value_col]
      .sum()
      .rename(columns={"Area name": "NAME"})
)

# Chronological list of months across all years
months = sorted(monthly["Period"].unique().tolist())
if len(months) <= 1:
    print(f"⚠️ Only {len(months)} month(s) found after filtering.")

# ------------------------
# Projection & layout locks
# ------------------------
# Project to British National Grid for better geometry/labels
map_proj = map_df.to_crs(epsg=27700)

# Precompute label points inside polygons
label_pts = map_proj.representative_point()
labels_df = map_proj[["NAME"]].copy()
labels_df["x"] = label_pts.x
labels_df["y"] = label_pts.y

# FIXED plot extents (prevents auto-rescale)
xmin, ymin, xmax, ymax = map_proj.total_bounds

# FIXED colour scale across all months
vmin = monthly[value_col].min()
vmax = monthly[value_col].max()
if pd.isna(vmin) or pd.isna(vmax):
    raise ValueError(f"'{value_col}' has no numeric values after filtering.")

norm = mpl.colors.Normalize(vmin=vmin, vmax=vmax)
cmap = mpl.cm.get_cmap("Reds")

print(f"Animating {len(months)} months: {months[0].strftime('%B %Y')} → {months[-1].strftime('%B %Y')}")
print(f"Value column: {value_col!r} | Fixed scale vmin={vmin}, vmax={vmax}")

# ========================
# BUILD FRAMES (UNIFORM SIZE)
# ========================
frame_files = []

for m in months:
    mdf = monthly[monthly["Period"] == m].copy()
    merged = map_proj.merge(mdf[["NAME", value_col]], on="NAME", how="left")

    # Create a figure with explicit axes rectangles so layout is identical every frame
    fig = plt.figure(figsize=(FIG_W, FIG_H))
    # main map axes (left, bottom, width, height in 0..1)
    ax = fig.add_axes([0.02, 0.02, 0.80, 0.88])
    # fixed colorbar axes
    cax = fig.add_axes([0.85, 0.20, 0.03, 0.60])

    ax.set_axis_off()
    ax.set_xlim(xmin, xmax)
    ax.set_ylim(ymin, ymax)
    ax.set_aspect("equal")

    merged.plot(
        column=value_col,
        cmap="Reds",
        linewidth=0.7,
        edgecolor="black",
        vmin=vmin,
        vmax=vmax,
        ax=ax,
        missing_kwds={
            "color": "gainsboro",
            "edgecolor": "black",
            "hatch": "///",
            "label": "No data"
        }
    )

    # ================================
    # LIGHTER BOROUGH LABELS (slightly darker grey)
    # ================================
    if ADD_LABELS:
        for _, r in labels_df.iterrows():
            ax.text(
                r["x"], r["y"], r["NAME"],
                fontsize=5.9,
                fontweight="regular",
                ha="center",
                va="center",
                color="#000000",  
                path_effects=[pe.withStroke(linewidth=0.3)])

    # Fixed colorbar placement
    sm = plt.cm.ScalarMappable(cmap="Reds", norm=norm)
    sm._A = []
    cbar = fig.colorbar(sm, cax=cax)
    cbar.set_label(f"{MEASURE} ({OFFENCE})")

    # Title at a fixed location (doesn't affect tight bbox)
    fig.suptitle(
        f"London Crime Choropleth\n{OFFENCE} – {MEASURE} ({pd.Timestamp(m).strftime('%B %Y')})",
        fontsize=14, fontweight="bold", y=0.97
    )
    fig.patch.set_facecolor("white")

    tag = pd.Timestamp(m).strftime("%Y_%m")
    frame_path = FRAMES_DIR / f"frame_{safe_name(OFFENCE)}_{safe_name(MEASURE)}_{tag}.png"

    # IMPORTANT: Do NOT use bbox_inches='tight' (it changes pixel size per frame)
    fig.savefig(frame_path, dpi=DPI, facecolor="white")
    plt.close(fig)

    frame_files.append(frame_path)

if not frame_files:
    raise RuntimeError("No frames created — check filters and monthly periods.")

# ========================
# ENSURE SAME PIXEL SIZE (safety net)
# ========================
pil_imgs = [Image.open(p).convert("RGB") for p in frame_files]
w0, h0 = pil_imgs[0].size
for i in range(1, len(pil_imgs)):
    if pil_imgs[i].size != (w0, h0):
        pil_imgs[i] = pil_imgs[i].resize((w0, h0), Image.BICUBIC)

# Convert to numpy arrays for imageio
frames_img = [np.array(im) for im in pil_imgs]

# ========================
# WRITE GIF (always works)
# ========================
gif_path = WORK_DIR / f"animation_{safe_name(OFFENCE)}_{safe_name(MEASURE)}_month.gif"
gif_duration_ms = int(1000 / FPS if FPS > 0 else 800)
imageio.mimsave(str(gif_path), frames_img, duration=gif_duration_ms, loop=0)
print("GIF created:", gif_path.resolve())

# ========================
# WRITE MP4 (needs ffmpeg; with fallback)
# ========================
mp4_path = WORK_DIR / f"animation_{safe_name(OFFENCE)}_{safe_name(MEASURE)}_month.mp4"

mp4_ok = False
mp4_error = None

try:
    imageio.mimsave(str(mp4_path), frames_img, fps=FPS, macro_block_size=None, quality=8)
    mp4_ok = True
except Exception as e:
    mp4_error = e
    print("Default MP4 writer failed. Trying with imageio-ffmpeg…")
    try:
        import imageio_ffmpeg
        ffmpeg_exe = imageio_ffmpeg.get_ffmpeg_exe()
        if ffmpeg_exe and os.path.isfile(ffmpeg_exe):
            os.environ["IMAGEIO_FFMPEG_EXE"] = ffmpeg_exe
            imageio.mimsave(str(mp4_path), frames_img, fps=FPS, macro_block_size=None, quality=8)
            mp4_ok = True
        else:
            print("imageio-ffmpeg installed, but ffmpeg binary not found.")
    except Exception as ee:
        mp4_error = ee

if mp4_ok:
    print("MP4 created:", mp4_path.resolve())
else:
    print("\n⚠️ MP4 not created.")
    print("Install FFmpeg and restart Spyder, then re-run:")
    print("  • conda install -c conda-forge ffmpeg   (recommended)")
    print("  • or: pip install imageio-ffmpeg        (bundled ffmpeg)")
    if mp4_error:
        print("Last MP4 error:", mp4_error)