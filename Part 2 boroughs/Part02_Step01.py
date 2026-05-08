
import os
import platform
from pathlib import Path
from datetime import datetime
import pandas as pd

# ========= SETTINGS =========
INPUT_XLSX = "M1045_MonthlyCrimeDashboard_TNOCrimeData.xlsx"  # your input file
SHEET_INDEX = 0  # change if data isn't on the first sheet
AREA_TYPE_VALUE = "Borough"

# You can ignore these now (no dialog / Downloads used in Option B)
USE_SAVE_DIALOG = False
OUTPUT_XLSX_NAME = "M1045_borough_grouped_two_rows_per_group.xlsx"

# Which sheets to include in the single workbook:
# - "agg"   : (required) subgroup rolled up to offence groups with 2 rows (Offences, Positive Outcomes)
# - "clean" : (optional) cleaned borough-level data before the roll-up (subgroups intact)
# - "pivot" : (optional) two helpful pivots (Offences-only, Positive Outcomes-only)
SHEETS_TO_EXPORT = ("agg", "clean", "pivot")

# (Optionally) add timestamp to filename to avoid overwrites
APPEND_TIMESTAMP = False  # set True if you want '..._YYYYMMDD_HHMMSS.xlsx'
# ============================


def reveal_in_file_browser(path: Path):
    """Open the output folder in the OS file explorer."""
    try:
        if platform.system() == "Windows":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":  # macOS
            os.system(f'open "{path}"')
        else:  # Linux
            os.system(f'xdg-open "{path}"')
    except Exception:
        pass


def build_output_name(base_name: str, add_timestamp: bool) -> str:
    """
    If add_timestamp is True and base_name is like 'file.xlsx',
    returns 'file_YYYYMMDD_HHMMSS.xlsx'.
    """
    if not add_timestamp:
        return base_name
    stem, ext = os.path.splitext(base_name)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{stem}_{stamp}{ext}"


# ---------- 1) Load ----------
df = pd.read_excel(INPUT_XLSX, sheet_name=SHEET_INDEX, engine="openpyxl")
df.columns = [str(c).strip() for c in df.columns]  # tidy headers

# Ensure Count is numeric (coerce non-numeric to NaN -> 0)
if "Count" in df.columns:
    df["Count"] = pd.to_numeric(df["Count"], errors="coerce").fillna(0)

# ---------- 2) Keep only Borough rows ----------
if "Area Type" not in df.columns:
    raise KeyError("Missing 'Area Type' column. Check your sheet headers.")

df_borough = df[df["Area Type"].astype(str).str.strip().eq(AREA_TYPE_VALUE)].copy()

# ---------- 3) Drop unneeded columns ----------
cols_to_drop = [c for c in ["Borough_SNT", "Area code", "FY_FYIndex", "Refresh Date"]
                if c in df_borough.columns]
df_borough.drop(columns=cols_to_drop, inplace=True)

# ---------- 4) Consolidate duplicates at raw borough level (pre-rollup) ----------
# (This collapses any accidental duplicates before we roll up subgroups)
raw_keys = [c for c in [
    "Month_Year", "Area Type", "Borough", "Area name",
    "Offence Group", "Offence Subgroup", "Measure", "Financial Year"
] if c in df_borough.columns]

if "Count" in df_borough.columns and raw_keys:
    df_borough = (
        df_borough
        .groupby(raw_keys, dropna=False, as_index=False)["Count"]
        .sum()
    )
else:
    df_borough = df_borough.drop_duplicates()

# ---------- 5) Roll up subgroups → offence groups, keeping measures separate ----------
group_keys = [c for c in [
    "Month_Year", "Borough", "Area name", "Area Type",
    "Offence Group", "Measure", "Financial Year"
] if c in df_borough.columns]

if "Count" not in df_borough.columns:
    raise KeyError("Missing 'Count' column—cannot aggregate without it.")

df_agg = (
    df_borough
    .groupby(group_keys, dropna=False, as_index=False)["Count"]
    .sum()
)

# ---------- 6) Guarantee two rows per (Month_Year × Borough × Offence Group) ----------
# We always want BOTH measures: Offences and Positive Outcomes.
desired_measures = ["Offences", "Positive Outcomes"]

# Base keys WITHOUT Measure
base_keys = [c for c in [
    "Month_Year", "Borough", "Area name", "Area Type",
    "Offence Group", "Financial Year"
] if c in df_agg.columns]

# Build the full grid: (all base combos) × (Offences, Positive Outcomes)
base = df_agg[base_keys].drop_duplicates().copy()
base["__k__"] = 1
meas_df = pd.DataFrame({"Measure": desired_measures})
meas_df["__k__"] = 1

grid = base.merge(meas_df, on="__k__", how="outer").drop(columns="__k__")

# Attach aggregated counts to the full grid and fill missing with 0
df_agg_completed = (
    grid
    .merge(df_agg, on=base_keys + ["Measure"], how="left")
    .assign(Count=lambda d: d["Count"].fillna(0).astype(int))
    .sort_values(base_keys + ["Measure"])
    .reset_index(drop=True)
)

# ---------- 7) Optional Pivots (separate sheets for each measure) ----------
pivot_offences = None
pivot_outcomes = None

if "pivot" in SHEETS_TO_EXPORT:
    if "Offences" in df_agg_completed["Measure"].unique():
        src_o = df_agg_completed[df_agg_completed["Measure"].eq("Offences")]
        if {"Month_Year", "Borough", "Offence Group", "Count"}.issubset(src_o.columns):
            pivot_offences = src_o.pivot_table(
                index=["Month_Year", "Borough"],
                columns="Offence Group",
                values="Count",
                aggfunc="sum",
                fill_value=0
            )
    if "Positive Outcomes" in df_agg_completed["Measure"].unique():
        src_p = df_agg_completed[df_agg_completed["Measure"].eq("Positive Outcomes")]
        if {"Month_Year", "Borough", "Offence Group", "Count"}.issubset(src_p.columns):
            pivot_outcomes = src_p.pivot_table(
                index=["Month_Year", "Borough"],
                columns="Offence Group",
                values="Count",
                aggfunc="sum",
                fill_value=0
            )

# ---------- 8) Save ONE Excel workbook to the script's folder ----------
# Always save next to this script (independent of Spyder's working directory)
script_dir = Path(__file__).resolve().parent
script_dir.mkdir(parents=True, exist_ok=True)

final_name = build_output_name(OUTPUT_XLSX_NAME, APPEND_TIMESTAMP)
out_xlsx = script_dir / final_name

with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
    if "agg" in SHEETS_TO_EXPORT:
        df_agg_completed.to_excel(writer, index=False, sheet_name="borough_offenceGroup_agg")
    if "clean" in SHEETS_TO_EXPORT:
        df_borough.to_excel(writer, index=False, sheet_name="borough_clean")
    if "pivot" in SHEETS_TO_EXPORT and pivot_offences is not None:
        pivot_offences.to_excel(writer, sheet_name="pivot_offences")
    if "pivot" in SHEETS_TO_EXPORT and pivot_outcomes is not None:
        pivot_outcomes.to_excel(writer, sheet_name="pivot_positive_outcomes")

print("\nSaved workbook to:", out_xlsx.resolve())

# ---------- 9) Open the folder so you see the file immediately ----------
reveal_in_file_browser(script_dir)

