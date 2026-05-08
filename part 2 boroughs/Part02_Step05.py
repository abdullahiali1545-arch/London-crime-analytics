# m1045_full_pipeline_dual_borough_plot_loop_demeaned.py
# End-to-end pipeline (Stages 1–3 unchanged from your original),
# plus Stage 4 plotting that:
#   • lets the user choose 1 or 2 boroughs
#   • de-means each borough’s selected series before plotting
#   • computes correlation on the de-meaned series
#   • loops until the user opts to exit (y/n)
#   • validates inputs to avoid crashes

import os
import platform
from pathlib import Path
from datetime import datetime
import sys
import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

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


# ========================= STAGE 1: CLEAN + ROLLUP =========================

# ---------- 1) Load ----------
try:
    df = pd.read_excel(INPUT_XLSX, sheet_name=SHEET_INDEX, engine="openpyxl")
except FileNotFoundError:
    print(f"Error: Input file not found at: {INPUT_XLSX}")
    sys.exit(1)

# tidy headers
df.columns = [str(c).strip() for c in df.columns]

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
desired_measures = ["Offences", "Positive Outcomes"]

base_keys = [c for c in [
    "Month_Year", "Borough", "Area name", "Area Type",
    "Offence Group", "Financial Year"
] if c in df_agg.columns]

base = df_agg[base_keys].drop_duplicates().copy()
base["__k__"] = 1
meas_df = pd.DataFrame({"Measure": desired_measures})
meas_df["__k__"] = 1
grid = base.merge(meas_df, on="__k__", how="outer").drop(columns="__k__")

df_agg_completed = (
    grid
    .merge(df_agg, on=base_keys + ["Measure"], how="left")
    .assign(Count=lambda d: d["Count"].fillna(0).astype(int))
    .sort_values(base_keys + ["Measure"])
    .reset_index(drop=True)
)

# ---------- 7) Optional Pivots ----------
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

# ---------- 8) Save Stage 1 ----------
script_dir = Path(__file__).resolve().parent if '__file__' in globals() else Path.cwd()
script_dir.mkdir(parents=True, exist_ok=True)

final_name = build_output_name(OUTPUT_XLSX_NAME, APPEND_TIMESTAMP)
stage1_xlsx = script_dir / final_name

with pd.ExcelWriter(stage1_xlsx, engine="openpyxl") as writer:
    if "agg" in SHEETS_TO_EXPORT:
        df_agg_completed.to_excel(writer, index=False, sheet_name="borough_offenceGroup_agg")
    if "clean" in SHEETS_TO_EXPORT:
        df_borough.to_excel(writer, index=False, sheet_name="borough_clean")
    if "pivot" in SHEETS_TO_EXPORT and pivot_offences is not None:
        pivot_offences.to_excel(writer, sheet_name="pivot_offences")
    if "pivot" in SHEETS_TO_EXPORT and pivot_outcomes is not None:
        pivot_outcomes.to_excel(writer, sheet_name="pivot_positive_outcomes")

print("\n[Stage 1] Saved workbook to:", stage1_xlsx.resolve())
reveal_in_file_browser(script_dir)


# ========================= STAGE 2: ADD TOTALS =========================

stage2_xlsx = script_dir / "M1045_borough_grouped_with_totals_v2.xlsx"

try:
    df = pd.read_excel(stage1_xlsx, engine="openpyxl")
except FileNotFoundError:
    print(f"Error: File '{stage1_xlsx}' not found.")
    sys.exit(1)

df.columns = df.columns.str.strip()

if "Borough" in df.columns:
    borough_col = "Borough"
elif "Area name" in df.columns:
    borough_col = "Area name"
else:
    print("Error: Could not find 'Borough' or 'Area name' column.")
    sys.exit(1)

required_cols = ["Month_Year", borough_col, "Offence Group", "Measure", "Count"]
for col in required_cols:
    if col not in df.columns:
        print(f"Error: Required column '{col}' is missing.")
        sys.exit(1)

df = df[df["Offence Group"] != "Total"].copy()

totals = (
    df.groupby(["Month_Year", borough_col, "Measure"], as_index=False)
      .agg({"Count": "sum"})
)
totals["Offence Group"] = "Total"

other_cols = [col for col in df.columns if col not in totals.columns]
additional_info = (
    df.groupby(["Month_Year", borough_col], as_index=False)
      .first()[["Month_Year", borough_col] + other_cols]
)
totals = totals.merge(additional_info, on=["Month_Year", borough_col], how="left")

totals = totals[df.columns]
df_final = pd.concat([df, totals], ignore_index=True)

df_final = df_final.sort_values(by=["Month_Year", borough_col, "Offence Group", "Measure"]).reset_index(drop=True)

try:
    df_final.to_excel(stage2_xlsx, index=False)
    print(f"[Stage 2] Success! File saved as '{stage2_xlsx.name}'")
except Exception as e:
    print("Error while saving file:")
    print(e)
    sys.exit(1)


# ========================= STAGE 3: SUCCESS RATES (v3) =========================

stage3_xlsx = script_dir / "M1045_borough_grouped_with_success_rates_v3.xlsx"

print("Loading dataset for success rates...")
df = pd.read_excel(stage2_xlsx, engine="openpyxl")
print("Dataset loaded successfully.\n")

required_columns = ["Month_Year", "Area name", "Offence Group", "Measure", "Count"]
for col in required_columns:
    if col not in df.columns:
        raise ValueError(f"Missing expected column: {col}")
print("All required columns are present.")

print("\nReshaping dataset (pivoting)...")
pivot_df = df.pivot_table(
    index=["Month_Year", "Area name", "Offence Group"],
    columns="Measure",
    values="Count",
    aggfunc="sum"
).reset_index()
pivot_df.columns.name = None
print("Pivot complete.")

if "Positive Outcomes" not in pivot_df.columns:
    pivot_df["Positive Outcomes"] = 0
pivot_df["Positive Outcomes"] = pivot_df["Positive Outcomes"].fillna(0)

if "Offences" not in pivot_df.columns:
    pivot_df["Offences"] = 0
pivot_df["Offences"] = pivot_df["Offences"].fillna(0)
print("Missing values handled.")

print("\nCalculating Success Rate...")
pivot_df["Success Rate"] = np.where(
    pivot_df["Offences"] == 0,
    0,
    pivot_df["Positive Outcomes"] / pivot_df["Offences"]
)
print("Success Rate column created.")

pivot_df = pivot_df.sort_values(by=["Month_Year", "Area name", "Offence Group"])
pivot_df.to_excel(stage3_xlsx, index=False)
print(f"\n[Stage 3] File successfully saved as: {stage3_xlsx.name}")

print("\nFinal dataset info:")
print(pivot_df.info())
print("\nProcess complete (success rates).")


# ========================= STAGE 4: PLOTTING (loop; de-meaned) =========================

print("\nLoading dataset for plotting...")
plot_df = pd.read_excel(stage3_xlsx, engine="openpyxl")
plot_df["Month_Year"] = pd.to_datetime(plot_df["Month_Year"], errors="coerce")
print("Dataset ready for plotting.\n")

boroughs = sorted(plot_df["Area name"].dropna().unique())
offence_groups = sorted(plot_df["Offence Group"].dropna().unique())
measures = ["Offences", "Positive Outcomes", "Success Rate"]

def choose_from_list(options, prompt_text):
    """Robust selection helper that lists options and re-prompts on bad input."""
    while True:
        print(f"\n{prompt_text}")
        for i, opt in enumerate(options):
            print(f"{i}: {opt}")
        raw = input("Enter number: ").strip()
        try:
            idx = int(raw)
            if 0 <= idx < len(options):
                return options[idx]
            print("Invalid number. Please try again.")
        except ValueError:
            print("Invalid input. Please enter a number.")

def choose_borough(prompt_text):
    return choose_from_list(boroughs, prompt_text)

def choose_borough_count():
    while True:
        raw = input("\nHow many boroughs do you want to plot? (enter 1 or 2): ").strip()
        if raw in {"1", "2"}:
            return int(raw)
        print("Please enter 1 or 2.")

def ask_yes_no(prompt):
    """Return True for yes, False for no. Re-prompts on bad input."""
    while True:
        raw = input(prompt).strip().lower()
        if raw in {"y", "yes"}:
            return True
        if raw in {"n", "no"}:
            return False
        print("Please enter 'y' or 'n'.")

def get_series_for(borough, offence, measure):
    """Return Month_Year + measure series for a borough/offence (numeric)."""
    if borough is None:
        return pd.DataFrame(columns=["Month_Year", measure])
    out = (plot_df[(plot_df["Area name"] == borough) & (plot_df["Offence Group"] == offence)]
           .loc[:, ["Month_Year", measure]]
           .dropna(subset=["Month_Year"])
           .sort_values("Month_Year")
           .reset_index(drop=True))
    # ensure numeric for the measure
    out[measure] = pd.to_numeric(out[measure], errors="coerce")
    return out.dropna(subset=[measure])

def demean_series(df_series, measure):
    """Add a '{measure}_demeaned' column = value - series mean; return df and the mean."""
    if df_series.empty:
        return df_series.assign(**{f"{measure}_demeaned": []}), np.nan
    mu = df_series[measure].mean(skipna=True)
    df_dm = df_series.copy()
    df_dm[f"{measure}_demeaned"] = df_dm[measure] - mu
    return df_dm, mu

# ---- LOOP until user chooses to exit ----
while True:
    # 1 or 2 boroughs
    num_boros = choose_borough_count()

    # Select borough(s)
    borough_A = choose_borough("Available Boroughs – select FIRST borough")
    borough_B = None
    if num_boros == 2:
        while True:
            borough_B = choose_borough("Available Boroughs – select SECOND borough")
            if borough_B != borough_A:
                break
            print("Second borough must be different from the first. Please choose again.")

    # Offence group & measure
    selected_offence = choose_from_list(offence_groups, "Available Offence Groups – select one")
    selected_measure = choose_from_list(measures, "Available Measures – select one")

    # Prepare data and de-mean per borough
    df_A = get_series_for(borough_A, selected_offence, selected_measure)
    if df_A.empty:
        print(f"\nNo data found for {borough_A} – {selected_offence} – {selected_measure}. Try again.")
        if not ask_yes_no("Plot another? (y/n): "):
            print("Exiting plotting loop. All steps complete.")
            break
        plt.close('all')
        continue

    df_A_dm, muA = demean_series(df_A, selected_measure)

    df_B_dm, muB = None, np.nan
    if num_boros == 2:
        df_B = get_series_for(borough_B, selected_offence, selected_measure)
        if df_B.empty:
            print(f"\nNo data found for {borough_B} – {selected_offence} – {selected_measure}. Try again.")
            if not ask_yes_no("Plot another? (y/n): "):
                print("Exiting plotting loop. All steps complete.")
                break
            plt.close('all')
            continue
        df_B_dm, muB = demean_series(df_B, selected_measure)

    # Correlation on de-meaned values (only when two boroughs)
    correlation = np.nan
    if num_boros == 2 and df_B_dm is not None and not df_B_dm.empty:
        merged = pd.merge(
            df_A_dm.rename(columns={f"{selected_measure}_demeaned": f"dm_{borough_A}"}),
            df_B_dm.rename(columns={f"{selected_measure}_demeaned": f"dm_{borough_B}"}),
            on="Month_Year", how="inner"
        )
        if len(merged) >= 2:
            a = merged[f"dm_{borough_A}"].astype(float)
            b = merged[f"dm_{borough_B}"].astype(float)
            if a.nunique(dropna=True) > 1 and b.nunique(dropna=True) > 1:
                correlation = a.corr(b)
            else:
                correlation = np.nan
        if np.isnan(correlation):
            print("\nCorrelation (demeaned) could not be computed (insufficient or constant aligned data).")
        else:
            print(f"\nCorrelation coefficient (demeaned) between {borough_A} and {borough_B}: {correlation:.4f}")

    # Plot (demeaned)
    plt.figure(figsize=(12, 6))
    sns.scatterplot(
        data=df_A_dm, x="Month_Year", y=f"{selected_measure}_demeaned",
        label=f"{borough_A} (μ={muA:.3f})", s=80
    )
    if num_boros == 2 and df_B_dm is not None and not df_B_dm.empty:
        sns.scatterplot(
            data=df_B_dm, x="Month_Year", y=f"{selected_measure}_demeaned",
            label=f"{borough_B} (μ={muB:.3f})", s=80
        )

    # Zero line for reference
    ax = plt.gca()
    ax.axhline(0, color="gray", linewidth=1, linestyle="--", alpha=0.7)

    # X-axis formatting
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))
    plt.xticks(rotation=45)
    plt.xlabel("Month (Year)")

    # Y-axis label
    if selected_measure == "Success Rate":
        plt.ylabel("De-meaned Success Rate (difference from borough mean, proportion)")
    else:
        plt.ylabel("De-meaned Value (difference from borough mean, count)")

    # Title
    if num_boros == 2:
        corr_text = "" if np.isnan(correlation) else f" | r (demeaned) = {correlation:.3f}"
        plt.title(f"Demeaned {selected_measure} Over Time – {selected_offence}\n{borough_A} vs {borough_B}{corr_text}")
    else:
        plt.title(f"Demeaned {selected_measure} Over Time – {selected_offence}\n{borough_A}")

    plt.legend()
    plt.tight_layout()

    # Save
    if num_boros == 2:
        output_filename = f"scatter_demeaned_{borough_A}_vs_{borough_B}_{selected_offence}_{selected_measure}.png"
    else:
        output_filename = f"scatter_demeaned_{borough_A}_{selected_offence}_{selected_measure}.png"

    plt.savefig(output_filename, dpi=150)
    plt.show()
    print(f"\nPlot saved as: {output_filename}")

    # Ask to loop again or exit
    if not ask_yes_no("Plot another? (y/n): "):
        print("Exiting plotting loop. All steps complete.")
        break
    plt.close('all')
