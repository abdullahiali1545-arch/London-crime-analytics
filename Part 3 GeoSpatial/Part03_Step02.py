
#PART3 STEP 2

import numpy as np
import pandas as pd
from pathlib import Path
import sys
import os
import platform
import subprocess


# ─────────────────────────────────────────────────────────────────────────────
# SETTINGS
# ─────────────────────────────────────────────────────────────────────────────
EXCEL_PATH = Path("M1045_borough_grouped_with_success_rates_v3.xlsx")
SHEET_NAME = 0
OUTPUT_FILE = "London_crime_data.xlsx"   # Excel output (always this name)

# ─────────────────────────────────────────────────────────────────────────────
# NORMALISE MONTH COLUMN
# ─────────────────────────────────────────────────────────────────────────────
def normalize_month_column(df: pd.DataFrame) -> pd.DataFrame:
    s = df["Month_Year"]

    # Handle numeric Excel serials, actual datetimes, or date strings
    if np.issubdtype(s.dtype, np.number):
        month = pd.to_datetime(s, origin="1899-12-30", unit="D")
    elif np.issubdtype(s.dtype, np.datetime64):
        month = s
    else:
        month = pd.to_datetime(s, errors="coerce")

    df = df.copy()
    df["Month"] = month
    df["Month_fmt"] = df["Month"].dt.to_period("M").astype(str)
    return df


# ─────────────────────────────────────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────────────────────────────────────
def load_data(path: Path, sheet=0) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")

    expected = {
        "Month_Year", "Area name", "Offence Group",
        "Offences", "Positive Outcomes", "Success Rate"
    }
    missing = expected - set(df.columns)
    if missing:
        raise ValueError(f"Missing expected columns: {sorted(missing)}")

    return normalize_month_column(df)

# ─────────────────────────────────────────────────────────────────────────────
# REMOVE INVALID BOROUGHS (USING WHITELIST)
# ─────────────────────────────────────────────────────────────────────────────
VALID_BOROUGHS = {
    "barking and dagenham","barnet","bexley","brent","bromley","camden",
    "croydon","ealing","enfield","greenwich","hackney","hammersmith and fulham",
    "haringey","harrow","havering","hillingdon","hounslow","islington",
    "kensington and chelsea","kingston upon thames","lambeth","lewisham",
    "merton","newham","redbridge","richmond upon thames","southwark",
    "sutton","tower hamlets","waltham forest","wandsworth","westminster"
}

def keep_only_valid_boroughs(df: pd.DataFrame) -> pd.DataFrame:
    """Filters dataset so only real London boroughs remain."""
    return df[df["Area name"].str.strip().str.casefold().isin(VALID_BOROUGHS)].copy()

# ─────────────────────────────────────────────────────────────────────────────
# SPYDER‑SAFE INPUT
# ─────────────────────────────────────────────────────────────────────────────
def prompt_from_list(prompt_text: str, options: list[str]) -> str:
    while True:
        print(f"\n{prompt_text}")
        for i, opt in enumerate(options, start=1):
            print(f"  {i}. {opt}")

        # Flush so Spyder shows the prompt
        sys.stdout.write("Type number (or value) and press Enter: ")
        sys.stdout.flush()
        raw = input().strip()

        # Numeric choice
        if raw.isdigit():
            idx = int(raw)
            if 1 <= idx <= len(options):
                return options[idx - 1]
            print(f"Please enter a number between 1 and {len(options)}.")
            continue

        # Exact text match (case-insensitive)
        matches = [opt for opt in options if opt.lower() == raw.lower()]
        if matches:
            return matches[0]

        print("Sorry, that wasn't recognised. Please try again.")

# ─────────────────────────────────────────────────────────────────────────────
# MENU CREATION
# ─────────────────────────────────────────────────────────────────────────────
def build_menus(df: pd.DataFrame):
    months = df["Month_fmt"].dropna().sort_values().unique().tolist()
    offences = df["Offence Group"].dropna().sort_values().unique().tolist()
    measures = ["Offences", "Positive Outcomes", "Success Rate"]
    return months, offences, measures

# ─────────────────────────────────────────────────────────────────────────────
# EXPORT EXACTLY WHAT IS PRINTED
# ─────────────────────────────────────────────────────────────────────────────
def save_output_excel(table_to_save: pd.DataFrame):
    """
    Save the exact table that is printed to terminal to Excel,
    preserving column order and values.
    """
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        table_to_save.to_excel(writer, index=False, sheet_name="data")

    print(f"\nSaved Excel file → {OUTPUT_FILE}")

  

# ─────────────────────────────────────────────────────────────────────────────
# MAIN INTERACTION
# ─────────────────────────────────────────────────────────────────────────────
def run_interactive():
    print("Loading data...")
    df = load_data(EXCEL_PATH, sheet=SHEET_NAME)
    df = keep_only_valid_boroughs(df)
    month_options, offence_options, measure_options = build_menus(df)

    chosen_month   = prompt_from_list("Select Month & Year", month_options)
    chosen_offence = prompt_from_list("Select Offence Group", offence_options)
    chosen_measure = prompt_from_list("Select Measure", measure_options)

    # Build the exact table we want to show AND save
    # (Columns and order here define both terminal & Excel outputs)
    columns_for_view_and_export = [
        "Month_fmt",
        "Area name",
        "Offence Group",
        chosen_measure
        
       
    ]

    filtered = df.loc[
        (df["Month_fmt"] == chosen_month)
        & (df["Offence Group"].str.casefold() == chosen_offence.casefold()),
        columns_for_view_and_export
    ].copy()

    if filtered.empty:
        print("\nNo rows found for that combination. Try again.")
        return

    # Print exactly what we'll export
    print("\n──────────────── RESULT (first 34 rows) ────────────────")
    print(filtered.head(34).to_string(index=False))

    # Save EXACTLY the same table (all rows, same columns, same order)
    save_output_excel(filtered)

# ─────────────────────────────────────────────────────────────────────────────
# PROGRAM ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    run_interactive()