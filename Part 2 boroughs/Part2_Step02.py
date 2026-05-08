import pandas as pd
import sys
from pathlib import Path

# -------------------------------------------------------
# CONFIGURATION
# -------------------------------------------------------
INPUT_FILE = "M1045_borough_grouped_two_rows_per_group.xlsx"      # Change this
OUTPUT_FILE = "M1045_borough_grouped_with_totals_v2.xlsx"

# -------------------------------------------------------
# LOAD DATA
# -------------------------------------------------------
try:
    df = pd.read_excel(INPUT_FILE)
except FileNotFoundError:
    print(f"Error: File '{INPUT_FILE}' not found.")
    sys.exit(1)

# -------------------------------------------------------
# CLEAN COLUMN NAMES (remove hidden whitespace)
# -------------------------------------------------------
df.columns = df.columns.str.strip()

# -------------------------------------------------------
# DETECT BOROUGH COLUMN
# -------------------------------------------------------
if "Borough" in df.columns:
    borough_col = "Borough"
elif "Area name" in df.columns:
    borough_col = "Area name"
else:
    print("Error: Could not find 'Borough' or 'Area name' column.")
    sys.exit(1)

# -------------------------------------------------------
# VALIDATE REQUIRED COLUMNS
# -------------------------------------------------------
required_cols = [
    "Month_Year",
    borough_col,
    "Offence Group",
    "Measure",
    "Count"
]

for col in required_cols:
    if col not in df.columns:
        print(f"Error: Required column '{col}' is missing.")
        sys.exit(1)

# -------------------------------------------------------
# REMOVE EXISTING 'Total' ROWS
# -------------------------------------------------------
df = df[df["Offence Group"] != "Total"].copy()

# -------------------------------------------------------
# CALCULATE TOTALS
# IMPORTANT: Do NOT group by Offence Group
# -------------------------------------------------------
totals = (
    df.groupby(["Month_Year", borough_col, "Measure"], as_index=False)
      .agg({"Count": "sum"})
)

# Add Offence Group column as 'Total'
totals["Offence Group"] = "Total"

# -------------------------------------------------------
# ADD BACK OTHER COLUMNS (Area Type, Financial Year)
# These must be preserved in output
# -------------------------------------------------------
other_cols = [col for col in df.columns if col not in totals.columns]

# For remaining columns, take first valid value per group
additional_info = (
    df.groupby(["Month_Year", borough_col], as_index=False)
      .first()[["Month_Year", borough_col] + other_cols]
)

# Merge additional info into totals
totals = totals.merge(
    additional_info,
    on=["Month_Year", borough_col],
    how="left"
)

# -------------------------------------------------------
# REORDER COLUMNS TO MATCH ORIGINAL ORDER
# -------------------------------------------------------
totals = totals[df.columns]

# -------------------------------------------------------
# APPEND TOTALS TO ORIGINAL DATA
# -------------------------------------------------------
df_final = pd.concat([df, totals], ignore_index=True)

# -------------------------------------------------------
# SORT OUTPUT
# -------------------------------------------------------
df_final = df_final.sort_values(
    by=["Month_Year", borough_col, "Offence Group", "Measure"]
).reset_index(drop=True)

# -------------------------------------------------------
# SAVE TO NEW EXCEL FILE
# -------------------------------------------------------
try:
    df_final.to_excel(OUTPUT_FILE, index=False)
    print(f"Success! File saved as '{OUTPUT_FILE}'")
except Exception as e:
    print("Error while saving file:")
    print(e)