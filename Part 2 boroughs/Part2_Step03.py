
"""
Part 2 – Step 03
Reshape borough crime data and calculate Success Rates.

This script:
1. Loads the cleaned borough dataset (v2)
2. Converts long format into wide format
3. Creates Offences and Positive Outcomes columns
4. Calculates Success Rate safely
5. Saves new version (v3)
"""

import pandas as pd
import numpy as np

# -------------------------------------------------------
# STEP 1: File names (must be in same directory)
# -------------------------------------------------------

INPUT_FILE = "M1045_borough_grouped_with_totals_v2.xlsx"
OUTPUT_FILE = "M1045_borough_grouped_with_success_rates_v3.xlsx"

# -------------------------------------------------------
# STEP 2: Load dataset
# -------------------------------------------------------

print("Loading dataset...")

df = pd.read_excel(INPUT_FILE)

print("Dataset loaded successfully.\n")

print("Columns in dataset:")
print(df.columns)
print("\nFirst 5 rows:")
print(df.head())

# -------------------------------------------------------
# STEP 3: Validate required columns exist
# -------------------------------------------------------

required_columns = [
    "Month_Year",
    "Area name",
    "Offence Group",
    "Measure",
    "Count"
]

for col in required_columns:
    if col not in df.columns:
        raise ValueError(f"Missing expected column: {col}")

print("\nAll required columns are present.")

# -------------------------------------------------------
# STEP 4: Pivot (Long format → Wide format)
# -------------------------------------------------------

print("\nReshaping dataset (pivoting)...")

pivot_df = df.pivot_table(
    index=["Month_Year", "Area name", "Offence Group"],
    columns="Measure",
    values="Count",
    aggfunc="sum"
).reset_index()

# Remove automatic column grouping name
pivot_df.columns.name = None

print("Pivot complete.\n")
print("Preview after pivot:")
print(pivot_df.head())

# -------------------------------------------------------
# STEP 5: Handle missing values
# -------------------------------------------------------

# If a borough had zero positive outcomes,
# the column may contain NaN — replace with 0

if "Positive Outcomes" not in pivot_df.columns:
    pivot_df["Positive Outcomes"] = 0

pivot_df["Positive Outcomes"] = pivot_df["Positive Outcomes"].fillna(0)

# If Offences somehow missing (unlikely but safe)
if "Offences" not in pivot_df.columns:
    pivot_df["Offences"] = 0

pivot_df["Offences"] = pivot_df["Offences"].fillna(0)

print("\nMissing values handled.")

# -------------------------------------------------------
# STEP 6: Calculate Success Rate safely
# -------------------------------------------------------

print("\nCalculating Success Rate...")

pivot_df["Success Rate"] = np.where(
    pivot_df["Offences"] == 0,
    0,
    pivot_df["Positive Outcomes"] / pivot_df["Offences"]
)

print("Success Rate column created.")

# -------------------------------------------------------
# STEP 7: Optional – Sort for neatness
# -------------------------------------------------------

pivot_df = pivot_df.sort_values(
    by=["Month_Year", "Area name", "Offence Group"]
)

# -------------------------------------------------------
# STEP 8: Save new version of file
# -------------------------------------------------------

pivot_df.to_excel(OUTPUT_FILE, index=False)

print(f"\nFile successfully saved as: {OUTPUT_FILE}")

# -------------------------------------------------------
# STEP 9: Final checks (for testing explanation)
# -------------------------------------------------------

print("\nFinal dataset info:")
print(pivot_df.info())

print("\nProcess complete.")