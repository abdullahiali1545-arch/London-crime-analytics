# -*- coding: utf-8 -*-
"""
Part 2 – Step 04
Improved Single Borough Scatterplot
With proper month intervals and axis units
"""

import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# -------------------------------------------------------
# STEP 1: Load dataset
# -------------------------------------------------------

INPUT_FILE = "M1045_borough_grouped_with_success_rates_v3.xlsx"

print("Loading dataset...")
df = pd.read_excel(INPUT_FILE)

df["Month_Year"] = pd.to_datetime(df["Month_Year"])

print("Dataset loaded successfully.\n")

# -------------------------------------------------------
# STEP 2: User selections
# -------------------------------------------------------

boroughs = sorted(df["Area name"].unique())
print("Available Boroughs:\n")
for i, b in enumerate(boroughs):
    print(f"{i}: {b}")

borough_choice = int(input("\nEnter borough number: "))
selected_borough = boroughs[borough_choice]

offence_groups = sorted(df["Offence Group"].unique())
print("\nAvailable Offence Groups:\n")
for i, o in enumerate(offence_groups):
    print(f"{i}: {o}")

offence_choice = int(input("\nEnter offence group number: "))
selected_offence = offence_groups[offence_choice]

measures = ["Offences", "Positive Outcomes", "Success Rate"]
print("\nAvailable Measures:\n")
for i, m in enumerate(measures):
    print(f"{i}: {m}")

measure_choice = int(input("\nEnter measure number: "))
selected_measure = measures[measure_choice]

# -------------------------------------------------------
# STEP 3: Filter data
# -------------------------------------------------------

filtered_df = df[
    (df["Area name"] == selected_borough) &
    (df["Offence Group"] == selected_offence)
].sort_values("Month_Year")

# -------------------------------------------------------
# STEP 4: Plot
# -------------------------------------------------------

plt.figure(figsize=(12, 6))

sns.scatterplot(
    data=filtered_df,
    x="Month_Year",
    y=selected_measure
)

# --- Monthly tick intervals ---
ax = plt.gca()
ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))  # every 3 months
ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))

plt.xticks(rotation=45)

# --- Axis labels with units ---
plt.xlabel("Month (Year)")

if selected_measure == "Success Rate":
    plt.ylabel("Success Rate (Proportion 0–1)")
else:
    plt.ylabel("Number of Recorded Cases")

plt.title(f"{selected_measure} Over Time\n{selected_borough} – {selected_offence}")

plt.tight_layout()

# Save file
output_filename = f"scatter_{selected_borough}_{selected_offence}_{selected_measure}.png"
plt.savefig(output_filename)

plt.show()

print(f"\nPlot saved as: {output_filename}")
print("Process complete.")