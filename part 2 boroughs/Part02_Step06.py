

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# -------------------------------
# LOAD CLEANED DATA (Step 3 file)
# -------------------------------
file_path = "M1045_borough_grouped_with_success_rates_v3.xlsx"  #change to your file path
df = pd.read_excel(file_path)

df["Month_Year"] = pd.to_datetime(df["Month_Year"])

# -------------------------------
# FILTER TOTAL CRIME ONLY
# -------------------------------
total_df = df[df["Offence Group"] == "Total"].copy()

# -------------------------------
# PIVOT DATA (Date x Borough)
# -------------------------------
pivot = total_df.pivot(
    index="Month_Year",
    columns="Area name",
    values="Success Rate"
)

# -------------------------------
# DE-MEAN DATA
# -------------------------------
demeaned = pivot - pivot.mean()

# -------------------------------
# ASK USER FOR BOROUGH (case-insensitive)
# -------------------------------
print("\nAvailable boroughs:")
print(list(demeaned.columns))

user_input = input("\nEnter a borough name as shown above: ")

# Convert user input to match the column names (case-insensitive)
chosen_borough = None
for col in demeaned.columns:
    if col.lower() == user_input.lower():
        chosen_borough = col
        break

if chosen_borough is None:
    raise ValueError("Invalid borough name entered.")

# -------------------------------
# CALCULATE CORRELATIONS
# -------------------------------
correlations = {}

for borough in demeaned.columns:
    if borough != chosen_borough:
        corr = demeaned[chosen_borough].corr(demeaned[borough])
        correlations[borough] = corr

corr_df = pd.DataFrame.from_dict(
    correlations,
    orient="index",
    columns=[f"Correlation with {chosen_borough}"]
).sort_values(by=f"Correlation with {chosen_borough}", ascending=False)

# Add proper title for the first column
corr_df.index.name = "Borough"

print("\nCorrelation Coefficients:")
print(corr_df)

# -------------------------------
# FIND HIGHEST & LOWEST
# -------------------------------
highest_borough = corr_df.idxmax()[0]
lowest_borough = corr_df.idxmin()[0]

print("\nHighest correlation:", highest_borough)
print("Lowest correlation:", lowest_borough)

# -------------------------------
# PLOT HIGHEST
# -------------------------------
plt.figure()
plt.plot(demeaned.index, demeaned[chosen_borough], label=chosen_borough)
plt.plot(demeaned.index, demeaned[highest_borough], label=highest_borough)
plt.title(f"De-meaned Success Rate: {chosen_borough} vs {highest_borough}")
plt.legend()
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# -------------------------------
# PLOT LOWEST
# -------------------------------
plt.figure()
plt.plot(demeaned.index, demeaned[chosen_borough], label=chosen_borough)
plt.plot(demeaned.index, demeaned[lowest_borough], label=lowest_borough)
plt.title(f"De-meaned Success Rate: {chosen_borough} vs {lowest_borough}")
plt.legend()
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# -------------------------------
# SAVE RESULTS
# -------------------------------
corr_df.to_excel("correlation_results.xlsx")

print("\nResults saved to correlation_results.xlsx")
