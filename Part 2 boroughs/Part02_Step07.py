
# ========================= STAGE 4: CORRELATION (NO PLOTS) =========================

# ========================= CORRELATION: DEMEANED TOTAL SUCCESS RATE =========================

import pandas as pd
import numpy as np

# ---- FILE PATH (EDIT IF NEEDED) ----
excel_path = "M1045_borough_grouped_with_success_rates_v3.xlsx"

print("\nLoading dataset...")
df = pd.read_excel(excel_path, engine="openpyxl")

# Ensure proper datetime format
df["Month_Year"] = pd.to_datetime(df["Month_Year"], errors="coerce")

# Keep TOTAL crime only
df_total = df[df["Offence Group"] == "Total"].copy()

if df_total.empty:
    raise ValueError("No 'Total' offence group found in dataset.")

# Get list of boroughs
boroughs = sorted(df_total["Area name"].dropna().unique())

# ---- User selects borough ----
print("\nAvailable Boroughs:")
for i, b in enumerate(boroughs):
    print(f"{i}: {b}")

while True:
    try:
        idx = int(input("\nSelect a borough (enter number): "))
        if 0 <= idx < len(boroughs):
            selected_borough = boroughs[idx]
            break
        else:
            print("Invalid number. Try again.")
    except ValueError:
        print("Please enter a valid number.")

print(f"\nSelected borough: {selected_borough}")

# ---- Selected borough series ----
base_series = (
    df_total[df_total["Area name"] == selected_borough]
    [["Month_Year", "Success Rate"]]
    .dropna()
    .sort_values("Month_Year")
)

# De-mean selected borough
base_mean = base_series["Success Rate"].mean()
base_series["demeaned"] = base_series["Success Rate"] - base_mean

# ---- Compute correlation vs all other boroughs ----
results = []

for borough in boroughs:
    if borough == selected_borough:
        continue

    compare_series = (
        df_total[df_total["Area name"] == borough]
        [["Month_Year", "Success Rate"]]
        .dropna()
        .sort_values("Month_Year")
    )

    # De-mean comparison borough
    comp_mean = compare_series["Success Rate"].mean()
    compare_series["demeaned"] = compare_series["Success Rate"] - comp_mean

    # Align by Month_Year
    merged = pd.merge(
        base_series[["Month_Year", "demeaned"]],
        compare_series[["Month_Year", "demeaned"]],
        on="Month_Year",
        suffixes=("_base", "_comp"),
        how="inner"
    )

    if len(merged) >= 2:
        r = merged["demeaned_base"].corr(merged["demeaned_comp"])
    else:
        r = np.nan

    results.append((borough, r))

# ---- Create results table ----
corr_df = pd.DataFrame(
    results,
    columns=["Other Borough", "Correlation (demeaned Total Success Rate)"]
)

corr_df = corr_df.sort_values(
    by="Correlation (demeaned Total Success Rate)",
    ascending=False
)

print("\nCorrelation results (highest to lowest):\n")
print(corr_df.to_string(index=False))

print("\nAnalysis complete.")

# ========================= CUBIC REGRESSION WITH GRAPH =========================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.preprocessing import PolynomialFeatures
from sklearn.linear_model import LinearRegression
from sklearn.model_selection import train_test_split
from sklearn.metrics import r2_score, mean_squared_error

# ---- FILE PATH (EDIT IF NEEDED) ----
excel_path = "M1045_borough_grouped_with_success_rates_v3.xlsx"

print("Loading dataset...")
df = pd.read_excel(excel_path, engine="openpyxl")

df["Month_Year"] = pd.to_datetime(df["Month_Year"], errors="coerce")

# Keep TOTAL crime only
df = df[df["Offence Group"] == "Total"].copy()

if df.empty:
    raise ValueError("No 'Total' offence group found.")

# ---- Choose Borough ----
boroughs = sorted(df["Area name"].dropna().unique())

print("\nAvailable Boroughs:")
for i, b in enumerate(boroughs):
    print(f"{i}: {b}")

while True:
    try:
        idx = int(input("\nSelect a borough (enter number): "))
        if 0 <= idx < len(boroughs):
            selected_borough = boroughs[idx]
            break
        else:
            print("Invalid number. Try again.")
    except ValueError:
        print("Please enter a valid number.")

print(f"\nSelected borough: {selected_borough}")

# ---- Prepare Borough Data ----
borough_df = df[df["Area name"] == selected_borough].copy()
borough_df = borough_df.sort_values("Month_Year")
borough_df = borough_df.dropna(subset=["Success Rate"])

if len(borough_df) < 10:
    raise ValueError("Not enough data points for modelling.")

# ---- De-mean Success Rate ----
mean_sr = borough_df["Success Rate"].mean()
borough_df["demeaned_sr"] = borough_df["Success Rate"] - mean_sr

# ---- Create numeric time index ----
borough_df["time_index"] = np.arange(len(borough_df))

X = borough_df[["time_index"]].values
y = borough_df["demeaned_sr"].values

# ---- 75% Train / 25% Test split ----
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.25, shuffle=False
)

# ---- Polynomial (Cubic) Features ----
poly = PolynomialFeatures(degree=3)
X_train_poly = poly.fit_transform(X_train)
X_test_poly = poly.transform(X_test)

# ---- Fit Model ----
model = LinearRegression()
model.fit(X_train_poly, y_train)

# ---- Predictions ----
y_pred_test = model.predict(X_test_poly)

# Full curve for smooth plotting
X_full_poly = poly.transform(X)
y_full_pred = model.predict(X_full_poly)

# ---- Evaluation ----
r2 = r2_score(y_test, y_pred_test)
rmse = np.sqrt(mean_squared_error(y_test, y_pred_test))

# ---- Extract Equation ----
intercept = model.intercept_
coefs = model.coef_

equation = (
    f"y = {intercept:.6f} "
    f"+ {coefs[1]:.6f}x "
    f"+ {coefs[2]:.6f}x^2 "
    f"+ {coefs[3]:.6f}x^3"
)

print("\nCubic Equation (demeaned success rate):")
print(equation)
print(f"\nTest R²: {r2:.4f}")
print(f"Test RMSE: {rmse:.6f}")

# ========================= PLOT =========================

plt.figure(figsize=(12, 6))

# Actual demeaned data
# plt.scatter(borough_df["Month_Year"], y, label="Actual (demeaned)", s=60)
# ========================= PLOT =========================

plt.figure(figsize=(12, 6))

# Training data (75%) - ORANGE
plt.scatter(
    borough_df["Month_Year"].iloc[:len(X_train)],
    y_train,
    color="orange",
    label="Training Data (75%)",
    s=60
)

# Test data (25%) - BLUE
plt.scatter(
    borough_df["Month_Year"].iloc[len(X_train):],
    y_test,
    color="blue",
    label="Test Data (25%)",
    s=60
)

# Cubic regression line
plt.plot(borough_df["Month_Year"], y_full_pred, linewidth=3, label="Cubic Fit")

plt.axhline(0, linestyle="--", linewidth=1)
plt.title(f"Cubic Regression – Demeaned Total Success Rate\n{selected_borough}")
plt.xlabel("Month")
plt.ylabel("Demeaned Success Rate")
plt.legend()
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

print("\nGraph displayed in Spyder.")