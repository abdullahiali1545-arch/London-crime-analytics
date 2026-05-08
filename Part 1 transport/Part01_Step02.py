
"""
Interactive Transport Crime Analysis Tool
Works with the CLEANED dataset produced by the cleaning script.

Dataset:
public-transport-crime-london.xlsx
Sheet: Clean data
"""

import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import sys


# -------------------------------------------------
# LOAD CLEANED DATASET
# -------------------------------------------------

file = "public-transport-crime-london.xlsx"
sheet = "Clean data"

df = pd.read_excel(file, sheet_name=sheet)

# Ensure Date column is datetime
df["Date"] = pd.to_datetime(df["Date"])


# -------------------------------------------------
# IDENTIFY RATE AND VOLUME COLUMNS
# -------------------------------------------------

rate_cols = [c for c in df.columns if c.startswith("Crime_Rate_")]
volume_cols = [c for c in df.columns if c.startswith("Crime_Volume_")]

# Extract transport mode names
rate_modes = [c.replace("Crime_Rate_", "") for c in rate_cols]
volume_modes = [c.replace("Crime_Volume_", "") for c in volume_cols]

# Union of all transport modes
all_modes = sorted(set(rate_modes + volume_modes))


# -------------------------------------------------
# MENU
# -------------------------------------------------

def print_menu():

    print("\n==============================")
    print(" PUBLIC TRANSPORT CRIME TOOL ")
    print("==============================")
    print("1. Crime RATES over time")
    print("2. Crime VOLUMES over time")
    print("3. Compare two transport modes (correlation)")
    print("4. Exit")
    print("==============================")


# -------------------------------------------------
# MODE SELECTION
# -------------------------------------------------

def choose_modes(available_modes):

    print("\nAvailable transport modes:\n")

    for m in available_modes:
        print("-", m)

    selected = input("\nEnter modes (comma separated): ")

    selected_list = [x.strip() for x in selected.split(",")]

    valid = [m for m in selected_list if m in available_modes]

    if not valid:
        print("\nNo valid modes selected.")
        return None

    return valid


# -------------------------------------------------
# CRIME RATE PLOT
# -------------------------------------------------

def plot_rates():

    modes = choose_modes(rate_modes)

    if modes is None:
        return

    long_df = df.melt(
        id_vars="Date",
        value_vars=[f"Crime_Rate_{m}" for m in modes],
        var_name="Transport Mode",
        value_name="Crime Rate"
    )

    long_df["Transport Mode"] = long_df["Transport Mode"].str.replace("Crime_Rate_", "")

    plt.figure(figsize=(14,7))

    ax = sns.scatterplot(
        data=long_df,
        x="Date",
        y="Crime Rate",
        hue="Transport Mode"
    )

    ax.set_title("Crime Rate Over Time by Transport Mode")
    ax.set_ylabel("Crime Rate (per million journeys)")

    plt.xticks(rotation=45)
    plt.tight_layout()

    plt.savefig("rates_over_time.png", dpi=300)

    plt.show()

    print("\nSaved as rates_over_time.png")


# -------------------------------------------------
# CRIME VOLUME PLOT
# -------------------------------------------------

def plot_volumes():

    modes = choose_modes(volume_modes)

    if modes is None:
        return

    long_df = df.melt(
        id_vars="Date",
        value_vars=[f"Crime_Volume_{m}" for m in modes],
        var_name="Transport Mode",
        value_name="Crime Volume"
    )

    long_df["Transport Mode"] = long_df["Transport Mode"].str.replace("Crime_Volume_", "")

    plt.figure(figsize=(14,7))

    ax = sns.scatterplot(
        data=long_df,
        x="Date",
        y="Crime Volume",
        hue="Transport Mode"
    )

    ax.set_title("Crime Volume Over Time by Transport Mode")
    ax.set_ylabel("Crime Volume")

    ax.yaxis.set_major_formatter(
        ticker.FuncFormatter(lambda x, pos: f"{x/1_000_000:.1f}M")
    )

    plt.xticks(rotation=45)
    plt.tight_layout()

    plt.savefig("volumes_over_time.png", dpi=300)

    plt.show()

    print("\nSaved as volumes_over_time.png")


# -------------------------------------------------
# CORRELATION COMPARISON
# -------------------------------------------------

def compare_two_modes():

    print("\nAvailable transport modes:\n")

    for m in rate_modes:
        print("-", m)

    mode1 = input("\nFirst mode: ").strip()
    mode2 = input("Second mode: ").strip()

    if mode1 not in rate_modes or mode2 not in rate_modes:
        print("\nInvalid mode names.")
        return

    col1 = f"Crime_Rate_{mode1}"
    col2 = f"Crime_Rate_{mode2}"

    merged = df[["Date", col1, col2]].dropna()

    if merged.empty:
        print("\nNo matching data for these modes.")
        return

    corr = merged[col1].corr(merged[col2])
    corr_value = round(corr,3)

    print(f"\nCorrelation: {corr_value}")

    plt.figure(figsize=(10,8))

    ax = sns.scatterplot(
        data=merged,
        x=col1,
        y=col2,
        s=90,
        alpha=0.85,
        edgecolor="black",
        linewidth=0.4
    )

    sns.regplot(
        data=merged,
        x=col1,
        y=col2,
        scatter=False,
        color="#d62728"
    )

    ax.set_title(f"{mode1} vs {mode2} Crime Rates\nCorrelation = {corr_value}")

    ax.set_xlabel(f"{mode1} Crime Rate")
    ax.set_ylabel(f"{mode2} Crime Rate")

    plt.grid(True, linestyle="--", alpha=0.4)

    plt.tight_layout()

    plt.savefig("transport_comparison.png", dpi=300)

    plt.show()

    print("\nSaved as transport_comparison.png")


# -------------------------------------------------
# MAIN LOOP
# -------------------------------------------------

while True:

    print_menu()

    choice = input("Enter your choice (1–4): ")

    if choice == "1":
        plot_rates()

    elif choice == "2":
        plot_volumes()

    elif choice == "3":
        compare_two_modes()

    elif choice == "4":
        print("\nExiting tool.")
        sys.exit()

    else:
        print("\nInvalid choice.")