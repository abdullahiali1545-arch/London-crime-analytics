import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import numbers
import seaborn as sns
import matplotlib.pyplot as plt

file = "public-transport-crime-london.xlsx"
source_sheet = "Volume and Rates"
target_sheet = "Clean data"

def extract_and_transform(file_path, source_sheet, target_sheet, start_row, end_row):

    df = pd.read_excel(file_path, sheet_name=source_sheet, header=None)
    selected = df.iloc[start_row-1:end_row].reset_index(drop=True)
    selected = selected.replace(',', '', regex=True)

    result = []
    i = 0
    current_year = None

    month_map = {
        "Apr": 4, "May": 5, "Jun": 6, "Jul": 7,
        "Aug": 8, "Sept": 9, "Oct": 10, "Nov": 11,
        "Dec": 12, "Jan": 1, "Feb": 2, "Mar": 3
    }

    month_order = [
        "Apr", "May", "Jun", "Jul", "Aug", "Sept",
        "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"
    ]

    while i < len(selected):

        if selected.iloc[i].isna().all():
            i += 1
            continue

        first_cell = str(selected.iloc[i, 0]).strip()

        if "Network-wide" in first_cell:

            fy_text = first_cell.split("FY")[-1]
            start_year = int(fy_text.split("/")[0])
            current_year = start_year

            i += 2
            block_rows = []

            while i < len(selected):

                if selected.iloc[i].isna().all():
                    i += 1
                    continue

                if "Network-wide" in str(selected.iloc[i, 0]):
                    break

                block_rows.append(selected.iloc[i])
                i += 1

            block_df = pd.DataFrame(block_rows).reset_index(drop=True)

            for row_index in range(len(block_df)):

                transport_mode = str(block_df.iloc[row_index, 0]).strip()
                col_index = 1

                for month in month_order:

                    volume = pd.to_numeric(block_df.iloc[row_index, col_index], errors="coerce")
                    rate = pd.to_numeric(block_df.iloc[row_index, col_index + 1], errors="coerce")

                    month_number = month_map[month]
                    year = current_year if month_number >= 4 else current_year + 1
                    date_string = f"{year}-{month_number:02d}"

                    result.append([
                        transport_mode,
                        date_string,
                        volume,
                        rate
                    ])

                    col_index += 2
        else:
            i += 1

    result_df = pd.DataFrame(result, columns=[
        "Transport Mode",
        "Date",
        "Crime_Volume",
        "Crime_Rate"
    ])

    # -------------------------
    # CLEANING & MERGING
    # -------------------------

    result_df.loc[result_df["Transport Mode"].eq("Trams"), "Transport Mode"] = "London Tramlink"

    # Merge LU + DLR
    mask_lu = result_df["Transport Mode"].eq("London Underground")
    mask_dlr = result_df["Transport Mode"].eq("Docklands Light Railway")
    result_df.loc[mask_lu | mask_dlr, "Transport Mode"] = "London Underground / Docklands Light Railway"

    result_df = (
        result_df.groupby(["Transport Mode", "Date"], as_index=False)
        .apply(lambda g: pd.Series({
            "Crime_Volume": g["Crime_Volume"].sum(),
            "Crime_Rate": (g["Crime_Volume"] * g["Crime_Rate"]).sum() / g["Crime_Volume"].sum()
            if g["Crime_Volume"].sum() != 0 else np.nan
        }))
        .reset_index(drop=True)
    )

    # Merge TfL Rail + Elizabeth Line
    mask_tfl = result_df["Transport Mode"].eq("TfL Rail")
    mask_eliz = result_df["Transport Mode"].eq("Elizabeth Line (formerly TfL Rail)")
    result_df.loc[mask_tfl | mask_eliz, "Transport Mode"] = "Elizabeth Line"

    result_df = (
        result_df.groupby(["Transport Mode", "Date"], as_index=False)
        .apply(lambda g: pd.Series({
            "Crime_Volume": g["Crime_Volume"].sum(),
            "Crime_Rate": (g["Crime_Volume"] * g["Crime_Rate"]).sum() / g["Crime_Volume"].sum()
            if g["Crime_Volume"].sum() != 0 else np.nan
        }))
        .reset_index(drop=True)
    )

    # Remove old totals
    result_df = result_df[
        ~result_df["Transport Mode"].str.contains("all transport modes", case=False, na=False)
    ]

    # Add weighted totals
    totals = (
        result_df.groupby("Date")
        .apply(lambda x: pd.Series({
            "Transport Mode": "All transport modes",
            "Crime_Volume": x["Crime_Volume"].sum(),
            "Crime_Rate": (x["Crime_Volume"] * x["Crime_Rate"]).sum() / x["Crime_Volume"].sum()
            if x["Crime_Volume"].sum() != 0 else np.nan
        }))
        .reset_index()
    )

    result_df = pd.concat([result_df, totals], ignore_index=True)

    # Keep date as YYYY-MM string (NO timestamp)
    result_df["Date"] = pd.to_datetime(result_df["Date"]).dt.strftime("%Y-%m")

    result_df["Crime_Volume"] = result_df["Crime_Volume"].astype(float).round(1)
    result_df["Crime_Rate"] = result_df["Crime_Rate"].astype(float).round(1)

    # Remove Overground early period
    result_df = result_df[
        ~(
            (result_df["Transport Mode"] == "London Overground") &
            (result_df["Date"] >= "2009-04") &
            (result_df["Date"] <= "2011-03")
        )
    ]

    # -------------------------
    # SAVE WIDE FORMAT
    # -------------------------

    wide = result_df.pivot(
        index="Date",
        columns="Transport Mode",
        values=["Crime_Volume", "Crime_Rate"]
    )

    wide.columns = [
        f"{metric}_{mode.replace(' ', '_').replace('/', '_')}"
        for metric, mode in wide.columns
    ]

    wide = wide.sort_index()

    with pd.ExcelWriter(
        file_path,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:
        wide.to_excel(writer, sheet_name=target_sheet, index=True)

    print("Wide-format dataset created successfully.")

    return result_df


# Run transformation
long_df = extract_and_transform(file, source_sheet, target_sheet, 1, 400)

# Convert Date to datetime ONLY for plotting
long_df["Date"] = pd.to_datetime(long_df["Date"])

# ============================
# SCATTERPLOTS
# ============================

# Crime Volume
plt.figure(figsize=(12,6))
sns.scatterplot(
    data=long_df,
    x="Date",
    y="Crime_Volume",
    hue="Transport Mode"
)
plt.xticks(rotation=45)
plt.title("Crime Volume Over Time by Transport Mode")
plt.tight_layout()
plt.show()

# Crime Rate
plt.figure(figsize=(12,6))
sns.scatterplot(
    data=long_df,
    x="Date",
    y="Crime_Rate",
    hue="Transport Mode"
)
plt.xticks(rotation=45)
plt.title("Crime Rate Over Time by Transport Mode")
plt.tight_layout()
plt.show()