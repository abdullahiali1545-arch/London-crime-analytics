# ============================================================
# PART 3 – STEP 3
# Generate clean choropleth labelled with offence, measure,
# and month/year from London_crime_data.xlsx
# ============================================================

import pandas as pd
import geopandas as gpd
import matplotlib.pyplot as plt


# ------------------------------------------------------------
# STEP 1 – LOAD SHAPEFILE
# ------------------------------------------------------------
map_df = gpd.read_file("London_Borough_Excluding_MHW.shp")


# ------------------------------------------------------------
# STEP 2 – LOAD FILTERED DATA FROM STEP 2
# ------------------------------------------------------------
crime_df = pd.read_excel("London_crime_data.xlsx")


# ------------------------------------------------------------
# STEP 3 – EXTRACT INFORMATION FOR TITLE
# ------------------------------------------------------------
offence_type = crime_df["Offence Group"].iloc[0]
month_year   = crime_df["Month_fmt"].iloc[0]
measure_column = crime_df.columns[-1]


# ------------------------------------------------------------
# STEP 4 – RENAME COLUMN TO MATCH SHAPEFILE
# ------------------------------------------------------------
crime_df = crime_df.rename(columns={"Area name": "NAME"})

crime_df["NAME"] = crime_df["NAME"].astype(str).str.strip()
map_df["NAME"]   = map_df["NAME"].astype(str).str.strip()


# ------------------------------------------------------------
# STEP 5 – MERGE DATA WITH MAP
# ------------------------------------------------------------
merged_df = map_df.merge(crime_df, on="NAME", how="left")


# ------------------------------------------------------------
# STEP 6 – CREATE CLEAN CHOROPLETH
# ------------------------------------------------------------
fig, ax = plt.subplots(figsize=(10, 10))

merged_df.plot(
    column=measure_column,
    cmap="Reds",
    linewidth=0.7,
    edgecolor="white",
    legend=True,
    legend_kwds={
        "label": f"{measure_column} ({offence_type})",
        "orientation": "vertical",
        "shrink": 0.6
    },
    ax=ax
)


# ------------------------------------------------------------
# STEP 7 – OPTIONAL BOROUGH LABELS (small and clean)
# ------------------------------------------------------------
for idx, row in merged_df.iterrows():

    if row.geometry is None:
        continue

    centroid = row.geometry.centroid

    ax.text(
        centroid.x,
        centroid.y,
        row["NAME"],
        fontsize=6,
        ha="center",
        va="center"
    )


# ------------------------------------------------------------
# STEP 8 – TITLE WITH OFFENCE AND DATE
# ------------------------------------------------------------
ax.set_title(
    f"London Crime Choropleth\n"
    f"{offence_type} – {measure_column} ({month_year})",
    fontsize=16,
    weight="bold",
    pad=15
)

ax.axis("off")

fig.patch.set_facecolor("white")


# ------------------------------------------------------------
# STEP 9 – SAVE IMAGE
# ------------------------------------------------------------
output_file = f"London_crime_{offence_type.replace(' ','_')}_{month_year}.png"

plt.savefig(
    output_file,
    dpi=300,
    bbox_inches="tight"
)

plt.show()


# ------------------------------------------------------------
# CONFIRM
# ------------------------------------------------------------
print("\nChoropleth created successfully.")
print(f"Offence: {offence_type}")
print(f"Month: {month_year}")
print(f"Measure: {measure_column}")
print(f"Saved as: {output_file}")