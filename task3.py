import os
import numpy as np
import pandas as pd

# =========================================================
# LOAD EXCEL DATA
# =========================================================

BASE_DIR = r"C:\Users\pc\Desktop\PROJECTS\luhum"
excel_path = os.path.join(BASE_DIR, "Deep Earth Mining Data.xlsx")

xls = pd.ExcelFile(excel_path)
comp = pd.read_excel(xls, "Composition")
cost = pd.read_excel(xls, "Cost")
market = pd.read_excel(xls, "Market")
refining = pd.read_excel(xls, "Refining Costs")

YEAR_MAP = {5: 2030, 10: 2035, 15: 2040}

# =========================================================
# CLEAN INPUTS
# =========================================================
# Forward-fill Location to handle rows where Location is NaN
# (these rows belong to the previous location)

comp = comp.copy()
comp["Location"] = comp["Location"].ffill()
comp = comp[comp["Location"].str.startswith("Location", na=False)]
comp["Depth_km"] = pd.to_numeric(comp["Depth_km"], errors="coerce")
comp = comp.dropna(subset=["Depth_km"])

cost = cost.copy()
cost["Location"] = cost["Location"].ffill()
cost = cost[cost["Location"].str.startswith("Location", na=False)]
cost["Depth_km"] = pd.to_numeric(cost["Depth_km"], errors="coerce")
cost = cost.dropna(subset=["Depth_km"])

# =========================================================
# SELECT TOP 4 MINERALS BY DEMAND–SUPPLY GAP
# =========================================================

market["gap"] = market["Demand ('000 Tonnes)"] - market["Supply ('000 Tonnes)"]

top4 = (
    market.groupby("Mineral")["gap"]
          .mean()
          .sort_values(ascending=False)
          .head(4)
          .index.tolist()
)

# Map mineral names between sheets
name_map = {
    "Lithium": "Lithium",
    "Nickel (Million Tonnes)": "Nickel",
    "Cobalt": "Cobalt",
    "Graphite": "Graphite",
    "Manganese": "Manganese",
    "Copper (Million Tones)": "Copper",
    "RareEarth": "RareEarth",
    "Zinc": "Zinc",
    "Tin": "Tin",
    "Aluminum ('000 Mil tonnes)": "Aluminum",
    "Iron ('000 mil ton)": "Iron",
    "Lead": "Lead",
    "Silver (per Kg)": "Silver",
    "Gold (per Kg)": "Gold",
    "Platinum (per Kg)": "Platinum",
    "Phosphorus": "Phosphorus",
    "Potash": "Potash",
    "Silicon ('000 mil tons)": "Silicon",
    "Germanium": "Germanium",
    "Gallium": "Gallium",
    "Antimony": "Antimony",
    "Molybdenum": "Molybdenum",
    "Vanadium": "Vanadium",
    "Tungsten": "Tungsten",
    "Selenium": "Selenium",
    "Indium": "Indium",
    "Tellurium": "Tellurium",
    "Bismuth": "Bismuth",
    "Cadmium": "Cadmium",
    "Chromium": "Chromium"
}

top4_cols = [name_map[m] for m in top4]
rev_map = {v: k for k, v in name_map.items()}

# =========================================================
# LOOKUP DICTIONARIES
# =========================================================

price_dict = {
    (row["Mineral"], int(row["Year"])): row["Price_USD_per_ton"]
    for _, row in market.iterrows()
}

ref_cost_dict = dict(zip(refining["Unnamed: 0"], refining["Refining Cost (USD/Ton)"]))

def get_price(col, year):
    return price_dict[(rev_map[col], year)]

def get_ref_cost(col):
    return ref_cost_dict[col]

# =========================================================
# COST MODEL (matching test_profit.py exactly)
# =========================================================

cost["mining_cost_per_ton"] = (
    cost["Total Extraction Cost ('000 USD/ton)"] * 1000 +
    cost["Manpower Cost (USD/ton)"]
)

# =========================================================
# ORE TONNAGE ASSUMPTION
# =========================================================
# Required to convert composition % → actual metal mass.

ORE_TONNAGE = 100000   # 100,000 tons of ore (matching test_profit.py)

# =========================================================
# PROFIT CALCULATION FOR ONE (LOCATION, DEPTH, HORIZON)
# =========================================================

def calc_profit(location, depth, horizon):
    year = YEAR_MAP[horizon]

    c_row = comp[(comp["Location"] == location) & (comp["Depth_km"] == depth)]
    if c_row.empty:
        return -np.inf
    c_row = c_row.iloc[0]

    k_row = cost[(cost["Location"] == location) & (cost["Depth_km"] == depth)]
    if k_row.empty:
        return -np.inf
    k_row = k_row.iloc[0]

    mining_cost_per_ton_ore = k_row["mining_cost_per_ton"]
    total_profit = 0

    for col in top4_cols:
        pct = c_row[col]
        if pd.isna(pct) or pct <= 0:
            continue

        # STEP 1: Mass of metal from ore tonnage
        mass_metal = (pct / 100) * ORE_TONNAGE

        mineral = rev_map[col]
        mkt = market[(market["Mineral"] == mineral) & (market["Year"] == year)].iloc[0]

        # STEP 2: Demand-Supply Limit
        gap_tons = max(mkt["gap"] * 1000, 0)
        if gap_tons == 0:
            continue

        effective_mass = min(mass_metal, gap_tons)

        # STEP 3: Convert mining cost (per ton ore → per ton metal)
        mining_cost_per_ton_metal = mining_cost_per_ton_ore / (pct / 100)

        # STEP 4: Cost per ton of metal (mining + refining)
        refining_cost = get_ref_cost(col)
        total_cost_per_ton = mining_cost_per_ton_metal + refining_cost

        # STEP 5: Apply assignment formula
        price = get_price(col, year)
        profit_m = effective_mass * (price - total_cost_per_ton)

        total_profit += profit_m

    return total_profit

# =========================================================
# OPTIMIZE DEPTH FOR EACH LOCATION & HORIZON
# =========================================================

locations = sorted(comp["Location"].unique())
depths = sorted(comp["Depth_km"].unique())

rows = []

for loc in locations:
    for h in [5, 10, 15]:
        best_p = -np.inf
        best_d = None

        for d in depths:
            p = calc_profit(loc, d, h)
            if p > best_p:
                best_p = p
                best_d = d

        rows.append({
            "Horizon": f"{h} yrs ({YEAR_MAP[h]})",
            "Location": loc.replace("Location ", ""),
            "Optimal Depth": f"{int(best_d)} km",
            "Profit (B USD)": best_p / 1e9
        })

result = pd.DataFrame(rows)

result["H"] = result["Horizon"].map({

    "5 yrs (2030)": 0,
    "10 yrs (2035)": 1,
    "15 yrs (2040)": 2
})
result["L"] = result["Location"].map({"A": 0, "B": 1, "C": 2})

result = result.sort_values(["H", "L"]).drop(columns=["H", "L"])

# =========================================================
# OUTPUT
# =========================================================

print("\n==================== FINAL TASK 3 OUTPUT ====================\n")
print(result.to_string(index=False))
print("\n=============================================================\n")

result.to_csv(os.path.join(BASE_DIR, "task3_output.csv"), index=False)
print("Saved to task3_output.csv\n")

