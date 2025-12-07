import os
import numpy as np
import pandas as pd
from itertools import combinations

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
# GET OPTIMAL DEPTHS FROM TASK 3
# =========================================================

# Read task3 output to get optimal depths for Location A
try:
    task3_output = pd.read_csv(os.path.join(BASE_DIR, "task3_output.csv"))
    optimal_depths = {}
    for _, row in task3_output.iterrows():
        if row["Location"] == "A":
            horizon = row["Horizon"]
            depth_str = row["Optimal Depth"].replace(" km", "")
            optimal_depths[horizon] = float(depth_str)
    print("Optimal depths from Task 3:")
    for h, d in optimal_depths.items():
        print(f"  {h}: {d} km")
except:
    # Fallback: use depth 0 for all horizons
    optimal_depths = {
        "5 yrs (2030)": 0.0,
        "10 yrs (2035)": 0.0,
        "15 yrs (2040)": 0.0
    }
    print("Using default depth 0 for all horizons")

# =========================================================
# PREPARE DATA FOR LOCATION A
# =========================================================

LOCATION = "Location A"

# Get all available minerals with positive composition at optimal depths
market["gap"] = market["Demand ('000 Tonnes)"] - market["Supply ('000 Tonnes)"]

# Map mineral names
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

rev_map = {v: k for k, v in name_map.items()}

# Lookup dictionaries
price_dict = {
    (row["Mineral"], int(row["Year"])): row["Price_USD_per_ton"]
    for _, row in market.iterrows()
}

ref_cost_dict = dict(zip(refining["Unnamed: 0"], refining["Refining Cost (USD/Ton)"]))

def get_price(col, year):
    return price_dict[(rev_map[col], year)]

def get_ref_cost(col):
    return ref_cost_dict[col]

# Mining cost calculation
cost["mining_cost_per_ton"] = (
    cost["Total Extraction Cost ('000 USD/ton)"] * 1000 +
    cost["Manpower Cost (USD/ton)"]
)

# =========================================================
# GET AVAILABLE MINERALS AT OPTIMAL DEPTHS
# =========================================================

available_minerals = {}  # {horizon: [list of mineral columns]}
for horizon in [5, 10, 15]:
    year = YEAR_MAP[horizon]
    depth = optimal_depths[f"{horizon} yrs ({year})"]
    
    comp_row = comp[(comp["Location"] == LOCATION) & (comp["Depth_km"] == depth)]
    if comp_row.empty:
        available_minerals[horizon] = []
        continue
    
    comp_row = comp_row.iloc[0]
    minerals = []
    
    # Get all minerals with positive composition and positive gap
    for col_name, market_name in rev_map.items():
        if col_name not in comp_row.index:
            continue
        pct = comp_row[col_name]
        if pd.isna(pct) or pct <= 0:
            continue
        
        # Check if there's a demand gap
        mkt = market[(market["Mineral"] == market_name) & (market["Year"] == year)]
        if not mkt.empty:
            gap = mkt.iloc[0]["gap"]
            if gap > 0:
                minerals.append(col_name)
    
    available_minerals[horizon] = minerals
    print(f"\nHorizon {horizon} ({year}): {len(minerals)} available minerals")

# =========================================================
# GET ADDITIONAL LOGISTICS COST
# =========================================================

# Logistics cost increases with number of minerals refined
# Create mapping: number_of_minerals -> logistics_cost_per_ton_ore
cost_A = cost[cost["Location"] == LOCATION].copy()
logistics_map = {}

for _, row in cost_A.iterrows():
    num_minerals = row.get("Number of minerals")
    add_cost = row.get("Additional Cost ")
    
    if pd.notna(num_minerals) and pd.notna(add_cost):
        try:
            # Skip header row (contains text)
            if isinstance(add_cost, str):
                continue
            # Convert from '000 USD/ton to USD/ton
            logistics_map[int(num_minerals)] = float(add_cost) * 1000
        except:
            continue

print(f"\nLogistics Cost Mapping (per ton of ore):")
for k in sorted(logistics_map.keys())[:10]:
    print(f"  {k} minerals: ${logistics_map[k]:,.0f}/ton")

def get_logistics_cost(num_minerals):
    """Get logistics cost per ton of ore for given number of minerals."""
    if num_minerals <= 0:
        return 0.0
    # Use the highest available cost if num_minerals exceeds table
    keys = sorted(logistics_map.keys())
    if num_minerals > keys[-1]:
        return logistics_map[keys[-1]]
    return logistics_map.get(num_minerals, 0.0)

# =========================================================
# OPTIMIZATION: SELECT MINERALS AND ORE QUANTITY
# =========================================================

def calculate_profit_for_minerals(selected_minerals, ore_tonnage, horizon, depth):
    """
    Calculate profit for a set of selected minerals and ore tonnage.
    
    Formula: Profit = Σ mass × (price - (mining_cost + refining_cost)) - logistics_cost
    """
    year = YEAR_MAP[horizon]
    
    comp_row = comp[(comp["Location"] == LOCATION) & (comp["Depth_km"] == depth)].iloc[0]
    cost_row = cost[(cost["Location"] == LOCATION) & (cost["Depth_km"] == depth)].iloc[0]
    
    mining_cost_per_ton_ore = cost_row["mining_cost_per_ton"]
    total_profit = 0
    
    # Logistics cost: applies to ore tonnage, increases with number of minerals
    num_minerals = len(selected_minerals)
    logistics_cost_per_ton = get_logistics_cost(num_minerals)
    logistics_total = logistics_cost_per_ton * ore_tonnage
    
    for col in selected_minerals:
        pct = comp_row[col]
        if pd.isna(pct) or pct <= 0:
            continue
        
        # Mass of metal from ore
        mass_metal = (pct / 100) * ore_tonnage
        
        # Get market data
        mineral = rev_map[col]
        mkt = market[(market["Mineral"] == mineral) & (market["Year"] == year)].iloc[0]
        
        gap_tons = max(mkt["gap"] * 1000, 0)
        if gap_tons == 0:
            continue
        
        # Production limited by gap
        effective_mass = min(mass_metal, gap_tons)
        
        # Convert mining cost to per ton of metal
        mining_cost_per_ton_metal = mining_cost_per_ton_ore / (pct / 100)
        
        # Get costs
        refining_cost = get_ref_cost(col)
        price = get_price(col, year)
        
        # Cost per ton of metal
        total_cost_per_ton = mining_cost_per_ton_metal + refining_cost
        
        # Profit per mineral
        profit_m = effective_mass * (price - total_cost_per_ton)
        total_profit += profit_m
    
    # Subtract logistics cost
    total_profit -= logistics_total
    
    return total_profit

# =========================================================
# OPTIMIZE FOR EACH HORIZON
# =========================================================

def rank_minerals_by_margin(minerals, horizon, depth):
    """Rank minerals by profit margin (price - cost per ton of metal)."""
    year = YEAR_MAP[horizon]
    comp_row = comp[(comp["Location"] == LOCATION) & (comp["Depth_km"] == depth)].iloc[0]
    cost_row = cost[(cost["Location"] == LOCATION) & (cost["Depth_km"] == depth)].iloc[0]
    
    mining_cost_per_ton_ore = cost_row["mining_cost_per_ton"]
    mineral_margins = []
    
    for col in minerals:
        pct = comp_row[col]
        if pd.isna(pct) or pct <= 0:
            continue
        
        mineral = rev_map[col]
        mkt = market[(market["Mineral"] == mineral) & (market["Year"] == year)].iloc[0]
        
        gap_tons = max(mkt["gap"] * 1000, 0)
        if gap_tons == 0:
            continue
        
        price = get_price(col, year)
        refining_cost = get_ref_cost(col)
        mining_cost_per_ton_metal = mining_cost_per_ton_ore / (pct / 100)
        
        # Profit margin per ton of metal
        margin = price - (mining_cost_per_ton_metal + refining_cost)
        
        mineral_margins.append({
            "mineral": col,
            "margin": margin,
            "gap_tons": gap_tons,
            "composition_pct": pct
        })
    
    # Sort by margin (highest first)
    mineral_margins.sort(key=lambda x: x["margin"], reverse=True)
    return mineral_margins

results = []

for horizon in [5, 10, 15]:
    year = YEAR_MAP[horizon]
    depth = optimal_depths[f"{horizon} yrs ({year})"]
    minerals = available_minerals[horizon]
    
    if not minerals:
        continue
    
    print(f"\n{'='*60}")
    print(f"Optimizing for {horizon}-year horizon ({year}), Depth {depth} km")
    print(f"{'='*60}")
    
    # Rank minerals by profit margin
    ranked_minerals = rank_minerals_by_margin(minerals, horizon, depth)
    print(f"\nTop 5 minerals by profit margin:")
    for i, m in enumerate(ranked_minerals[:5]):
        print(f"  {i+1}. {rev_map[m['mineral']]}: ${m['margin']:,.0f}/ton margin")
    
    best_profit = -np.inf
    best_minerals = []
    best_ore_tonnage = 0
    
    # Try different numbers of top minerals (1 to min(10, available))
    max_minerals_to_try = min(len(ranked_minerals), 10)
    
    for num_minerals in range(1, max_minerals_to_try + 1):
        # Use top N minerals by margin
        selected = [m["mineral"] for m in ranked_minerals[:num_minerals]]
        
        # Try different ore tonnages
        # Use finer grid for optimization
        ore_grid = list(range(50000, 1000001, 50000))  # 50k to 1M in 50k steps
        
        for ore_tonnage in ore_grid:
            profit = calculate_profit_for_minerals(
                selected, ore_tonnage, horizon, depth
            )
            
            if profit > best_profit:
                best_profit = profit
                best_minerals = selected.copy()
                best_ore_tonnage = ore_tonnage
    
    print(f"\nBest solution:")
    print(f"  Minerals selected: {len(best_minerals)}")
    print(f"  Minerals: {[rev_map[m] for m in best_minerals]}")
    print(f"  Ore tonnage: {best_ore_tonnage:,} tons")
    print(f"  Total profit: ${best_profit/1e9:.3f} B USD")
    
    results.append({
        "Horizon": f"{horizon} yrs ({year})",
        "Optimal Depth": f"{int(depth)} km",
        "Number of Minerals": len(best_minerals),
        "Minerals Selected": ", ".join([rev_map[m] for m in best_minerals]),
        "Ore Tonnage (tons)": best_ore_tonnage,
        "Profit (B USD)": best_profit / 1e9
    })

# =========================================================
# OUTPUT RESULTS
# =========================================================

result_df = pd.DataFrame(results)

print("\n" + "="*60)
print("FINAL TASK 4 OUTPUT")
print("="*60)
print(result_df.to_string(index=False))
print("="*60)

# Save to CSV
output_path = os.path.join(BASE_DIR, "task4_output.csv")
result_df.to_csv(output_path, index=False)
print(f"\nOutput saved successfully to:")
print(f"  {output_path}")
print(f"  File exists: {os.path.exists(output_path)}")
print(f"  File size: {os.path.getsize(output_path)} bytes")
