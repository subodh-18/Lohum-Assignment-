import pandas as pd
import numpy as np

xls = pd.ExcelFile("Deep Earth Mining Data.xlsx")
comp = pd.read_excel(xls, "Composition")
cost = pd.read_excel(xls, "Cost")
market = pd.read_excel(xls, "Market")
refining = pd.read_excel(xls, "Refining Costs")

# Clean data
comp = comp.dropna(subset=["Location", "Depth_km"])
comp["Depth_km"] = pd.to_numeric(comp["Depth_km"], errors="coerce")
comp = comp.dropna(subset=["Depth_km"])
comp = comp[comp["Location"].str.startswith("Location", na=False)]

cost = cost.dropna(subset=["Location", "Depth_km"])
cost["Depth_km"] = pd.to_numeric(cost["Depth_km"], errors="coerce")
cost = cost.dropna(subset=["Depth_km"])
cost = cost[cost["Location"].str.startswith("Location", na=False)]
cost["mining_cost"] = cost["Total Extraction Cost ('000 USD/ton)"] * 1000 + cost["Manpower Cost (USD/ton)"]

market["gap"] = market["Demand ('000 Tonnes)"] - market["Supply ('000 Tonnes)"]

# Test case: Location A, Depth 0
rcomp = comp[(comp["Location"] == "Location A") & (comp["Depth_km"] == 0.0)].iloc[0]
rcost = cost[(cost["Location"] == "Location A") & (cost["Depth_km"] == 0.0)].iloc[0]

ORE_TONNAGE = 100000
pct = rcomp["Lithium"]
mining_cost = rcost["mining_cost"]

print("="*60)
print("TEST CASE: Location A, Depth 0, Lithium")
print("="*60)
print(f"Lithium composition: {pct}%")
print(f"Mining cost per ton: ${mining_cost:,.0f}")
print(f"ORE_TONNAGE: {ORE_TONNAGE:,} tons")
print()

max_metal = (pct / 100) * ORE_TONNAGE
print(f"Max Lithium metal from ore: {max_metal:,.2f} tons")

# Check market for 2030
mkt = market[(market["Mineral"] == "Lithium") & (market["Year"] == 2030)].iloc[0]
gap_tons = mkt["gap"] * 1000
price = mkt["Price_USD_per_ton"]

print(f"Market gap (2030): {gap_tons:,.2f} tons")
print(f"Price (2030): ${price:,.0f}/ton")

# Get refining cost
ref = refining[refining["Unnamed: 0"] == "Lithium"].iloc[0]
ref_cost = ref["Refining Cost (USD/Ton)"]
print(f"Refining cost: ${ref_cost:,.0f}/ton")
print()

# Calculate profit
if max_metal > gap_tons:
    metal_produced = gap_tons
    ore_needed = metal_produced / (pct / 100)
else:
    metal_produced = max_metal
    ore_needed = ORE_TONNAGE

print(f"Metal produced: {metal_produced:,.2f} tons")
print(f"Ore needed: {ore_needed:,.2f} tons")
print()

revenue = metal_produced * price
cost_mining = ore_needed * mining_cost
cost_refining = metal_produced * ref_cost

print(f"Revenue: ${revenue:,.0f}")
print(f"Mining cost: ${cost_mining:,.0f}")
print(f"Refining cost: ${cost_refining:,.0f}")
print()

profit = revenue - cost_mining - cost_refining
print(f"Profit: ${profit:,.0f}")
print("="*60)

