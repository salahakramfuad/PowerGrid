# -*- coding: utf-8 -*-
"""
Created on Mon Feb 23 23:50:33 2026

@author: USER
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d
import os

# ========== CONFIG ==========
FILES = {"Max": "pssemaxval.xlsx", "Min": "psseminval.xlsx"}
STATION = "Panchagarh"
BASE_KV = 132.0
os.makedirs("panchagarh_analysis", exist_ok=True)

def calculate_improvement(v_base, v_facts):
    """Calculates % reduction in deviation from 1.0 p.u."""
    dev_base = abs(v_base - 1.0)
    dev_facts = abs(v_facts - 1.0)
    if dev_base < 0.0001: return 0  # Avoid div by zero if already perfect
    return ((dev_base - dev_facts) / dev_base) * 100

# ========== 1. DATA EXTRACTION ==========
monthly_improvement = {"Max": {"Daytime": [], "Nighttime": []}, 
                       "Min": {"Daytime": [], "Nighttime": []}}

for label, file in FILES.items():
    if not os.path.exists(file):
        print(f"❌ File not found: {file}")
        continue
        
    df = pd.read_excel(file, header=None)
    # Find the row where Panchagarh is mentioned
    idx = df[df[0] == STATION].index[0]
    # The 132kV data is in the row immediately below the station name
    row_132 = df.iloc[idx + 1] 
    
    # We expect 12 months of data (This assumes your Excel has months stacked vertically)
    # If your Excel format has months repeating, we parse all occurrences:
    indices = df[df[0] == STATION].index
    for i in indices:
        r = df.iloc[i + 1]
        # Col 1: Solar, Col 2: Solar+FACTS, Col 3: Normal, Col 4: Normal+FACTS
        monthly_improvement[label]["Daytime"].append(calculate_improvement(r[1]/BASE_KV, r[2]/BASE_KV))
        monthly_improvement[label]["Nighttime"].append(calculate_improvement(r[3]/BASE_KV, r[4]/BASE_KV))

# ========== 2. 365-DAY INTERPOLATION ==========
days_in_year = np.linspace(1, 12, 365)
months_x = np.arange(1, 13)
daily_results = []

for val_type in ["Max", "Min"]:
    for time_of_day in ["Daytime", "Nighttime"]:
        y_points = monthly_improvement[val_type][time_of_day]
        
        # Cubic Spline creates a smooth curve through the 12 months
        f = interp1d(months_x, y_points, kind='cubic', fill_value="extrapolate")
        y_365 = f(days_in_year)
        
        for d, val in enumerate(y_365):
            daily_results.append({
                "Day": d + 1,
                "Type": val_type,
                "Time": time_of_day,
                "Improvement_Pct": round(val, 4)
            })

# Save to Excel
df_365 = pd.DataFrame(daily_results)
# Pivot table to make it easier to use (Days as rows, Scenarios as columns)
df_pivot = df_365.pivot(index='Day', columns=['Type', 'Time'], values='Improvement_Pct')
df_pivot.to_excel("panchagarh_analysis/Panchagarh_365_Improvement.xlsx")

# ========== 3. PLOTTING THE IMPROVEMENT (MAX vs MIN) ==========
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10), sharex=True)

# Graph 1: Max Voltage Improvement
ax1.plot(days_in_year, df_pivot[('Max', 'Daytime')], label='Daytime (Solar) Improvement', color='#f39c12', linewidth=2)
ax1.plot(days_in_year, df_pivot[('Max', 'Nighttime')], label='Nighttime (Normal) Improvement', color='#d35400', linestyle='--')
ax1.scatter(months_x, monthly_improvement["Max"]["Daytime"], color='black', label='Simulated Points (Max)', zorder=5)
ax1.set_title(f"Panchagarh 132kV: FACTS Improvement for MAX Voltage", fontsize=13, fontweight='bold')
ax1.set_ylabel("% Improvement")
ax1.grid(True, alpha=0.3)
ax1.legend(loc='upper right')

# Graph 2: Min Voltage Improvement
ax2.plot(days_in_year, df_pivot[('Min', 'Daytime')], label='Daytime (Solar) Improvement', color='#3498db', linewidth=2)
ax2.plot(days_in_year, df_pivot[('Min', 'Nighttime')], label='Nighttime (Normal) Improvement', color='#2980b9', linestyle='--')
ax2.scatter(months_x, monthly_improvement["Min"]["Daytime"], color='black', label='Simulated Points (Min)', zorder=5)
ax2.set_title(f"Panchagarh 132kV: FACTS Improvement for MIN Voltage", fontsize=13, fontweight='bold')
ax2.set_ylabel("% Improvement")
ax2.set_xlabel("Month (1=Jan, 12=Dec)")
ax2.set_xticks(range(1, 13))
ax2.set_xticklabels(["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])
ax2.grid(True, alpha=0.3)
ax2.legend(loc='upper right')

plt.tight_layout()
plt.savefig("panchagarh_analysis/Panchagarh_Max_Min_Improvement.png", dpi=300)
plt.show()
print("✅ Task Complete! Created 365-day trend for Panchagarh 132kV.")