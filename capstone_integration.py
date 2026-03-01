# -*- coding: utf-8 -*-
"""
EEE Capstone: 132kV Northern Bangladesh — ML Forecast + PSS/E FACTS Integration
5-stage pipeline: ML forecast (2026-2027) → Validation (2025 MAPE) → PSS/E lookup → Mitigation math → Dual-axis plots
"""

import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.metrics import mean_absolute_percentage_error

try:
    from xgboost import XGBRegressor
    Regressor = XGBRegressor
    regressor_kwargs = dict(n_estimators=500, learning_rate=0.05, max_depth=6)
except Exception:
    from sklearn.ensemble import GradientBoostingRegressor
    Regressor = GradientBoostingRegressor
    regressor_kwargs = dict(n_estimators=500, learning_rate=0.05, max_depth=6)

BASE_KV = 132.0
PGCB_LOWER_PU = 0.94   # -6%
PGCB_UPPER_PU = 1.06   # +6%
KV_PLOT_MIN, KV_PLOT_MAX = 100.0, 145.0
OUTPUT_DIR = "capstone_outputs"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------------------------------------------------------------------
# Stage 1: ML forecasting — load data, train, 2-year forecast (2026-2027)
# ---------------------------------------------------------------------------
# Option A: from pre-processed CSV (run dataprocess.py first)
df = pd.read_csv('processed_data_step1.csv')
df['Date'] = pd.to_datetime(df['Date'])

targets = ['Load (MW)', 'Max_Voltage (kV)', 'Min_Voltage (kV)']
for col in targets:
    df[col] = pd.to_numeric(df[col], errors='coerce')
df = df.dropna(subset=targets)

features = ['Year', 'Month', 'DayOfYear', 'DayOfWeek', 'Is_Irrigation', 'Is_Weekend']
train_df = df[df['Year'] <= 2024]
val_df = df[df['Year'] == 2025]

models = {}
for target in targets:
    m = Regressor(**regressor_kwargs)
    m.fit(train_df[features], train_df[target])
    models[target] = m

# 2-year forecast 2026-2027
future_dates = pd.date_range(start='2026-01-01', end='2027-12-31', freq='D')
future_df = pd.DataFrame({'Date': future_dates})
future_df['Year'] = future_df['Date'].dt.year
future_df['Month'] = future_df['Date'].dt.month
future_df['DayOfYear'] = future_df['Date'].dt.dayofyear
future_df['DayOfWeek'] = future_df['Date'].dt.dayofweek
future_df['Is_Irrigation'] = future_df['Month'].apply(lambda x: 1 if x in [2, 3, 4, 5] else 0)
future_df['Is_Weekend'] = future_df['DayOfWeek'].apply(lambda x: 1 if x in [4, 5] else 0)

for target in targets:
    future_df[f'Predicted_{target}'] = models[target].predict(future_df[features])

# Stage 1 plot: 2-year forecast (2026-2027) for Load, Max V, Min V
fig1, axes = plt.subplots(3, 1, figsize=(12, 10), sharex=True)
for i, target in enumerate(targets):
    axes[i].plot(future_df['Date'], future_df[f'Predicted_{target}'], color='green', label='Forecast')
    axes[i].set_ylabel(target)
    axes[i].legend()
    axes[i].grid(True, alpha=0.3)
axes[0].set_title('Stage 1: 2-year ML forecast (2026-2027) — Load & Max/Min Voltage')
axes[-1].set_xlabel('Date')
plt.tight_layout()
plt.savefig(os.path.join(OUTPUT_DIR, 'stage1_forecast_2026_2027.png'), dpi=300, bbox_inches='tight')
plt.close()
print("Stage 1 done: forecast 2026-2027 saved.")

# ---------------------------------------------------------------------------
# Stage 2: Validation — MAPE for 2025 (predicted vs actual)
# ---------------------------------------------------------------------------
print("\n--- Stage 2: Validation (2025) ---")
for target in targets:
    pred_2025 = models[target].predict(val_df[features])
    mape = mean_absolute_percentage_error(val_df[target], pred_2025) * 100
    print(f"  {target}: MAPE = {mape:.2f}%")

# ---------------------------------------------------------------------------
# Stage 3: PSS/E scenario lookup — 4 conditions
# ---------------------------------------------------------------------------
# Conditions: (1) Normal, (2) Normal+FACTS, (3) Solar, (4) Solar+FACTS
# We use precomputed 365-day improvement (from mlprep.py) for Panchagarh 132kV.
# Keys: DayOfYear (1-365) -> improvement % for Max and Min (average Daytime/Nighttime).
PSSE_IMPROVEMENT_PATH = "panchagarh_analysis/Panchagarh_365_Improvement.xlsx"
if not os.path.exists(PSSE_IMPROVEMENT_PATH):
    raise FileNotFoundError(f"Run mlprep.py first to create {PSSE_IMPROVEMENT_PATH}")

df_improve = pd.read_excel(PSSE_IMPROVEMENT_PATH, header=[0, 1])
# Expected: first column ('Type','Time') has Day index; columns ('Max','Daytime'), ('Max','Nighttime'), ('Min','Daytime'), ('Min','Nighttime')
first_col = df_improve.columns[0]
df_improve = df_improve.set_index(first_col)
if isinstance(df_improve.columns, pd.MultiIndex):
    df_improve['Improve_Max'] = (df_improve[('Max', 'Daytime')] + df_improve[('Max', 'Nighttime')]) / 2
    df_improve['Improve_Min'] = (df_improve[('Min', 'Daytime')] + df_improve[('Min', 'Nighttime')]) / 2
else:
    improve_cols = [c for c in df_improve.columns if 'Max' in str(c) or 'Min' in str(c)]
    df_improve['Improve_Max'] = df_improve.get('Improve_Max', df_improve.iloc[:, 0] if len(df_improve.columns) > 0 else 0)
    df_improve['Improve_Min'] = df_improve.get('Improve_Min', df_improve.iloc[:, 1] if len(df_improve.columns) > 1 else 0)
# Index may be 'Day' (string) or numeric; ensure 1..365
df_improve.index = pd.to_numeric(df_improve.index, errors='coerce')
df_improve = df_improve.dropna(how='all').loc[1:365]

# Lookup: day_of_year (1-365) -> (improve_max_pct, improve_min_pct)
day_to_improve_max = df_improve['Improve_Max'].to_dict()
day_to_improve_min = df_improve['Improve_Min'].to_dict()
# Reindex so Day 1..365 exist
improve_max = np.array([day_to_improve_max.get(d, 0) for d in range(1, 366)])
improve_min = np.array([day_to_improve_min.get(d, 0) for d in range(1, 366)])
print("Stage 3 done: PSS/E 4-condition lookup (Normal, Normal+FACTS, Solar, Solar+FACTS) applied via 365-day improvement.")

# ---------------------------------------------------------------------------
# Stage 4: Integration math — apply PSS/E % improvement to raw ML predictions
# ---------------------------------------------------------------------------
# Mitigated: raw voltage in p.u. -> reduce deviation from 1.0 by improvement_pct
def raw_to_mitigated_kv(V_kV, improvement_pct):
    raw_pu = V_kV / BASE_KV
    dev = raw_pu - 1.0
    improved_dev = dev * (1.0 - improvement_pct / 100.0)
    mitigated_pu = 1.0 + improved_dev
    return mitigated_pu * BASE_KV

future_df = future_df.copy()
# DayOfYear 1-365; for 2026-2027 we have 365+365 days, repeat improvement for second year
doy = future_df['DayOfYear'].values
imax = np.where(doy <= 365, improve_max[doy - 1], improve_max[(doy - 1) % 365])
imin = np.where(doy <= 365, improve_min[doy - 1], improve_min[(doy - 1) % 365])

future_df['Predicted_Max_Voltage_kV_Raw'] = future_df['Predicted_Max_Voltage (kV)']
future_df['Predicted_Min_Voltage_kV_Raw'] = future_df['Predicted_Min_Voltage (kV)']
future_df['Predicted_Max_Voltage_kV_Mitigated'] = raw_to_mitigated_kv(
    future_df['Predicted_Max_Voltage (kV)'].values, imax)
future_df['Predicted_Min_Voltage_kV_Mitigated'] = raw_to_mitigated_kv(
    future_df['Predicted_Min_Voltage (kV)'].values, imin)
print("Stage 4 done: Mitigated voltage profile computed.")

# ---------------------------------------------------------------------------
# Stage 5: Dual-axis plots — Future voltage fluctuation: raw vs FACTS-mitigated
# FACTS reduces deviation from nominal and dampens fluctuation (PGCB ±6%).
# ---------------------------------------------------------------------------
# PGCB limits in p.u. and kV
pu_lo, pu_hi = PGCB_LOWER_PU, PGCB_UPPER_PU
kv_lo, kv_hi = pu_lo * BASE_KV, pu_hi * BASE_KV

# Summary: % of days within PGCB ±6% and max deviation from 1 p.u. (before vs after FACTS)
raw_max_pu = future_df['Predicted_Max_Voltage_kV_Raw'].values / BASE_KV
raw_min_pu = future_df['Predicted_Min_Voltage_kV_Raw'].values / BASE_KV
mit_max_pu = future_df['Predicted_Max_Voltage_kV_Mitigated'].values / BASE_KV
mit_min_pu = future_df['Predicted_Min_Voltage_kV_Mitigated'].values / BASE_KV
within_lim_raw = np.sum((raw_max_pu <= pu_hi) & (raw_min_pu >= pu_lo)) / len(future_df) * 100
within_lim_mit = np.sum((mit_max_pu <= pu_hi) & (mit_min_pu >= pu_lo)) / len(future_df) * 100
max_dev_raw = max(np.abs(raw_max_pu - 1.0).max(), np.abs(raw_min_pu - 1.0).max()) * 100
max_dev_mit = max(np.abs(mit_max_pu - 1.0).max(), np.abs(mit_min_pu - 1.0).max()) * 100
print("\n--- FACTS impact on future voltage fluctuation (2026-2027) ---")
print(f"  Days within PGCB ±6%:  Raw {within_lim_raw:.1f}%  →  With FACTS {within_lim_mit:.1f}%")
print(f"  Max deviation from 1 p.u.:  Raw {max_dev_raw:.2f}%  →  With FACTS {max_dev_mit:.2f}% (reduction: {max_dev_raw - max_dev_mit:.2f}% points)")

for label, raw_col, mit_col in [
    ('Max_Voltage', 'Predicted_Max_Voltage_kV_Raw', 'Predicted_Max_Voltage_kV_Mitigated'),
    ('Min_Voltage', 'Predicted_Min_Voltage_kV_Raw', 'Predicted_Min_Voltage_kV_Mitigated'),
]:
    fig, ax1 = plt.subplots(figsize=(14, 6))
    ax1.set_ylabel("Voltage (p.u.), base 132 kV", fontsize=11)
    ax1.set_ylim(KV_PLOT_MIN / BASE_KV, KV_PLOT_MAX / BASE_KV)
    ax1.axhline(pu_lo, color='red', linestyle='--', linewidth=0.9, alpha=0.7, label='PGCB −6%')
    ax1.axhline(pu_hi, color='red', linestyle='--', linewidth=0.9, alpha=0.7, label='PGCB +6%')
    ax1.axhline(1.0, color='gray', linestyle=':', alpha=0.5)
    ax1.plot(future_df['Date'], future_df[raw_col] / BASE_KV, color='green', linewidth=2, label='Raw (No FACTS)')
    ax1.plot(future_df['Date'], future_df[mit_col] / BASE_KV, color='red', linewidth=2, label='Mitigated (With FACTS)')
    ax1.grid(True, alpha=0.3)
    ax1.legend(loc='upper right')
    ax1.set_title(f"Future voltage fluctuation (2026-2027): {label}\nRaw forecast vs FACTS-mitigated; FACTS reduces deviation from nominal and dampens fluctuation. PGCB ±6%.")
    ax2 = ax1.twinx()
    ax2.set_ylabel("Voltage (kV)", fontsize=11)
    ax2.set_ylim(KV_PLOT_MIN, KV_PLOT_MAX)
    ax2.plot(future_df['Date'], future_df[raw_col], color='green', alpha=0.4, linewidth=1)
    ax2.plot(future_df['Date'], future_df[mit_col], color='red', alpha=0.4, linewidth=1)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, f'stage5_dual_axis_{label}.png'), dpi=300, bbox_inches='tight')
    plt.close()
    print(f"  Saved: stage5_dual_axis_{label}.png")

# Optional: save forecast table with raw + mitigated voltages
future_df.to_excel(os.path.join(OUTPUT_DIR, 'forecast_2026_2027_with_mitigation.xlsx'), index=False)
print("\nAll stages complete. Outputs in '%s'." % OUTPUT_DIR)
