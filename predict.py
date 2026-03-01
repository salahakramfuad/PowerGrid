# predictions + graphs: XGBoost vs Random Forest, Load (MW) includes irrigation (Feb–May) impact

import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.metrics import mean_absolute_percentage_error, mean_absolute_error

try:
    from xgboost import XGBRegressor
    XGB_kwargs = dict(n_estimators=500, learning_rate=0.05, max_depth=6)
except Exception:
    XGBRegressor = None
    XGB_kwargs = {}

from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor

# Second algorithm for comparison: Random Forest
RF_kwargs = dict(n_estimators=500, max_depth=12, random_state=42)

OUTPUT_DIR = "new output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 1. Load the data
df = pd.read_csv('processed_data_step1.csv')

targets = ['Load (MW)', 'Max_Voltage (kV)', 'Min_Voltage (kV)']
for col in targets:
    df[col] = pd.to_numeric(df[col], errors='coerce')

original_count = len(df)
df = df.dropna(subset=targets)
df['Date'] = pd.to_datetime(df['Date'])
print(f"Dropped {original_count - len(df)} rows containing NaN values.")

# 2. Train/validation split: 2018-2024 for training, 2025 for validation
train_df = df[df['Year'] <= 2024]
val_df = df[df['Year'] == 2025]
assert train_df['Year'].min() <= 2018, "Training data should include at least 2018"
assert 2025 in val_df['Year'].unique(), "Validation data should include 2025"

features = ['Year', 'Month', 'DayOfYear', 'DayOfWeek', 'Is_Irrigation', 'Is_Weekend']

# 3. Train both algorithms per target
models_xgb = {}
models_rf = {}
if XGBRegressor is not None:
    for target in targets:
        models_xgb[target] = XGBRegressor(**XGB_kwargs)
        models_xgb[target].fit(train_df[features], train_df[target])
else:
    for target in targets:
        models_xgb[target] = GradientBoostingRegressor(n_estimators=500, learning_rate=0.05, max_depth=6)
        models_xgb[target].fit(train_df[features], train_df[target])

for target in targets:
    models_rf[target] = RandomForestRegressor(**RF_kwargs)
    models_rf[target].fit(train_df[features], train_df[target])

# 4. Validation 2025: MAPE and MAE per model per target
print("\n--- Validation 2025 (error % and MAE) ---")
val_metrics = []
for target in targets:
    actual = val_df[target].values
    pred_xgb = models_xgb[target].predict(val_df[features])
    pred_rf = models_rf[target].predict(val_df[features])
    mape_xgb = mean_absolute_percentage_error(actual, pred_xgb) * 100
    mape_rf = mean_absolute_percentage_error(actual, pred_rf) * 100
    mae_xgb = mean_absolute_error(actual, pred_xgb)
    mae_rf = mean_absolute_error(actual, pred_rf)
    print(f"{target}: XGBoost MAPE={mape_xgb:.2f}% MAE={mae_xgb:.4g}  |  RandomForest MAPE={mape_rf:.2f}% MAE={mae_rf:.4g}")
    val_metrics.append({
        'Target': target,
        'MAPE_XGB': mape_xgb, 'MAPE_RF': mape_rf,
        'MAE_XGB': mae_xgb, 'MAE_RF': mae_rf,
    })

# 5. Save 2025 validation with both models' predictions (includes Load impact)
val_predictions_df = val_df[['Date', 'Year', 'Month', 'DayOfYear', 'DayOfWeek', 'Is_Irrigation', 'Is_Weekend']].copy()
for target in targets:
    val_predictions_df[f'Actual_{target}'] = val_df[target].values
    val_predictions_df[f'Predicted_{target}_XGB'] = models_xgb[target].predict(val_df[features])
    val_predictions_df[f'Predicted_{target}_RF'] = models_rf[target].predict(val_df[features])
val_predictions_df.to_excel('predictions_2025_validation.xlsx', index=False)
print("Saved: predictions_2025_validation.xlsx (actual vs XGBoost vs Random Forest for 2025).")

# 6. Forecast 2025-2027 with both models
future_dates = pd.date_range(start='2025-01-01', end='2027-12-31', freq='D')
future_df = pd.DataFrame({'Date': future_dates})
future_df['Year'] = future_df['Date'].dt.year
future_df['Month'] = future_df['Date'].dt.month
future_df['DayOfYear'] = future_df['Date'].dt.dayofyear
future_df['DayOfWeek'] = future_df['Date'].dt.dayofweek
future_df['Is_Irrigation'] = future_df['Month'].apply(lambda x: 1 if x in [2, 3, 4, 5] else 0)
future_df['Is_Weekend'] = future_df['DayOfWeek'].apply(lambda x: 1 if x in [4, 5] else 0)

for target in targets:
    future_df[f'Predicted_{target}_XGB'] = models_xgb[target].predict(future_df[features])
    future_df[f'Predicted_{target}_RF'] = models_rf[target].predict(future_df[features])
# Primary forecast columns (XGB) for backward compatibility
for target in targets:
    future_df[f'Predicted_{target}'] = future_df[f'Predicted_{target}_XGB']

future_df.to_excel('predictions_2025_2027_forecast.xlsx', index=False)
print("Success! File 'predictions_2025_2027_forecast.xlsx' created (XGB + RF).")

# 7. Error comparison plot: MAPE and MAE (XGBoost vs Random Forest) — includes Load (MW)
sns.set_theme(style="whitegrid")
metrics_df = pd.DataFrame(val_metrics)
target_labels = ['Load (MW)', 'Max_Voltage (kV)', 'Min_Voltage (kV)']
short_labels = ['Load (MW)', 'Max Voltage', 'Min Voltage']

fig, (ax_mape, ax_mae) = plt.subplots(1, 2, figsize=(12, 5))
x = np.arange(len(targets))
w = 0.35
ax_mape.bar(x - w/2, metrics_df['MAPE_XGB'], w, label='XGBoost', color='steelblue')
ax_mape.bar(x + w/2, metrics_df['MAPE_RF'], w, label='Random Forest', color='coral')
ax_mape.set_xticks(x)
ax_mape.set_xticklabels(short_labels)
ax_mape.set_ylabel('MAPE (%)')
ax_mape.set_title('Validation 2025: MAPE comparison (XGBoost vs Random Forest)\nLoad (MW) includes irrigation Feb–May impact')
ax_mape.legend()
ax_mape.grid(True, alpha=0.3)

ax_mae.bar(x - w/2, metrics_df['MAE_XGB'], w, label='XGBoost', color='steelblue')
ax_mae.bar(x + w/2, metrics_df['MAE_RF'], w, label='Random Forest', color='coral')
ax_mae.set_xticks(x)
ax_mae.set_xticklabels(short_labels)
ax_mae.set_ylabel('MAE')
ax_mae.set_title('Validation 2025: MAE comparison (XGBoost vs Random Forest)')
ax_mae.legend()
ax_mae.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig(os.path.join(OUTPUT_DIR, 'error_comparison_XGB_vs_RF.png'), dpi=300, bbox_inches='tight')
plt.close()
print(f"Saved: {OUTPUT_DIR}/error_comparison_XGB_vs_RF.png")

# 8. Optional: time-series of daily APE for 2025 (both models) per target
fig2, axes = plt.subplots(3, 1, figsize=(12, 9), sharex=True)
for i, target in enumerate(targets):
    actual = val_df[target].values
    pred_xgb = models_xgb[target].predict(val_df[features])
    pred_rf = models_rf[target].predict(val_df[features])
    ape_xgb = np.abs(actual - pred_xgb) / (np.abs(actual) + 1e-10) * 100
    ape_rf = np.abs(actual - pred_rf) / (np.abs(actual) + 1e-10) * 100
    axes[i].plot(val_df['Date'], ape_xgb, label='XGBoost', color='steelblue', alpha=0.8)
    axes[i].plot(val_df['Date'], ape_rf, label='Random Forest', color='coral', alpha=0.8)
    axes[i].set_ylabel('Absolute % Error')
    axes[i].set_title(f'{target} — Daily error 2025')
    axes[i].legend()
    axes[i].grid(True, alpha=0.3)
axes[-1].set_xlabel('Date (2025)')
fig2.suptitle('Validation 2025: Daily absolute percentage error — XGBoost vs Random Forest', fontsize=12, fontweight='bold')
plt.tight_layout(rect=[0, 0, 1, 0.96])
plt.savefig(os.path.join(OUTPUT_DIR, 'error_timeseries_2025_XGB_vs_RF.png'), dpi=300, bbox_inches='tight')
plt.close()
print(f"Saved: {OUTPUT_DIR}/error_timeseries_2025_XGB_vs_RF.png")

# 9. Per-target analysis figures (using XGBoost as primary model)
for target in targets:
    fig, axes = plt.subplots(2, 2, figsize=(18, 10))
    fig.suptitle(f'Project Analysis: {target} (Northern Bangladesh 132kV)\nLoad (MW) includes irrigation Feb–May impact', fontsize=16, fontweight='bold')

    actual_2025 = val_df[target]
    predicted_2025 = models_xgb[target].predict(val_df[features])
    mape = mean_absolute_percentage_error(actual_2025, predicted_2025) * 100

    axes[0, 0].plot(val_df['Date'], val_df[target], label='Actual 2025', color='black', alpha=0.4)
    axes[0, 0].plot(val_df['Date'], predicted_2025, label='XGBoost Predicted', color='red', linestyle='--')
    axes[0, 0].set_title(f'Validation (MAPE Analysis)\nMAPE: {mape:.2f}%')
    axes[0, 0].legend()
    axes[0, 0].set_xlabel('Date (calendar days in 2025)')

    axes[0, 1].plot(future_df['Date'], future_df[f'Predicted_{target}'], color='green', linewidth=1)
    axes[0, 1].set_title('Future Forecast (2025-2027)')
    axes[0, 1].tick_params(axis='x', rotation=45)
    if 'Voltage' in target:
        axes[0, 1].set_ylim(100, 145)
        axes[0, 1].axhline(132, color='black', alpha=0.2, label='132kV Nominal')

    ape_pct = np.abs(actual_2025.values - predicted_2025) / (np.abs(actual_2025.values) + 1e-10) * 100
    axes[1, 0].plot(val_df['Date'], ape_pct, color='darkorange', alpha=0.8, linewidth=0.8)
    axes[1, 0].axhline(mape, color='red', linestyle='--', alpha=0.7, label=f'Mean MAPE: {mape:.2f}%')
    axes[1, 0].set_title('Error Percentage (2025): Daily Absolute % Error')
    axes[1, 0].set_ylabel('Absolute percentage error (%)')
    axes[1, 0].set_xlabel('Date (calendar days in 2025)')
    axes[1, 0].legend()
    axes[1, 0].grid(True, alpha=0.3)

    importances = models_xgb[target].feature_importances_
    feat_imp = pd.Series(importances, index=features).sort_values()
    feat_imp.plot(kind='barh', ax=axes[1, 1], color='skyblue')
    axes[1, 1].set_title('Feature Importance (XGBoost)')

    plt.tight_layout(rect=[0, 0.03, 1, 0.95])
    filename = f"analysis_{target.replace(' ', '_').replace('(', '').replace(')', '')}.png"
    plt.savefig(os.path.join(OUTPUT_DIR, filename), dpi=300)
    print(f"Saved: {OUTPUT_DIR}/{filename}")
    plt.close()

print(f"\nDone. All PNGs saved in '{OUTPUT_DIR}/'.")
