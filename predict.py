#predictions+graph

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.metrics import mean_absolute_percentage_error

try:
    from xgboost import XGBRegressor
    Regressor = XGBRegressor
    regressor_kwargs = dict(n_estimators=500, learning_rate=0.05, max_depth=6)
except Exception:
    from sklearn.ensemble import GradientBoostingRegressor
    Regressor = GradientBoostingRegressor
    regressor_kwargs = dict(n_estimators=500, learning_rate=0.05, max_depth=6)

# 1. Load the data
df = pd.read_csv('processed_data_step1.csv')

# --- NEW DATA CLEANING STEPS ---
targets = ['Load (MW)', 'Max_Voltage (kV)', 'Min_Voltage (kV)']
for col in targets:
    df[col] = pd.to_numeric(df[col], errors='coerce')

original_count = len(df)
df = df.dropna(subset=targets)
df['Date'] = pd.to_datetime(df['Date'])
print(f"Dropped {original_count - len(df)} rows containing NaN values.")
# -------------------------------

# 2. Train/validation split: 2018-2024 for training, 2025 for validation only
train_df = df[df['Year'] <= 2024]
val_df = df[df['Year'] == 2025]
assert train_df['Year'].min() <= 2018, "Training data should include at least 2018"
assert 2025 in val_df['Year'].unique(), "Validation data should include 2025"

features = ['Year', 'Month', 'DayOfYear', 'DayOfWeek', 'Is_Irrigation', 'Is_Weekend']

# 3. Training and Validation Loop
models = {}
for target in targets:
    model = Regressor(**regressor_kwargs)
    model.fit(train_df[features], train_df[target])
    models[target] = model
    
    # Validation
    actual_2025 = val_df[target]
    predicted_2025 = model.predict(val_df[features])
    
    error = mean_absolute_percentage_error(actual_2025, predicted_2025) * 100
    print(f"Validation for {target}: Error is {error:.2f}%")

# 3b. Save 2025 validation (actual vs predicted) for accuracy evaluation
val_predictions_df = val_df[['Date', 'Year', 'Month', 'DayOfYear', 'DayOfWeek', 'Is_Irrigation', 'Is_Weekend']].copy()
for target in targets:
    val_predictions_df[f'Actual_{target}'] = val_df[target].values
    val_predictions_df[f'Predicted_{target}'] = models[target].predict(val_df[features])
val_predictions_df.to_excel('predictions_2025_validation.xlsx', index=False)
print("Saved: predictions_2025_validation.xlsx (actual vs predicted for 2025).")

# 4. Generate Predictions for 2025, 2026, 2027
future_dates = pd.date_range(start='2025-01-01', end='2027-12-31', freq='D')
future_df = pd.DataFrame({'Date': future_dates})

future_df['Year'] = future_df['Date'].dt.year
future_df['Month'] = future_df['Date'].dt.month
future_df['DayOfYear'] = future_df['Date'].dt.dayofyear
future_df['DayOfWeek'] = future_df['Date'].dt.dayofweek
future_df['Is_Irrigation'] = future_df['Month'].apply(lambda x: 1 if x in [2, 3, 4] else 0)
future_df['Is_Weekend'] = future_df['DayOfWeek'].apply(lambda x: 1 if x in [4, 5] else 0)

for target in targets:
    future_df[f'Predicted_{target}'] = models[target].predict(future_df[features])

# 5. Save forecast results (2025-2027)
future_df.to_excel('predictions_2025_2027_forecast.xlsx', index=False)
print("Success! File 'predictions_2025_2027_forecast.xlsx' created.")

import matplotlib.pyplot as plt
import seaborn as sns

print("Generating and saving graphs...")
sns.set_theme(style="whitegrid")

for target in targets:
    # Create a figure with 2x2 subplots: validation, forecast, error %, feature importance
    fig, axes = plt.subplots(2, 2, figsize=(18, 10))
    fig.suptitle(f'Project Analysis: {target} (Northern Bangladesh 132kV)', fontsize=16, fontweight='bold')

    # 1. Validation Plot: 2025 Actual vs Predicted (error validation)
    actual_2025 = val_df[target]
    predicted_2025 = models[target].predict(val_df[features])
    mape = mean_absolute_percentage_error(actual_2025, predicted_2025) * 100
    axes[0, 0].plot(val_df['Date'], val_df[target], label='Actual 2025', color='black', alpha=0.4)
    axes[0, 0].plot(val_df['Date'], predicted_2025, label='XGBoost Predicted', color='red', linestyle='--')
    axes[0, 0].set_title(f'Step 4: Validation (MAPE Analysis)\nMAPE: {mape:.2f}%')
    axes[0, 0].legend()
    axes[0, 0].set_xlabel('Date (calendar days in 2025)')

    # 2. Future Forecast: 2025-2027
    axes[0, 1].plot(future_df['Date'], future_df[f'Predicted_{target}'], color='green', linewidth=1)
    axes[0, 1].set_title('Step 3: Future Forecast (2025-2027)')
    axes[0, 1].tick_params(axis='x', rotation=45)
    if 'Voltage' in target:
        axes[0, 1].set_ylim(100, 145)
        axes[0, 1].axhline(132, color='black', alpha=0.2, label='132kV Nominal')

    # 3. Error percentage over time (daily APE for 2025)
    ape_pct = np.abs(actual_2025.values - predicted_2025) / (np.abs(actual_2025.values) + 1e-10) * 100
    axes[1, 0].plot(val_df['Date'], ape_pct, color='darkorange', alpha=0.8, linewidth=0.8)
    axes[1, 0].axhline(mape, color='red', linestyle='--', alpha=0.7, label=f'Mean (MAPE): {mape:.2f}%')
    axes[1, 0].set_title('Error Percentage (2025): Daily Absolute % Error')
    axes[1, 0].set_ylabel('Absolute percentage error (%)')
    axes[1, 0].set_xlabel('Date (calendar days in 2025)')
    axes[1, 0].legend()

    # 4. Feature Importance
    importances = models[target].feature_importances_
    feat_imp = pd.Series(importances, index=features).sort_values()
    feat_imp.plot(kind='barh', ax=axes[1, 1], color='skyblue')
    axes[1, 1].set_title('What Drives the Forecast?')

    plt.tight_layout(rect=[0, 0.03, 1, 0.95])
    
    # Save the file so you can open it manually
    filename = f"analysis_{target.replace(' ', '_').replace('(', '').replace(')', '')}.png"
    plt.savefig(filename, dpi=300)
    print(f"Saved: {filename}")
    
    plt.show() # This will attempt to open a window
