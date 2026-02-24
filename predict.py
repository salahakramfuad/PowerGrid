#predictions+graph

import pandas as pd
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
    # Create a figure with 3 subplots for a full report
    fig, axes = plt.subplots(1, 3, figsize=(20, 6))
    fig.suptitle(f'Analysis for {target}', fontsize=16)

    # 1. Validation Plot: 2025 Actual vs Predicted
    # This helps you see where the MAPE error is actually happening
    axes[0].plot(val_df.index, val_df[target], label='Actual 2025', color='black', alpha=0.5)
    axes[0].plot(val_df.index, models[target].predict(val_df[features]), label='Predicted 2025', color='red', linestyle='--')
    axes[0].set_title('Model Accuracy (2025)')
    axes[0].legend()

    # 2. Future Forecast: 2025-2027
    axes[1].plot(future_df['Date'], future_df[f'Predicted_{target}'], color='green')
    axes[1].set_title('Forecast (2025-2027)')
    axes[1].tick_params(axis='x', rotation=45)

    # 3. Feature Importance
    # Shows which factors (Irrigation, Weekends, etc.) moved the needle most
    importances = models[target].feature_importances_
    feat_imp = pd.Series(importances, index=features).sort_values()
    feat_imp.plot(kind='barh', ax=axes[2], color='skyblue')
    axes[2].set_title('What Drives these Predictions?')

    plt.tight_layout(rect=[0, 0.03, 1, 0.95])
    
    # Save the file so you can open it manually
    filename = f"analysis_{target.replace(' ', '_').replace('(', '').replace(')', '')}.png"
    plt.savefig(filename)
    print(f"Saved: {filename}")
    
    plt.show() # This will attempt to open a window
