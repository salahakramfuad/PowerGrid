# PowerGrid — 132 kV Northern Bangladesh

**EEE Capstone: ML-based load & voltage forecasting with PSS/E FACTS integration**

This project forecasts load and voltage for the 132 kV Northern Bangladesh grid (Panchagarh substation), validates accuracy on 2025 data, and integrates PSS/E-derived FACTS improvement to produce raw vs mitigated voltage profiles (2026–2027) against PGCB statutory limits (±6%).

---

## Technologies & Libraries Used

| Category | Tool / Library | Purpose |
|----------|----------------|---------|
| **Language** | Python 3.9 | Core implementation |
| **Data** | **pandas** (≥1.5.0) | DataFrames, date handling, Excel/CSV I/O |
| | **numpy** (≥1.21.0) | Numerical arrays and indexing |
| | **openpyxl** (≥3.0.0) | Read/write Excel (.xlsx) |
| **ML** | **XGBoost** (≥1.7.0) | Gradient-boosted regression (primary) |
| | **scikit-learn** (≥1.0.0) | Fallback `GradientBoostingRegressor`, MAPE, train/val split |
| **Interpolation** | **scipy** (≥1.9.0) | Cubic spline (`interp1d`) for 12-month → 365-day improvement |
| **Visualization** | **matplotlib** (≥3.5.0) | Time-series and dual-axis plots |
| | **seaborn** (≥0.12.0) | Styling and analysis plots in `predict.py` |

---

## Data Sources & Inputs

- **`combinedtill25.xlsx`** — Historical daily data: Date, Substation, Load (MW), Max/Min Voltage (kV) and times. Used to build training/validation and to create `processed_data_step1.csv`.
- **`pssemaxval.xlsx`** — PSS/E simulation results for **max** voltage (monthly, 4 conditions: Solar, Solar+FACTS, Normal, Normal+FACTS) at Panchagarh 132 kV.
- **`psseminval.xlsx`** — PSS/E simulation results for **min** voltage (same structure).

---

## Pipeline Overview

The workflow is split into **data preparation**, **PSS/E improvement processing**, **ML training/validation**, and **capstone integration** (forecast + mitigation + plots).

```
combinedtill25.xlsx  ──► dataprocess.py  ──► processed_data_step1.csv
                                                      │
pssemaxval.xlsx  ──► mlprep.py  ──► panchagarh_analysis/Panchagarh_365_Improvement.xlsx
psseminval.xlsx  ──┘                                      │
                                                          │
processed_data_step1.csv  ──► predict.py  ──► predictions + analysis plots
                                                          │
processed_data_step1.csv  ──► capstone_integration.py  ──► forecast + mitigation + dual-axis plots
Panchagarh_365_Improvement.xlsx ──────────────────────┘
```

---

## ML Flowchart

We **train** ML on historical data from **2018 to 2024**, **predict** daily Load and Max/Min Voltage for **2025, 2026, and 2027**, **compare** 2025 predictions to actuals to compute **error percentage** (MAPE/MAE), and **apply FACTS scaling** to the voltage forecast to show how **future fluctuation** is impacted by FACTS application (reduced deviation from nominal, more days within PGCB ±6%).

```mermaid
flowchart LR
  subgraph data [Data]
    A[combinedtill25.xlsx]
    B[processed_data_step1.csv]
  end
  subgraph prep [Preprocessing]
    C[Date features]
    D[Is_Irrigation Feb-May]
    E[Is_Weekend]
  end
  subgraph train [Training 2018-2024]
    F[XGBoost]
    G[Random Forest]
  end
  subgraph pred [Prediction]
    H[2025 Validation]
    I[2025-2027 Forecast]
  end
  subgraph eval [Evaluation]
    J[Error percent MAPE/MAE]
    K[Model comparison plot]
  end
  subgraph facts [FACTS]
    L[PSS/E 365-day improvement]
    M[Mitigated voltage 2026-2027]
  end
  A --> C
  C --> D
  D --> E
  E --> B
  B --> F
  B --> G
  F --> H
  G --> H
  F --> I
  G --> I
  H --> J
  J --> K
  I --> L
  L --> M
```

---

## Methods Used to Produce the Output

### 1. Data preprocessing (`dataprocess.py`)

- **Input:** `combinedtill25.xlsx`
- **Steps:**
  - Parse `Date` with `pd.to_datetime`.
  - Create **time-based features:** `Year`, `Month`, `Day`, `DayOfYear`, `DayOfWeek`.
  - Create **Bangladesh-specific features:**
    - **Irrigation season:** `Is_Irrigation = 1` for Feb–May (months 2–5), else 0.
    - **Weekend:** `Is_Weekend = 1` for Fri/Sat (`DayOfWeek` 4, 5), else 0.
- **Output:** `processed_data_step1.csv` (same rows, added feature columns).

---

### 2. PSS/E improvement to 365-day lookup (`mlprep.py`)

- **Input:** `pssemaxval.xlsx`, `psseminval.xlsx` (monthly PSS/E values for Panchagarh 132 kV).
- **Steps:**
  - **Extract** monthly improvement: for each of Max/Min, compute % reduction in voltage deviation from 1.0 p.u.:
    - **Daytime:** Solar vs Solar+FACTS (cols 1 vs 2).
    - **Nighttime:** Normal vs Normal+FACTS (cols 3 vs 4).
  - **Interpolate** 12 months → 365 days using **cubic spline** (`scipy.interpolate.interp1d`, `kind='cubic'`, `fill_value='extrapolate'`).
  - **Pivot** to a table: Day (1–365) × (Max/Min × Daytime/Nighttime) → improvement %.
- **Outputs:**
  - `panchagarh_analysis/Panchagarh_365_Improvement.xlsx` — used by `capstone_integration.py`.
  - `panchagarh_analysis/Panchagarh_Max_Min_Improvement.png` — improvement curves.

---

### 3. ML training, validation & forecast (`predict.py`)

- **Input:** `processed_data_step1.csv`
- **Features:** `Year`, `Month`, `DayOfYear`, `DayOfWeek`, `Is_Irrigation` (Feb–May), `Is_Weekend`
- **Targets:** `Load (MW)` (includes irrigation impact), `Max_Voltage (kV)`, `Min_Voltage (kV)`
- **Methods:**
  - **Cleaning:** Convert targets to numeric; drop rows with NaN in targets.
  - **Split:** Train = years 2018–2024; Validation = year 2025 only.
  - **Models:** **XGBoost** and **Random Forest** (one per target each); fit on `train_df[features]` → `train_df[target]`.
  - **Validation:** Predict 2025 with both models; report **MAPE** and **MAE** per target; save **error comparison plot** (XGBoost vs Random Forest).
  - **Forecast:** Generate daily dates 2025–2027; predict all three targets (Load, Max_Voltage, Min_Voltage) with both models.
- **Outputs:**
  - `predictions_2025_validation.xlsx` — 2025 actual vs XGBoost vs Random Forest.
  - `predictions_2025_2027_forecast.xlsx` — daily forecasts 2025–2027.
  - `error_comparison_XGB_vs_RF.png` — MAPE and MAE bar chart (XGBoost vs Random Forest) for Load, Max Voltage, Min Voltage.
  - `error_timeseries_2025_XGB_vs_RF.png` — daily absolute % error over 2025 for both models.
  - `analysis_*.png` — per-target report: validation, forecast, error %, feature importance (XGBoost).

---

### Understanding the analysis plots (`analysis_*.png`)

Each `analysis_*.png` file (one per target: Load, Max_Voltage, Min_Voltage) contains four panels:

1. **Step 4: Validation (MAPE Analysis)**  
   - **X-axis:** Calendar date — each point is one day in the year **2025**. The axis shows the actual dates (e.g. Jan 2025 through Dec 2025) so you can see how well the model tracks the real values over time.  
   - **Y-axis:** The target variable (e.g. Load in MW or Voltage in kV).  
   - **What it shows:** Black line = actual 2025 values; red dashed line = model predictions for 2025. The title includes the overall **MAPE** (Mean Absolute Percentage Error) for that target.

2. **Step 3: Future Forecast (2025-2027)**  
   - Model predictions for daily values from 2025 to 2027. For voltage targets, the y-axis is fixed at 100–145 kV with a faint 132 kV nominal reference line.

3. **Error percentage graph (2025)**  
   - **X-axis:** Same as the validation plot — **calendar date** (each day in 2025).  
   - **Y-axis:** **Absolute percentage error (%)** — for each day, \(|actual - predicted| / actual \times 100\).  
   - **What it shows:** How wrong the model was on each day in percentage terms. The orange line is the daily error; the red dashed line is the mean (MAPE) across the year. This panel makes it easy to see which periods had the largest errors.

4. **What Drives the Forecast?**  
   - This is **feature importance** from the XGBoost model. Each horizontal bar is one input feature (`Year`, `Month`, `DayOfYear`, `DayOfWeek`, `Is_Irrigation`, `Is_Weekend`). Bar length shows how much that feature contributed to the model’s predictions — longer bars mean the model relied on that variable more. For example, if `Is_Irrigation` has the longest bar for Min_Voltage, irrigation season is a strong driver of minimum voltage in the model.

---

### 4. Capstone integration: forecast + FACTS mitigation + plots (`capstone_integration.py`)

Five stages:

- **Stage 1 — ML forecast (2026–2027)**  
  Same data and model setup as `predict.py` (train ≤2024, features/targets). Forecast only 2026–2027. Plot Load, Max Voltage, Min Voltage time series. Save `stage1_forecast_2026_2027.png`.

- **Stage 2 — Validation (2025)**  
  Predict 2025 with the same models; print MAPE for each target.

- **Stage 3 — PSS/E scenario lookup**  
  Load `panchagarh_analysis/Panchagarh_365_Improvement.xlsx`. For each day-of-year (1–365), get average improvement for Max and Min (average of Daytime and Nighttime). Build arrays `improve_max[1..365]`, `improve_min[1..365]` for use in Stage 4.

- **Stage 4 — Mitigation math**  
  For each forecast day, apply the day’s improvement % to the **raw** ML voltage (kV):
  - Convert kV → p.u. (base 132 kV).
  - Compute deviation from 1.0 p.u.; reduce deviation by `(1 - improvement_pct/100)`; convert back to p.u. then to kV.
  - Produces **Raw** and **Mitigated** voltage series.

- **Stage 5 — Dual-axis plots**  
  For Max and Min voltage: **Future voltage fluctuation (2026–2027):** raw forecast vs FACTS-mitigated; FACTS reduces deviation from nominal and dampens fluctuation.
  - Left axis: voltage in p.u. (base 132 kV); right axis: voltage in kV.
  - Y-axis range 100–145 kV (and equivalent p.u.).
  - Plot Raw (green) vs Mitigated (red); horizontal lines for PGCB ±6% (0.94 and 1.06 p.u.).
  - Save `stage5_dual_axis_Max_Voltage.png`, `stage5_dual_axis_Min_Voltage.png`.
  - Print summary: % of days within PGCB ±6% and max deviation from 1 p.u. (raw vs with FACTS).

- **Additional output:** `capstone_outputs/forecast_2026_2027_with_mitigation.xlsx` — full forecast table including raw and mitigated voltages.

---

## How to Run

1. **Environment**  
   Create a virtual environment, then:
   ```bash
   pip install -r requirements.txt
   ```

2. **Data**  
   Place in project root:
   - `combinedtill25.xlsx`
   - `pssemaxval.xlsx`
   - `psseminval.xlsx`

3. **Execution order**
   ```bash
   python dataprocess.py                    # → processed_data_step1.csv
   python mlprep.py                         # → panchagarh_analysis/ + 365-day improvement Excel
   python predict.py                        # → validation Excel, forecast Excel, analysis PNGs
   python capstone_integration.py           # → capstone_outputs/ (forecast, dual-axis plots, Excel)
   ```

---

## Output Summary

| Script | Main outputs |
|--------|----------------|
| `dataprocess.py` | `processed_data_step1.csv` |
| `mlprep.py` | `panchagarh_analysis/Panchagarh_365_Improvement.xlsx`, `Panchagarh_Max_Min_Improvement.png` |
| `predict.py` | `predictions_2025_validation.xlsx`, `predictions_2025_2027_forecast.xlsx`, `error_comparison_XGB_vs_RF.png`, `error_timeseries_2025_XGB_vs_RF.png`, `analysis_*.png` |
| `capstone_integration.py` | `capstone_outputs/stage1_forecast_2026_2027.png`, `stage5_dual_axis_*.png`, `forecast_2026_2027_with_mitigation.xlsx` |

---

## Key Parameters

- **Base voltage:** 132 kV  
- **PGCB limits:** 0.94 p.u. (–6%) and 1.06 p.u. (+6%)  
- **ML:** XGBoost (primary) and Random Forest for comparison — XGBoost: `n_estimators=500`, `learning_rate=0.05`, `max_depth=6`; Random Forest: `n_estimators=500`, `max_depth=12`  
- **Train period:** Up to 2024  
- **Validation year:** 2025  
- **Forecast period:** 2025–2027 (`predict.py`); 2026–2027 only for capstone plots and mitigation Excel  

### 2025 validation results (MAPE %) by learning rate

Validation uses year 2025 only (actual vs predicted). Same model setup: 500 trees, max_depth=6; only learning rate varies. Lower MAPE = better.

| Learning rate | Load (MW) | Max_Voltage (kV) | Min_Voltage (kV) |
|---------------|-----------|------------------|------------------|
| **0.01%**     | 18.23%    | 4.88%            | 7.52%            |
| **1%**        | 9.37%     | 3.26%            | 5.38%            |
| **3%**        | 9.72%     | 3.38%            | 5.37%            |
| **5%**        | 9.97%     | 3.43%            | 5.45%            |
| **7%**        | 10.20%    | 3.41%            | 5.55%            |
| **10%**       | 10.34%    | 3.36%            | 5.54%            |

The default in `predict.py` is 5% learning rate. For this dataset, load achieves best MAPE at 1%; voltage targets are similar across 1–10%.

---

This README documents the technologies used and the methods applied to obtain the forecasts, validation metrics, and FACTS-mitigated voltage profiles.
