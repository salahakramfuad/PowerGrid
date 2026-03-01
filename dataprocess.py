# -*- coding: utf-8 -*-
"""
Created on Tue Feb 24 03:07:42 2026

@author: USER
"""

import pandas as pd

# 1. Load the dataset
file_path = 'combinedtill25.xlsx'  # Ensure this file is in your project folder
df = pd.read_excel(file_path)

# 2. Convert 'Date' to a format Python understands
# This allows us to "pull out" the year, month, and day easily
df['Date'] = pd.to_datetime(df['Date'])

# 3. Create Time-Based Features
df['Year'] = df['Date'].dt.year
df['Month'] = df['Date'].dt.month
df['Day'] = df['Date'].dt.day
df['DayOfYear'] = df['Date'].dt.dayofyear
df['DayOfWeek'] = df['Date'].dt.dayofweek  # 0=Monday, 6=Sunday

# 4. Create Bangladesh-Specific Features
# Irrigation Season: Peak load in Northern BD — Feb, Mar, Apr, May (4 months)
df['Is_Irrigation'] = df['Month'].apply(lambda x: 1 if x in [2, 3, 4, 5] else 0)

# Weekend Feature: In Bangladesh, Friday (4) and Saturday (5) usually see lower industrial load
df['Is_Weekend'] = df['DayOfWeek'].apply(lambda x: 1 if x in [4, 5] else 0)

# 5. Clean Up for ML
# We keep the 'Date' for plotting later, but we create a new dataframe (X) 
# that only contains numbers for the actual Machine Learning training.
features_for_ml = ['Year', 'Month', 'DayOfYear', 'DayOfWeek', 'Is_Irrigation', 'Is_Weekend']
targets = ['Load (MW)', 'Max_Voltage (kV)', 'Min_Voltage (kV)']

# Show the first few rows of your "engineered" data
print("Feature Engineering Complete. New Columns Created:")
print(df[['Date', 'Year', 'Month', 'Is_Irrigation', 'Is_Weekend']].head())

# Save this version so you don't have to run this step again
df.to_csv('processed_data_step1.csv', index=False)