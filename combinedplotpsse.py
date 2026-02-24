import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# ========== CONFIGURATION ==========
# Ensure these filenames match your actual Excel files
FILES_TO_PROCESS = [
    {"filename": "pssemaxval.xlsx", "suffix": "Max"},
    {"filename": "psseminval.xlsx", "suffix": "Min"}
]

STATIONS = ["Panchagarh", "Thakurgaon", "Purbasadipur"]
MONTHS_ORDER = ["Jan","Feb","Mar","Apr","May","Jun",
                "Jul","Aug","Sep","Oct","Nov","Dec"]

FULL_TO_SHORT_MONTH = {
    "January":"Jan","February":"Feb","March":"Mar","April":"Apr","May":"May","June":"Jun",
    "July":"Jul","August":"Aug","September":"Sep","October":"Oct","November":"Nov","December":"Dec"
}

CONDITIONS = ["Solar", "Solar + FACTS", "Normal", "Normal + FACTS"]
BASE_KV_132 = 132.0  # Focusing only on 132kV as requested

# Original Visual Styles
LINESTYLES = {"Solar": "-", "Solar + FACTS": "-", "Normal": ":", "Normal + FACTS": "--"}
MARKERS = {"Solar": "o", "Solar + FACTS": "s", "Normal": "^", "Normal + FACTS": "D"}
COLORS = {"Solar": "#f39c12", "Solar + FACTS": "#e67e22", "Normal": "#3498db", "Normal + FACTS": "#2980b9"}

# Create a folder for the new plots
os.makedirs("plots_132kv_only", exist_ok=True)

# ========== PROCESSING ENGINE ==========
def run_simulation_plots():
    for file_info in FILES_TO_PROCESS:
        file_path = file_info["filename"]
        label = file_info["suffix"]

        if not os.path.exists(file_path):
            print(f"❌ Error: {file_path} not found. Skipping...")
            continue

        print(f"🔍 Processing {label} values from {file_path}...")
        df = pd.read_excel(file_path, header=None)
        
        # Data structure for 132kV only
        data = {st: {c: {} for c in CONDITIONS} for st in STATIONS}

        current_month = None
        expecting_132kv_row = False
        target_station = None

        # Parse Excel
        for r in range(df.shape[0]):
            raw_val = df.iat[r, 0]
            cellA = str(raw_val).strip() if not pd.isna(raw_val) else None
            
            # Check for Month
            if cellA in FULL_TO_SHORT_MONTH:
                current_month = FULL_TO_SHORT_MONTH[cellA]
                continue

            # Check for Station (33kV row)
            if cellA in STATIONS and current_month:
                target_station = cellA
                expecting_132kv_row = True # The next row is the 132kV row
                continue

            # Process the 132kV Row (The blank row below the station name)
            if expecting_132kv_row and (cellA is None or cellA == "nan" or cellA == "") and current_month:
                for cond, col in zip(CONDITIONS, [1, 2, 3, 4]):
                    val = df.iat[r, col]
                    data[target_station][cond][current_month] = float(val) if not pd.isna(val) else np.nan
                expecting_132kv_row = False

        # Generate the 3 plots for this file
        for station in STATIONS:
            # Grid Code Limits
            LOWER_LIMIT = 0.94
            UPPER_LIMIT = 1.06

            all_values = []
            for cond in CONDITIONS:
                pts = [data[station][cond].get(m, np.nan) for m in MONTHS_ORDER]
                all_values.extend([v/BASE_KV_132 for v in pts if not np.isnan(v)])

            if not all_values:
                continue

            # Dynamic scaling
            vmin, vmax = min(all_values), max(all_values)
            padding = max((vmax - vmin) * 0.2, 0.04)
            plot_min, plot_max = min(LOWER_LIMIT - 0.02, vmin - padding), max(UPPER_LIMIT + 0.02, vmax + padding)

            fig, ax1 = plt.subplots(figsize=(13, 6.5))

            # Visual decorations
            ax1.axhspan(LOWER_LIMIT, UPPER_LIMIT, color='green', alpha=0.07, label='PGCB Grid Code Limit')
            ax1.axhline(LOWER_LIMIT, color='red', linestyle='--', linewidth=0.9, alpha=0.4)
            ax1.axhline(UPPER_LIMIT, color='red', linestyle='--', linewidth=0.9, alpha=0.4)
            ax1.axhline(1.0, color='black', linewidth=1, alpha=0.2)

            # Plot Lines
            for cond in CONDITIONS:
                ys_pu = [data[station][cond].get(m, np.nan) / BASE_KV_132 for m in MONTHS_ORDER]
                series = pd.Series(ys_pu)
                mask = ~series.isna()
                
                if mask.any():
                    ax1.plot(np.array(MONTHS_ORDER)[mask], series[mask],
                             linestyle=LINESTYLES[cond], marker=MARKERS[cond],
                             color=COLORS[cond], linewidth=2.5, label=cond, zorder=5)

            ax1.set_ylim(plot_min, plot_max)
            ax1.set_ylabel("Voltage (per-unit)", fontsize=12, fontweight='bold')
            ax1.set_title(f"Simulation Analysis ({label}): {station} 132kV Bus\n$\pm$6% Tolerance Band", 
                          fontsize=14, pad=20, fontweight='bold')
            ax1.grid(True, which='both', linestyle=':', alpha=0.6)
            ax1.legend(loc="upper left", bbox_to_anchor=(1.08, 1.0))

            # Secondary kV Axis
            ax2 = ax1.twinx()
            ax2.set_ylim(plot_min * BASE_KV_132, plot_max * BASE_KV_132)
            ax2.set_ylabel("Absolute Voltage (kV)", fontsize=12, color='#7f8c8d')

            plt.tight_layout()
            save_path = f"plots_132kv_only/{station}_132kV_{label}_Analysis.png"
            plt.savefig(save_path, dpi=300, bbox_inches="tight")
            plt.close()
            print(f"✅ Saved: {save_path}")

if __name__ == "__main__":
    run_simulation_plots()
    print("\n🎯 Completed! 6 plots (3 Max, 3 Min) for 132kV are ready.")