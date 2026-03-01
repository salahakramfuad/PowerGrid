# -*- coding: utf-8 -*-
"""
Comparative Day/Night plots for xyz.xlsx and xyzmax.xlsx.
Loads data from sheets Panchagarh, Purbasadipur, Thakurgaon; produces one
comparative figure per substation with Day vs Night grouped and File 1 vs File 2
side-by-side. Supports alternative views: line, diff, heatmap, table.
"""

import argparse
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# ========== CONFIG ==========
FILE1 = "xyz.xlsx"
FILE2 = "xyzmax.xlsx"
SUBSTATIONS = ["Panchagarh", "Purbasadipur", "Thakurgaon"]
OUTPUT_DIR = "comparison_plots"
FILE1_LABEL = "File 1 (xyz)"
FILE2_LABEL = "File 2 (xyzmax)"


def load_excel(path):
    """Load Excel file; return dict {sheet_name_lower: DataFrame}."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"File not found: {path}")
    xl = pd.ExcelFile(path)
    return {str(s).strip().lower(): pd.read_excel(path, sheet_name=s) for s in xl.sheet_names}


def normalize_columns(df):
    """Strip and lowercase column names; return df with cols: operating_mode, day, night."""
    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]
    # Map common variants to canonical names
    col_map = {}
    for c in df.columns:
        c_clean = c.strip().lower()
        if "operating" in c_clean and "mode" in c_clean:
            col_map[c] = "operating_mode"
        elif c_clean == "day" or c_clean.rstrip() == "day":
            col_map[c] = "day"
        elif c_clean == "night":
            col_map[c] = "night"
    df = df.rename(columns=col_map)
    required = {"day", "night"}
    if not required.issubset(df.columns):
        return None
    if "operating_mode" not in df.columns:
        df["operating_mode"] = np.arange(len(df))
    return df[["operating_mode", "day", "night"]].copy()


def get_sheet_key(name):
    """Canonical key for case-insensitive matching."""
    return name.strip().lower().replace(" ", "")


def load_all_data():
    """
    Load both files, match sheets by case-insensitive name, align by row index.
    Returns dict[substation_name] -> DataFrame with columns:
      scenario (0..n-1), file1_day, file1_night, file2_day, file2_night [, operating_mode]
    """
    data1 = load_excel(FILE1)
    data2 = load_excel(FILE2)
    keys1 = {get_sheet_key(k): k for k in data1}
    keys2 = {get_sheet_key(k): k for k in data2}

    result = {}
    for display_name in SUBSTATIONS:
        key = get_sheet_key(display_name)
        if key not in keys1:
            print(f"Warning: sheet for '{display_name}' not found in {FILE1}; skipping.")
            continue
        if key not in keys2:
            print(f"Warning: sheet for '{display_name}' not found in {FILE2}; skipping.")
            continue
        df1 = normalize_columns(data1[keys1[key]])
        df2 = normalize_columns(data2[keys2[key]])
        if df1 is None:
            print(f"Warning: missing day/night columns for '{display_name}' in {FILE1}; skipping.")
            continue
        if df2 is None:
            print(f"Warning: missing day/night columns for '{display_name}' in {FILE2}; skipping.")
            continue
        n = min(len(df1), len(df2))
        combined = pd.DataFrame({
            "scenario": np.arange(n),
            "file1_day": df1["day"].values[:n],
            "file1_night": df1["night"].values[:n],
            "file2_day": df2["day"].values[:n],
            "file2_night": df2["night"].values[:n],
        })
        result[display_name] = combined
    return result


def plot_grouped_bar(substation, data, out_dir):
    """
    One figure per substation: two logical groups (Day, Night).
    Each group: side-by-side bars File1 vs File2 over scenario index.
    """
    n = len(data)
    x = np.arange(n)
    width = 0.35

    fig, (ax_day, ax_night) = plt.subplots(1, 2, figsize=(12, 5), sharey=False)
    fig.suptitle(f"{substation} — {FILE1_LABEL} vs {FILE2_LABEL}", fontsize=13, fontweight="bold")

    # Day: File1 vs File2
    ax_day.bar(x - width / 2, data["file1_day"], width, label=FILE1_LABEL, color="steelblue", alpha=0.9)
    ax_day.bar(x + width / 2, data["file2_day"], width, label=FILE2_LABEL, color="coral", alpha=0.9)
    ax_day.set_xlabel("Scenario (row index)")
    ax_day.set_ylabel("Value")
    ax_day.set_title("Day")
    ax_day.legend()
    ax_day.grid(True, alpha=0.3)

    # Night: File1 vs File2
    ax_night.bar(x - width / 2, data["file1_night"], width, label=FILE1_LABEL, color="steelblue", alpha=0.9)
    ax_night.bar(x + width / 2, data["file2_night"], width, label=FILE2_LABEL, color="coral", alpha=0.9)
    ax_night.set_xlabel("Scenario (row index)")
    ax_night.set_ylabel("Value")
    ax_night.set_title("Night")
    ax_night.legend()
    ax_night.grid(True, alpha=0.3)

    plt.tight_layout()
    safe_name = substation.replace(" ", "_")
    path = os.path.join(out_dir, f"comparison_{safe_name}.png")
    plt.savefig(path, dpi=300, bbox_inches="tight")
    plt.close()
    return path


def plot_line(substation, data, out_dir):
    """Single axes: 4 lines — File1 Day, File1 Night, File2 Day, File2 Night vs scenario."""
    fig, ax = plt.subplots(figsize=(10, 5))
    x = data["scenario"]
    ax.plot(x, data["file1_day"], "o-", label=f"{FILE1_LABEL} Day", color="steelblue", linewidth=2)
    ax.plot(x, data["file1_night"], "s--", label=f"{FILE1_LABEL} Night", color="steelblue", alpha=0.7)
    ax.plot(x, data["file2_day"], "o-", label=f"{FILE2_LABEL} Day", color="coral", linewidth=2)
    ax.plot(x, data["file2_night"], "s--", label=f"{FILE2_LABEL} Night", color="coral", alpha=0.7)
    ax.set_xlabel("Scenario (row index)")
    ax.set_ylabel("Value")
    ax.set_title(f"{substation} — Line comparison")
    ax.legend()
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    safe_name = substation.replace(" ", "_")
    path = os.path.join(out_dir, f"comparison_{safe_name}_line.png")
    plt.savefig(path, dpi=300, bbox_inches="tight")
    plt.close()
    return path


def plot_diff(substation, data, out_dir):
    """Plot (File2 - File1) for Day and for Night."""
    fig, ax = plt.subplots(figsize=(10, 5))
    x = data["scenario"]
    ax.bar(x - 0.2, data["file2_day"] - data["file1_day"], 0.4, label="Day (F2 − F1)", color="steelblue", alpha=0.9)
    ax.bar(x + 0.2, data["file2_night"] - data["file1_night"], 0.4, label="Night (F2 − F1)", color="coral", alpha=0.9)
    ax.axhline(0, color="black", linewidth=0.5)
    ax.set_xlabel("Scenario (row index)")
    ax.set_ylabel("Difference (File 2 − File 1)")
    ax.set_title(f"{substation} — Shift between files")
    ax.legend()
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    safe_name = substation.replace(" ", "_")
    path = os.path.join(out_dir, f"comparison_{safe_name}_diff.png")
    plt.savefig(path, dpi=300, bbox_inches="tight")
    plt.close()
    return path


def plot_heatmap(substation, data, out_dir):
    """Heatmap: rows = scenarios, columns = Day F1, Night F1, Day F2, Night F2."""
    try:
        import seaborn as sns
    except ImportError:
        # Fallback without seaborn: use matplotlib imshow
        M = np.column_stack([
            data["file1_day"], data["file1_night"],
            data["file2_day"], data["file2_night"],
        ])
        fig, ax = plt.subplots(figsize=(8, max(4, len(data) * 0.35)))
        im = ax.imshow(M, aspect="auto", cmap="viridis")
        ax.set_xticks([0, 1, 2, 3])
        ax.set_xticklabels(["Day F1", "Night F1", "Day F2", "Night F2"])
        ax.set_ylabel("Scenario")
        ax.set_title(f"{substation} — Value heatmap")
        plt.colorbar(im, ax=ax)
        plt.tight_layout()
    else:
        df_hm = pd.DataFrame({
            "Day F1": data["file1_day"],
            "Night F1": data["file1_night"],
            "Day F2": data["file2_day"],
            "Night F2": data["file2_night"],
        })
        fig, ax = plt.subplots(figsize=(8, max(4, len(data) * 0.35)))
        sns.heatmap(df_hm.T, ax=ax, cmap="viridis", cbar_kws={"label": "Value"})
        ax.set_xlabel("Scenario")
        ax.set_ylabel("")
        ax.set_title(f"{substation} — Value heatmap")
        plt.tight_layout()
    safe_name = substation.replace(" ", "_")
    path = os.path.join(out_dir, f"comparison_{safe_name}_heatmap.png")
    plt.savefig(path, dpi=300, bbox_inches="tight")
    plt.close()
    return path


def print_table(all_data):
    """Print summary table: substation, period (Day/Night), file, mean, min, max, std."""
    rows = []
    for substation, data in all_data.items():
        for period, col1, col2 in [
            ("Day", "file1_day", "file2_day"),
            ("Night", "file1_night", "file2_night"),
        ]:
            for label, col in [(FILE1_LABEL, col1), (FILE2_LABEL, col2)]:
                s = data[col]
                rows.append({
                    "Substation": substation,
                    "Period": period,
                    "File": label,
                    "mean": s.mean(),
                    "min": s.min(),
                    "max": s.max(),
                    "std": s.std(),
                })
    df = pd.DataFrame(rows)
    print(df.to_string(index=False))
    return df


def main():
    parser = argparse.ArgumentParser(description="Comparative Day/Night plots for xyz vs xyzmax")
    parser.add_argument(
        "--view",
        choices=["bar", "line", "diff", "heatmap", "table"],
        default="bar",
        help="bar = grouped Day/Night (default); line, diff, heatmap, table",
    )
    parser.add_argument("--output-dir", default=OUTPUT_DIR, help="Directory for output figures")
    args = parser.parse_args()

    all_data = load_all_data()
    if not all_data:
        print("No substation data loaded. Check file paths and sheet names.")
        return

    os.makedirs(args.output_dir, exist_ok=True)
    saved = []

    if args.view == "table":
        df = print_table(all_data)
        out_csv = os.path.join(args.output_dir, "comparison_summary.csv")
        df.to_csv(out_csv, index=False)
        print(f"Saved: {out_csv}")
        return

    for substation, data in all_data.items():
        if args.view == "bar":
            path = plot_grouped_bar(substation, data, args.output_dir)
        elif args.view == "line":
            path = plot_line(substation, data, args.output_dir)
        elif args.view == "diff":
            path = plot_diff(substation, data, args.output_dir)
        elif args.view == "heatmap":
            path = plot_heatmap(substation, data, args.output_dir)
        else:
            continue
        saved.append(path)

    for p in saved:
        print(f"Saved: {p}")
    if saved:
        print(f"Done. View: {args.view}, output dir: {args.output_dir}")


if __name__ == "__main__":
    main()
