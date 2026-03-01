# -*- coding: utf-8 -*-
"""Generate ML flowchart as PNG using matplotlib (no external diagram tools)."""

import os
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch

OUTPUT_DIR = "new output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Figure and axis (no axes frame, we draw everything manually)
fig, ax = plt.subplots(1, 1, figsize=(16, 10))
ax.set_xlim(0, 16)
ax.set_ylim(0, 10)
ax.set_aspect("equal")
ax.axis("off")

# Style
box_style = dict(boxstyle="round,pad=0.3", facecolor="lightgray", edgecolor="black", linewidth=1.2)
group_style = dict(facecolor="none", edgecolor="gray", linewidth=1, linestyle="--")

def box(ax, x, y, w, h, text, fontsize=9):
    """Draw a rounded box with text."""
    p = FancyBboxPatch((x - w/2, y - h/2), w, h, **box_style)
    ax.add_patch(p)
    ax.text(x, y, text, ha="center", va="center", fontsize=fontsize, wrap=True)

# Node positions (x, y) — layout left to right
A = (1.5, 7)
B = (1.5, 5)
C = (3.5, 7.5)
D = (3.5, 6)
E = (3.5, 4.5)
F = (6, 7.2)
G = (6, 5.8)
H = (8.5, 7.2)
I = (8.5, 5.8)
J = (11, 7.5)
K = (11, 6.5)
L = (13, 6.5)
M = (14.5, 6.5)

w_med, h_small = 2.0, 0.65

# Draw grouped regions (background rectangles)
ax.add_patch(mpatches.Rectangle((0.3, 4.2), 2.4, 3.4, **group_style))
ax.text(1.5, 7.85, "Data", fontsize=10, fontweight="bold", ha="center")
ax.add_patch(mpatches.Rectangle((2.4, 3.8), 1.8, 4.2, **group_style))
ax.text(3.3, 8.15, "Preprocessing", fontsize=10, fontweight="bold", ha="center")
ax.add_patch(mpatches.Rectangle((5.0, 5.2), 2.2, 2.6, **group_style))
ax.text(6, 7.95, "Training 2018-2024", fontsize=10, fontweight="bold", ha="center")
ax.add_patch(mpatches.Rectangle((7.5, 5.2), 2.2, 2.6, **group_style))
ax.text(8.5, 7.95, "Prediction", fontsize=10, fontweight="bold", ha="center")
ax.add_patch(mpatches.Rectangle((10.0, 5.8), 2.2, 2.2, **group_style))
ax.text(11, 8.05, "Evaluation", fontsize=10, fontweight="bold", ha="center")
ax.add_patch(mpatches.Rectangle((12.0, 5.8), 3.2, 1.6, **group_style))
ax.text(13.6, 7.55, "FACTS", fontsize=10, fontweight="bold", ha="center")

# Boxes
box(ax, A[0], A[1], 2.0, 0.5, "combinedtill25.xlsx", 8)
box(ax, B[0], B[1], 2.0, 0.5, "processed_data_step1.csv", 8)
box(ax, C[0], C[1], w_med, h_small, "Date features", 8)
box(ax, D[0], D[1], w_med, h_small, "Is_Irrigation Feb-May", 8)
box(ax, E[0], E[1], w_med, h_small, "Is_Weekend", 8)
box(ax, F[0], F[1], w_med, h_small, "XGBoost", 9)
box(ax, G[0], G[1], w_med, h_small, "Random Forest", 9)
box(ax, H[0], H[1], w_med, h_small, "2025 Validation", 8)
box(ax, I[0], I[1], w_med, h_small, "2025-2027 Forecast", 8)
box(ax, J[0], J[1], w_med, h_small, "Error % MAPE/MAE", 8)
box(ax, K[0], K[1], w_med, h_small, "Model comparison plot", 8)
box(ax, L[0], L[1], w_med, h_small, "PSS/E 365-day improvement", 8)
box(ax, M[0], M[1], 2.2, h_small, "Mitigated voltage 2026-2027", 8)

# Arrows
def draw_arrow(xy1, xy2):
    ax.annotate("", xy=xy2, xytext=xy1,
                arrowprops=dict(arrowstyle="->", color="black", lw=1.2,
                               connectionstyle="arc3,rad=0.1"))

draw_arrow(A, C)
draw_arrow(C, D)
draw_arrow(D, E)
draw_arrow(E, B)
draw_arrow(B, F)
draw_arrow(B, G)
draw_arrow(F, H)
draw_arrow(G, H)
draw_arrow(F, I)
draw_arrow(G, I)
draw_arrow(H, J)
draw_arrow(J, K)
draw_arrow(I, L)
draw_arrow(L, M)

# Title
ax.text(8, 9.5, "ML Flowchart: 132 kV Northern Bangladesh — Train (2018-2024), Predict (2025-2027), Validate, FACTS", fontsize=12, fontweight="bold", ha="center")

out_path = os.path.join(OUTPUT_DIR, "ML_flowchart.png")
plt.tight_layout()
plt.savefig(out_path, dpi=300, bbox_inches="tight", facecolor="white")
plt.close()
print(f"Saved: {out_path}")
