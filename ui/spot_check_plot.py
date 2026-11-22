"""
Spot-check validation matrix plot for the SD Data Descriptor.
Shows 30 cases × 12 fields with Correct/Incorrect/Uncertain/N/A colour coding.
"""
import tkinter as tk
from tkinter import ttk

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.patches as mpatches
import numpy as np

# --- embedded spot-check data ---
FIELD_NAMES = [
    "Country", "Building\nname", "City", "Construction\nyear",
    "Demolition\nyear", "Restoration\nyear", "Condition",
    "Spread\nmethod", "Quantity", "Function", "Designer",
    "Coordinates",
]

FIELD_ACC = [
    100, 96.7, 90, 76.7,
    None, 100, 100,
    100, 100, 100, 90,
    86.7,
]

# (label, density_group, [12 judgments])
# judgment: 0=Correct, 1=Incorrect, -1=N/A, -2=Uncertain
JUDGMENTS = [
    # --- High density (18 cases) ---
    ("Klenteng Sam Po Kong", "High",
     [0,0,0,0,0,0,0,0,0,0,0,0]),
    ("Peng Tao Gong Temple", "High",
     [0,0,1,1,-1,0,0,0,0,0,0,1]),
    ("Wat Ratchaorotsaram Ratchaworawihan", "High",
     [0,0,0,0,-1,0,0,0,0,0,0,0]),
    ("Lhong 1919", "High",
     [0,0,0,0,-1,0,0,0,0,0,0,0]),
    ("Chinese Garden, Tsurumi Ryokuchi Park", "High",
     [0,0,0,0,-1,-1,0,0,0,0,0,0]),
    ("Tekigaisō", "High",
     [0,0,0,0,-1,0,0,0,0,0,1,0]),
    ("Shinshuen Garden (Shenxiu Garden)", "High",
     [0,0,1,0,-1,-1,0,0,0,0,0,0]),
    ("Five-story Pagoda of Itsukushima Shrine", "High",
     [0,0,0,0,-1,0,0,0,0,0,0,0]),
    ("Yin Yu Tang", "High",
     [0,0,0,0,0,0,0,0,0,0,0,0]),
    ("Wat Kalayanamitr", "High",
     [0,0,0,0,-1,0,0,0,0,0,0,0]),
    ("Kadampa Temple for World Peace at Kadampa Meditation Center Washington", "High",
     [0,0,0,1,-1,-1,0,0,0,0,0,0]),
    ("Fa Kai Kuan Yin Temple", "High",
     [0,1,1,0,-1,0,0,0,0,0,0,1]),
    ("China Cultural Center in Bangkok", "High",
     [0,0,0,0,-1,-1,0,0,0,0,1,1]),
    ("Fook Tet Soo Khek Temple", "High",
     [0,0,0,0,-1,0,0,0,0,0,0,0]),
    ("Quang Trieu (Cantonese) Assembly Hall", "High",
     [0,0,0,1,-1,0,0,0,0,0,0,0]),
    ("Wat Benchamabophit", "High",
     [0,0,0,0,-1,0,0,0,0,0,0,0]),
    ("Ban Hin Kiong Klenteng", "High",
     [0,0,0,0,-1,0,0,0,0,0,0,0]),
    ("Wat Ben Huan, Thailand", "High",
     [0,0,0,1,-1,-1,0,0,0,0,0,1]),
    # --- Medium density (7 cases) ---
    ("Fishing Temple Virginia Water", "Medium",
     [0,0,0,0,0,-1,0,0,0,0,0,0]),
    ("The Great Pagoda, Kew", "Medium",
     [0,0,0,0,-1,0,0,0,0,0,0,0]),
    ("Yangtorp Sanctuary", "Medium",
     [0,0,0,0,-1,0,0,0,0,0,0,0]),
    ("La Pagode Chinoise de Romorantin-Lanthenay", "Medium",
     [0,0,0,1,-1,0,0,0,0,0,1,0]),
    ("Gold Coast Chinatown", "Medium",
     [0,0,0,0,-1,-1,0,0,0,0,0,0]),
    ("The Pagoda, Blackheath", "Medium",
     [0,0,0,0,-1,0,0,0,0,0,0,0]),
    ("Sze Yup Kwan Ti Temple", "Medium",
     [0,0,0,0,-1,0,0,0,0,0,0,0]),
    # --- Low density (5 cases) ---
    ("Golden Triangle Chinatown, Laos", "Low",
     [0,0,0,1,-1,-1,0,0,0,0,0,0]),
    ("Dunedin Chinese Garden (Lan Yuan)", "Low",
     [0,0,0,0,-1,-1,0,0,0,0,0,0]),
    ("Paro Taktsang", "Low",
     [0,0,0,0,-1,0,0,0,0,0,0,0]),
    ("Takht-e Soleyman", "Low",
     [0,0,0,0,-1,0,0,0,0,0,0,0]),
    ("Bagaya Kyaung", "Low",
     [0,0,0,1,-1,0,0,0,0,0,0,0]),
]

# Pastel colours
CLR_CORRECT   = "#c0e6c8"   # soft green
CLR_INCORRECT = "#f2c4c4"   # soft pink
CLR_NA        = "#e8e8e8"   # light grey
CLR_UNCERTAIN = "#fff2b8"   # soft yellow
CLR_TEXT       = "#333333"
CLR_GROUP_BG   = {"High": "#f5faf7", "Medium": "#fafaf5", "Low": "#faf7f5"}
GROUP_ORDER    = ["High", "Medium", "Low"]


def _clr(j):
    if j == 0:   return CLR_CORRECT
    if j == 1:   return CLR_INCORRECT
    if j == -2:  return CLR_UNCERTAIN
    return CLR_NA


def draw_spot_check_matrix(parent_frame):
    """Draw the spot-check matrix onto *parent_frame* (a tk Frame)."""
    for w in parent_frame.winfo_children():
        w.destroy()

    n_cases = len(JUDGMENTS)
    n_fields = len(FIELD_NAMES)

    fig = Figure(figsize=(14.5, 8.5), facecolor="white")
    gs = fig.add_gridspec(1, 1, left=0.32, right=0.90, top=0.96, bottom=0.12)
    ax = fig.add_subplot(gs[0, 0])

    # Build group position map
    group_positions = {}
    for i, (label, grp, _) in enumerate(JUDGMENTS):
        group_positions.setdefault(grp, []).append(i)

    # Group background shading (drawn first, behind cells)
    y_cursor = n_cases
    for grp in GROUP_ORDER:
        indices = group_positions.get(grp, [])
        n = len(indices)
        if n == 0:
            continue
        y0 = y_cursor - n
        rect = mpatches.Rectangle(
            (-0.5, y0), n_fields + 0.5, n,
            fc=CLR_GROUP_BG[grp], ec="none", zorder=-1, alpha=0.5,
        )
        ax.add_patch(rect)
        y_cursor = y0

    # Draw cells
    for row_i, (label, grp, vals) in enumerate(JUDGMENTS):
        for col_j, j in enumerate(vals):
            color = _clr(j)
            rect = mpatches.FancyBboxPatch(
                (col_j + 0.12, n_cases - 1 - row_i + 0.12),
                0.76, 0.76,
                boxstyle="round,pad=0.04", fc=color, ec="white", lw=0.6,
            )
            ax.add_patch(rect)

    # Section separators (in the gap between rows, not overlapping cells)
    prev = None
    for i, (_, grp, _) in enumerate(JUDGMENTS):
        if grp != prev and prev is not None:
            ax.axhline(y=n_cases - i, color="#999999", lw=1.0, linestyle="--")
        prev = grp

    # Y-axis: case labels (English only, no case IDs)
    y_labels = [label for label, _, _ in JUDGMENTS]
    ax.set_yticks(range(n_cases))
    ax.set_yticklabels(reversed(y_labels), fontsize=6.5, color=CLR_TEXT)
    ax.set_ylim(-0.5, n_cases - 0.5)

    # X-axis: field names
    ax.set_xticks(np.arange(n_fields) + 0.5)
    ax.set_xticklabels(FIELD_NAMES, fontsize=7.5, color=CLR_TEXT)
    ax.set_xlim(0, n_fields)

    # Column accuracy annotations at top
    for j, acc in enumerate(FIELD_ACC):
        if acc is not None:
            ax.text(j + 0.5, n_cases + 0.5, f"{acc:.0f}%", fontsize=7,
                    color="#555", ha="center", va="bottom")
        else:
            ax.text(j + 0.5, n_cases + 0.5, "N/A", fontsize=6.5,
                    color="#aaa", ha="center", va="bottom")

    ax.tick_params(bottom=False, left=False)

    # Group labels on the right
    y_cursor = n_cases - 0.5
    for grp in GROUP_ORDER:
        indices = group_positions.get(grp, [])
        n = len(indices)
        if n == 0:
            continue
        mid = y_cursor - n / 2
        ax.text(n_fields + 0.4, mid, grp, fontsize=8, fontweight="bold",
                color="#888", ha="left", va="center",
                transform=ax.transData)
        y_cursor -= n

    # Legend below matrix (wide, no overlap)
    legend_elements = [
        mpatches.Patch(facecolor=CLR_CORRECT,   edgecolor="white", label="Correct"),
        mpatches.Patch(facecolor=CLR_INCORRECT, edgecolor="white", label="Incorrect"),
        mpatches.Patch(facecolor=CLR_UNCERTAIN, edgecolor="white", label="Uncertain"),
        mpatches.Patch(facecolor=CLR_NA,        edgecolor="white", label="N/A"),
    ]
    ax.legend(
        handles=legend_elements,
        loc="upper center",
        bbox_to_anchor=(0.5, -0.10),
        ncol=4,
        fontsize=8.5,
        frameon=False,
        handletextpad=0.8,
        columnspacing=1.4,
    )

    for spine in ax.spines.values():
        spine.set_visible(False)

    canvas = FigureCanvasTkAgg(fig, parent_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True)

    parent_frame._spot_check_canvas = canvas
    return fig
