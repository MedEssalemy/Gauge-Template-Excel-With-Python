import matplotlib.pyplot as plt
import numpy as np

# ==============================================================================
# CONFIGURATION
# ==============================================================================

# Visual Configuration
FIGURE_SIZE = (16,9) #16:9 aspect ratio
FIGURE_DPI = 300
GAUGE_INNER_RADIUS = 0.75
GAUGE_OUTER_RADIUS = 1.0
NEEDLE_TIP_RADIUS = 0.87
NEEDLE_PIVOT_RADIUS = 0.05
NEEDLE_COLOR = "#cc0000"
SEGMENT_EDGE_COLOR = "white"
SEGMENT_EDGE_WIDTH = 1

# Performance Configuration
SEGMENT_RESOLUTION = 100
NEEDLE_RESOLUTION = 100

# Label Configuration
LABEL_TEXT_RADIUS = 1.12
LABEL_FONT_SIZE_NORMAL = 8
LABEL_FONT_SIZE_MATCHED = 9
LABEL_LINE_SPACING = 1.15
MATCHED_LABEL_COLOR = "red"
NORMAL_LABEL_COLOR = "black"

# Pointer Configuration
POINTER_OUTER_RADIUS = 1.02
POINTER_INNER_RADIUS = 0.76
POINTER_LINE_WIDTH = 2.2
POINTER_COLOR_MATCHED = "#cc0000"
POINTER_COLOR_NORMAL = "#0D3512"

# Value Label Configuration
VALUE_LABEL_RADIUS = 0.24
VALUE_LABEL_FONT_SIZE = 17
VALUE_LABEL_COLOR = '#153D64'
VALUE_LABEL_PADDING = 0.35
VALUE_LABEL_BG_ALPHA = 0.87

# Color Palette
SEGMENT_COLORS = ["#eaf7e4", "#c6e6b8", "#a3da92", "#80cd6c", "#5dbf46", "#3aa52e"]

# Gauge Range
GAUGE_MIN_SCORE = 0.0
GAUGE_ANGLE_START = -90
GAUGE_ANGLE_END = 90

# ==============================================================================
# DATA LOADING
# ==============================================================================

value = float(xl("progressValue"))
gauge_label = xl("gaugeLabel")
project_milestone = xl("projectMilestone")

milestones_data = xl("milestonesTargets3", headers=False)
rows = milestones_data.values

if isinstance(rows[0][1], str):
    rows = rows[1:]

pairs = sorted([(str(r[0]), float(r[1])) for r in rows], key=lambda x: x[1])
labels, scores = zip(*pairs)
labels = list(labels)
scores = list(scores)

max_score = scores[-1]
score_range = max_score - GAUGE_MIN_SCORE
angle_range = GAUGE_ANGLE_END - GAUGE_ANGLE_START

project_milestone_lower = str(project_milestone).strip().lower()

# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================

def score_to_radians(v):
    """Convert score value to radians for polar plot positioning."""
    v = max(GAUGE_MIN_SCORE, min(v, max_score))
    normalized = (v - GAUGE_MIN_SCORE) / score_range
    degrees = GAUGE_ANGLE_START + (normalized * angle_range)
    return np.radians(degrees)


def format_number(num):
    """Format number as integer if whole, otherwise use comma as decimal separator."""
    return str(int(num)) if num == int(num) else str(num).replace('.', ',')


def get_straight_line_in_polar(start_theta, start_r, end_theta, end_r, points):
    """Generate straight line in Cartesian space and convert to polar coordinates."""
    x1 = start_r * np.cos(start_theta - np.pi/2)
    y1 = start_r * np.sin(start_theta - np.pi/2)
    x2 = end_r * np.cos(end_theta - np.pi/2)
    y2 = end_r * np.sin(end_theta - np.pi/2)
    
    x_vals = np.linspace(x1, x2, points)
    y_vals = np.linspace(y1, y2, points)
    
    r_vals = np.sqrt(x_vals**2 + y_vals**2)
    theta_vals = np.arctan2(y_vals, x_vals) + np.pi/2
    
    return theta_vals, r_vals


# ==============================================================================
# PLOT SETUP
# ==============================================================================

fig, ax = plt.subplots(figsize=FIGURE_SIZE, subplot_kw={'projection': 'polar'})
fig.set_dpi(FIGURE_DPI)
fig.patch.set_facecolor('none')

ax.set_theta_zero_location("N")
ax.set_theta_direction(-1)
ax.set_thetamin(GAUGE_ANGLE_START)
ax.set_thetamax(GAUGE_ANGLE_END)
ax.set_ylim(0, 1.8)

ax.set_yticks([])
ax.set_xticks([])
ax.spines['polar'].set_visible(False)
ax.grid(False)
ax.set_frame_on(False)
ax.set_facecolor('none')
ax.axis('off')

# ==============================================================================
# DRAW GAUGE SEGMENTS
# ==============================================================================

segment_boundaries = [GAUGE_MIN_SCORE] + scores
num_segments = len(segment_boundaries) - 1

boundaries_rad = [score_to_radians(val) for val in segment_boundaries]

for i in range(num_segments):
    start_rad = boundaries_rad[i]
    end_rad = boundaries_rad[i + 1]
    
    theta = np.linspace(start_rad, end_rad, SEGMENT_RESOLUTION)
    color = SEGMENT_COLORS[i % len(SEGMENT_COLORS)]
    
    ax.fill_between(
        theta, 
        GAUGE_INNER_RADIUS, 
        GAUGE_OUTER_RADIUS,
        color=color,
        edgecolor=SEGMENT_EDGE_COLOR,
        linewidth=SEGMENT_EDGE_WIDTH
    )

# ==============================================================================
# DRAW MILESTONE LABELS WITH POINTERS
# ==============================================================================

num_labels = len(labels)

for idx in range(num_labels - 1):
    lbl = labels[idx]
    ang_rad = boundaries_rad[idx + 1]
    
    is_matched = (str(lbl).strip().lower() == project_milestone_lower)
    
    pointer_color = POINTER_COLOR_MATCHED if is_matched else POINTER_COLOR_NORMAL
    ax.plot(
        [ang_rad, ang_rad],
        [POINTER_OUTER_RADIUS, POINTER_INNER_RADIUS],
        color=pointer_color,
        linewidth=POINTER_LINE_WIDTH,
        zorder=8
    )
    
    full_text = f"{lbl}\n{format_number(scores[idx])}"
    ax.text(
        ang_rad, 
        LABEL_TEXT_RADIUS, 
        full_text,
        fontsize=LABEL_FONT_SIZE_MATCHED if is_matched else LABEL_FONT_SIZE_NORMAL,
        color=MATCHED_LABEL_COLOR if is_matched else NORMAL_LABEL_COLOR,
        ha="center",
        va="center",
        fontweight="bold",
        linespacing=LABEL_LINE_SPACING
    )

# ==============================================================================
# DRAW NEEDLE
# ==============================================================================

needle_angle_rad = score_to_radians(value)

needle_theta, needle_r = get_straight_line_in_polar(
    start_theta=0,
    start_r=NEEDLE_PIVOT_RADIUS,
    end_theta=needle_angle_rad,
    end_r=NEEDLE_TIP_RADIUS,
    points=NEEDLE_RESOLUTION
)

ax.plot(needle_theta, needle_r, color=NEEDLE_COLOR, lw=1.8, zorder=10)

ax.scatter(
    0, 
    NEEDLE_PIVOT_RADIUS, 
    s=100, 
    color=NEEDLE_COLOR, 
    zorder=11,
    edgecolor='white',
    linewidth=0
)

# ==============================================================================
# DRAW VALUE LABEL
# ==============================================================================

ax.text(
    0, 
    VALUE_LABEL_RADIUS,
    gauge_label,
    ha='center',
    va='center',
    fontsize=VALUE_LABEL_FONT_SIZE,
    fontweight='bold',
    color=VALUE_LABEL_COLOR,
    zorder=12,
    linespacing=1.3,
    bbox=dict(
        boxstyle=f"round,pad={VALUE_LABEL_PADDING}",
        facecolor="white",
        edgecolor="none",
        alpha=VALUE_LABEL_BG_ALPHA
    )
)

# ==============================================================================
# DISPLAY
# ==============================================================================

plt.subplots_adjust(0, 0, 1, 1)
plt.show()