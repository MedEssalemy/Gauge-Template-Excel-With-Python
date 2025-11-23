import matplotlib.pyplot as plt
import numpy as np


# CONFIGURATION

# Visual Configuration
FIGURE_SIZE = (16, 9)
FIGURE_DPI = 300

# Section donut (thin, lower layer)
SECTION_INNER_RADIUS = 0.68
SECTION_OUTER_RADIUS = 0.73
SECTION_EDGE_COLOR = "white"
SECTION_EDGE_WIDTH = 3

# Progress donut (thick, upper layer)
PROGRESS_INNER_RADIUS = 0.75
PROGRESS_OUTER_RADIUS = 1.0
PROGRESS_EDGE_COLOR = "white"
PROGRESS_EDGE_WIDTH = 0

# Performance Configuration
SEGMENT_RESOLUTION = 200

# Value Label Configuration
VALUE_LABEL_RADIUS = 0.1
VALUE_LABEL_FONT_SIZE = 45
VALUE_LABEL_COLOR = '#393939'
VALUE_LABEL_FONT_WEIGHT = 'bold'
VALUE_LABEL_FORMAT = xl("labelFormat")

# Progress background
PROGRESS_BACKGROUND_COLOR = '#E0E0E0'

# Gauge Angles (semicircular)
GAUGE_ANGLE_START = 180  # Left side
GAUGE_ANGLE_END = 0      # Right side


# DATA LOADING


# Load gauge value
gauge_value = float(xl("gaugeValue"))

# Load sections data from Excel table
sections_data = xl("gaugeSections", headers=False)
sections_rows = sections_data.values

# Parse sections
sections = []
for row in sections_rows:
    section_desc = str(row[0])
    section_span = float(row[1])
    section_color = str(row[2]).strip()
    sections.append({
        'description': section_desc,
        'span': section_span,
        'color': section_color
    })

# Calculate total range - starting from 0
total_span = sum(s['span'] for s in sections)
gauge_min = 0.0
gauge_max = total_span


# HELPER FUNCTIONS


def value_to_angle(val):
    """Convert value to angle in degrees."""
    val = max(gauge_min, min(val, gauge_max))
    normalized = (val - gauge_min) / (gauge_max - gauge_min)
    # Fixed: progress from left (180°) to right (0°)
    angle = GAUGE_ANGLE_START - (normalized * (GAUGE_ANGLE_START - GAUGE_ANGLE_END))
    return angle


def get_section_color_for_value(val):
    """Get the color of the section where the value falls."""
    cumulative = gauge_min
    
    for section in sections:
        section_end = cumulative + section['span']
        if cumulative <= val <= section_end:
            return section['color']
        cumulative = section_end
    
    return sections[-1]['color']

def format_number(num, fmt="decimal1"):
    """
    Number formatter for gauge progress values.
    Optimized for displaying metrics, KPIs, and progress indicators.
    """
    
    # Load decimal separator from Excel
    decimal_sep_tag = str(xl("decimalSep"))
    if decimal_sep_tag == "dot":
        decimal_sep = "."
        thousand_sep = ","
    else:
        decimal_sep = ","
        thousand_sep = " "
    
    # Helper function to apply separators
    def apply_separators(value_str):
        """Apply thousand and decimal separators to formatted string"""
        if '.' in value_str:
            int_part, dec_part = value_str.split('.')
            int_part = f"{int(int_part):,}".replace(',', thousand_sep)
            return f"{int_part}{decimal_sep}{dec_part}"
        else:
            return f"{int(float(value_str)):,}".replace(',', thousand_sep)
    
    # === INTEGER FORMAT ===
    if fmt == "integer":
        return f"{int(round(num)):,}".replace(',', thousand_sep)
    
    # === DECIMAL FORMATS ===
    if fmt == "decimal1":
        return apply_separators(f"{num:.1f}")
    
    if fmt == "decimal2":
        return apply_separators(f"{num:.2f}")
    
    # === PERCENTAGE FORMATS ===
    if fmt == "percent0":
        return f"{num:.0f}%"
    
    if fmt == "percent1":
        return f"{num:.1f}%".replace('.', decimal_sep)
    
    if fmt == "percent2":
        return f"{num:.2f}%".replace('.', decimal_sep)
    
    # === ABBREVIATED FORMATS ===
    if fmt == "abbreviated":
        if abs(num) >= 1_000_000_000:
            return f"{num/1_000_000_000:.1f}B".replace('.', decimal_sep)
        elif abs(num) >= 1_000_000:
            return f"{num/1_000_000:.1f}M".replace('.', decimal_sep)
        elif abs(num) >= 1_000:
            return f"{num/1_000:.1f}K".replace('.', decimal_sep)
        else:
            return apply_separators(f"{num:.1f}")
    
    if fmt == "thousands":
        return f"{num/1_000:.1f}K".replace('.', decimal_sep)
    
    if fmt == "millions":
        return f"{num/1_000_000:.1f}M".replace('.', decimal_sep)
    
    # === CURRENCY FORMATS ===
    if fmt == "currency_usd":
        formatted = apply_separators(f"{num:.0f}")
        return f"${formatted}" if decimal_sep == "." else f"{formatted} $"
    
    if fmt == "currency_euro":
        formatted = apply_separators(f"{num:.0f}")
        return f"€{formatted}" if decimal_sep == "." else f"{formatted} €"
    
    if fmt == "currency_mad":
        formatted = apply_separators(f"{num:.0f}")
        return f"{formatted} DH"
    
    # === DEFAULT FALLBACK ===
    return apply_separators(f"{num:.1f}")

# PLOT SETUP

fig, ax = plt.subplots(figsize=FIGURE_SIZE, subplot_kw={'projection': 'polar'})
fig.set_dpi(FIGURE_DPI)
fig.patch.set_facecolor('white')
ax.set_facecolor('white')

# Configure polar plot
ax.set_theta_zero_location("E")
ax.set_theta_direction(1)
ax.set_ylim(0, 1.3)

# Remove all default elements
ax.set_yticks([])
ax.set_xticks([])
ax.spines['polar'].set_visible(False)
ax.grid(False)


# DRAW SECTION DONUT (THIN, REFERENCE LAYER)

cumulative_value = gauge_min

for section in sections:
    start_value = cumulative_value
    end_value = cumulative_value + section['span']
    
    start_angle = value_to_angle(start_value)
    end_angle = value_to_angle(end_value)
    
    # Convert to radians (swap start and end to draw correctly)
    start_rad = np.radians(start_angle)
    end_rad = np.radians(end_angle)
    
    # Create arc - linspace from end to start for correct direction
    theta = np.linspace(end_rad, start_rad, SEGMENT_RESOLUTION)
    
    ax.fill_between(
        theta,
        SECTION_INNER_RADIUS,
        SECTION_OUTER_RADIUS,
        color=section['color'],
        edgecolor=SECTION_EDGE_COLOR,
        linewidth=SECTION_EDGE_WIDTH,
        zorder=5
    )
    
    cumulative_value = end_value

# DRAW PROGRESS DONUT (THICK, UPPER LAYER)

# First draw the full background in light gray
full_start_rad = np.radians(GAUGE_ANGLE_START)
full_end_rad = np.radians(GAUGE_ANGLE_END)
full_theta = np.linspace(full_end_rad, full_start_rad, SEGMENT_RESOLUTION)

ax.fill_between(
    full_theta,
    PROGRESS_INNER_RADIUS,
    PROGRESS_OUTER_RADIUS,
    color=PROGRESS_BACKGROUND_COLOR,
    edgecolor=PROGRESS_EDGE_COLOR,
    linewidth=PROGRESS_EDGE_WIDTH,
    zorder=10
)

# Calculate progress percentage and get corresponding color
progress_value = max(gauge_min, min(gauge_value, gauge_max))
progress_color = get_section_color_for_value(progress_value)

# Draw progress arc from start to current value (on top of gray background)
start_angle = GAUGE_ANGLE_START
end_angle = value_to_angle(progress_value)

start_rad = np.radians(start_angle)
end_rad = np.radians(end_angle)

# Draw from end to start for correct direction
theta = np.linspace(end_rad, start_rad, SEGMENT_RESOLUTION)

ax.fill_between(
    theta,
    PROGRESS_INNER_RADIUS,
    PROGRESS_OUTER_RADIUS,
    color=progress_color,
    edgecolor=PROGRESS_EDGE_COLOR,
    linewidth=PROGRESS_EDGE_WIDTH,
    zorder=11
)

# DRAW VALUE LABEL
ax.text(
    np.radians(90),
    VALUE_LABEL_RADIUS,
    format_number(gauge_value, VALUE_LABEL_FORMAT),
    ha='center',
    va='center',
    fontsize=VALUE_LABEL_FONT_SIZE,
    fontweight=VALUE_LABEL_FONT_WEIGHT,
    color=VALUE_LABEL_COLOR,
    zorder=12
)

# DISPLAY
plt.tight_layout()
plt.show()