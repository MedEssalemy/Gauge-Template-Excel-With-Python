# Excel Gauge Charts with Python

This repository contains Python scripts for creating professional gauge charts that integrate seamlessly with Microsoft Excel. These visualizations are perfect for dashboards, KPI tracking, and performance monitoring.

## 🎥 Quick Start Video

**New to this template? Watch the 1-minute demo:**

[![Excel Gauge Charts Template - 1 Minute Demo](https://img.youtube.com/vi/YOUR_VIDEO_ID/maxresdefault.jpg)](https://www.youtube.com/watch?v=YOUR_VIDEO_ID)

*Click above to see both gauge types in action and learn the key features*

---

## 📊 Available Gauge Types

1. **Semicircular Progress Gauge** (`semicircular_gauge.py`) - Modern donut-style gauge with colored sections
2. **Needle Gauge with Milestones** (`needle_gauge.py`) - Classic needle gauge with milestone markers

## 🔧 Supported Excel Versions

- **Excel 2021**
- **Microsoft 365** (Excel Online and Desktop)
- **Excel for Web**

> **Note**: These scripts use Excel's Python integration feature (available in Microsoft 365 and Excel for Web). Excel 2021 users may need the Python in Excel add-in.

## 📋 Prerequisites

Before using these gauge charts, ensure:

- You have a compatible Excel version with Python support enabled
- The matplotlib and numpy libraries are available (pre-installed in Excel's Python environment)
- You're familiar with creating Excel named ranges and tables

## 🚀 Getting Started

### Step 1: Download the Files

Download both files from this repository:
- The Python script (`.py` file)
- The Excel template (`.xlsx` file)

> 💡 **First time here?** Watch the [1-minute video tutorial](#-quick-start-video) above to see everything in action!

### Step 2: Open the Excel Template

Open the provided Excel template file. The template includes:
- Pre-configured named ranges
- Sample data tables
- Placeholder cells for your data

### Step 3: Understand the Template Structure

Each gauge type requires specific named ranges and data structures.

---

## 📈 Semicircular Progress Gauge

### Required Named Ranges

You must define these named ranges in Excel (Formulas tab > Define Name):

| Named Range | Type | Description | Example |
|-------------|------|-------------|---------|
| `gaugeValue` | Single Cell | The current value to display | `75.5` |
| `gaugeSections` | Table | Section definitions (3 columns) | See below |
| `labelFormat` | Single Cell | Number format string | `"decimal1"` |
| `decimalSep` | Single Cell | Decimal separator preference | `"dot"` or `"comma"` |

### Setting Up the Sections Table

Create an Excel table named `gaugeSections` with **no headers** and exactly 3 columns:

| Column 1: Description | Column 2: Span | Column 3: Color |
|----------------------|----------------|-----------------|
| Low | 30 | #FF6B6B |
| Medium | 40 | #FFD93D |
| High | 30 | #6BCF7F |

**Important Notes:**
- Sections are drawn from left to right (0° to 180°)
- The gauge automatically calculates the total range from section spans
- Colors must be in hex format (e.g., `#FF6B6B`)
- The gauge value should fall within 0 to total span

### Format Options

The `labelFormat` named range accepts these values:

| Format Code | Display Example | Use Case |
|-------------|-----------------|----------|
| `integer` | 1,234 | Whole numbers |
| `decimal1` | 1,234.5 | One decimal place |
| `decimal2` | 1,234.56 | Two decimal places |
| `percent0` | 75% | Percentage (no decimals) |
| `percent1` | 75.5% | Percentage (1 decimal) |
| `percent2` | 75.52% | Percentage (2 decimals) |
| `abbreviated` | 1.2M | Auto K/M/B abbreviation |
| `thousands` | 1.2K | Always show as thousands |
| `millions` | 1.2M | Always show as millions |
| `currency_usd` | $1,234 | US Dollar |
| `currency_euro` | €1,234 or 1,234 € | Euro |
| `currency_mad` | 1,234 DH | Moroccan Dirham |

### Step-by-Step Setup

1. **Create your data table**:
   - In a worksheet, create a 3-column table without headers
   - Fill in: Description | Span | Color (hex code)

2. **Define named ranges**:
   - Select your data table → Formulas → Define Name → Name it `gaugeSections`
   - Select the cell with your current value → Define Name → `gaugeValue`
   - Select the cell with format code → Define Name → `labelFormat`
   - Select the cell with separator preference → Define Name → `decimalSep`

3. **Insert Python code**:
   - Go to Insert tab → Python
   - Copy and paste the `semicircular_gauge.py` code
   - Run the cell

4. **Customize colors**:
   - Modify the hex color codes in your sections table
   - Re-run the Python cell to see changes

---

## 🎯 Needle Gauge with Milestones

### Required Named Ranges

| Named Range | Type | Description | Example |
|-------------|------|-------------|---------|
| `progressValue` | Single Cell | Current value on the gauge | `245` |
| `gaugeLabel` | Single Cell | Text label displayed in center | `"Q3 Sales\nTarget"` |
| `projectMilestone` | Single Cell | Which milestone to highlight | `"Phase 2"` |
| `milestonesTargets3` | Table | Milestone definitions (2 columns) | See below |

### Setting Up the Milestones Table

Create an Excel table named `milestonesTargets3` with 2 columns:

| Column 1: Milestone Name | Column 2: Score/Value |
|--------------------------|----------------------|
| Phase 1 | 100 |
| Phase 2 | 200 |
| Phase 3 | 300 |
| Complete | 400 |

**Important Notes:**
- The first row can be headers or data (script auto-detects)
- Milestones are automatically sorted by value
- The last milestone sets the maximum gauge value
- The highlighted milestone (red) is matched by name from `projectMilestone`

### Step-by-Step Setup

1. **Create milestones table**:
   - Create a 2-column table with milestone names and values
   - Include headers if desired (script handles both cases)

2. **Define named ranges**:
   - Select milestone table → Formulas → Define Name → `milestonesTargets3`
   - Select current value cell → Define Name → `progressValue`
   - Select label cell → Define Name → `gaugeLabel`
   - Select milestone name to highlight → Define Name → `projectMilestone`

3. **Insert Python code**:
   - Go to Insert tab → Python
   - Copy and paste the `needle_gauge.py` code
   - Run the cell

4. **Customize appearance**:
   - Add `\n` in gaugeLabel for multi-line text
   - Ensure projectMilestone exactly matches a milestone name
   - Adjust milestone values to fit your scale

---

## 💡 Best Practices

### General Recommendations

1. **Use descriptive range names**: Stick to the exact names specified above
2. **Keep data organized**: Place all gauge data on one dedicated worksheet
3. **Test with sample data first**: Use the provided template values before your own data
4. **Document your ranges**: Add comments in Excel to remind yourself what each range does

### Data Quality

- **Semicircular Gauge**: Ensure gaugeValue is between 0 and total section span
- **Needle Gauge**: Ensure progressValue doesn't exceed the highest milestone
- **Both Gauges**: Validate that all named ranges exist before running the script
- **Colors**: Always use 6-digit hex codes (e.g., `#FF0000`, not `#F00`)

### Performance Tips

- Keep section/milestone counts reasonable (3-7 items work best)
- Use higher DPI (300) for presentations, lower (100-150) for quick previews
- If the chart takes too long to render, reduce `SEGMENT_RESOLUTION` in the code

### Troubleshooting

| Issue | Solution |
|-------|----------|
| "Name not found" error | Check that all required named ranges are defined |
| Gauge doesn't display | Verify Python is enabled in your Excel version |
| Wrong colors | Ensure hex codes start with `#` |
| Text overlap | Reduce number of sections/milestones or adjust font sizes |
| Value not updating | Re-run the Python cell after changing data |

---

## 🎨 Customization Guide

Both scripts have configuration sections at the top. You can modify:

### Visual Settings
- `FIGURE_SIZE`: Chart dimensions (default: 16:9 ratio)
- `FIGURE_DPI`: Resolution (higher = sharper, but slower)
- Font sizes, colors, and spacing

### Gauge Geometry
- Inner/outer radius values
- Segment edge widths and colors
- Needle size and shape (needle gauge)

### Colors
- Modify `SEGMENT_COLORS` array (needle gauge)
- Change section colors in your Excel table (semicircular gauge)

**Pro Tip**: Change only one setting at a time to see its effect clearly.

---

## 📁 Repository Structure

```
├── semicircular_gauge.py          # Progress gauge script
├── needle_gauge.py                # Milestone needle gauge script
├── gauge_template.xlsx            # Excel template with sample data
└── README.md                      # This file
```

---

## 🤝 Contributing

Feel free to customize these gauges for your needs! If you create interesting variations, consider sharing them.

---

## 📄 License

These scripts are provided as-is for use in your Excel dashboards and reports.

---

## 🆘 Need Help?

If you encounter issues:
1. Watch the [video tutorial](#-quick-start-video) for a visual walkthrough
2. Verify all named ranges are correctly defined
3. Check that your Excel version supports Python
4. Ensure your data matches the expected format
5. Review the configuration section in the script

**Still stuck?** Open an issue on GitHub with:
- Your Excel version
- Screenshots of your named ranges
- The error message you're seeing

---

## 🌟 Show Your Support

If you find these gauge templates helpful:
- ⭐ Star this repository
- 🔔 Subscribe to the [YouTube channel](https://youtu.be/NSaglFuvu_o) for more Excel tips
- 🐛 Report issues or suggest features
- 📢 Share with colleagues who might benefit

---

**Happy Charting! 📊**

---

### 📺 More Resources

- [Watch the Demo](https://www.youtube.com/watch?v=YOUR_VIDEO_ID) - 1-minute overview
- [Excel Python Documentation](https://support.microsoft.com/en-us/office/get-started-with-python-in-excel-a33fbcbe-065b-41d3-82cf-23d05397f53d) - Official Microsoft guide
- [Report an Issue](../../issues) - Found a bug? Let us know!
