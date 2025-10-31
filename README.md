# Force Tester Data Visualization

A Streamlit application for visualizing force tester data from Excel files.

## Features

- Upload Excel files with paired x-y data columns
- Automatic grouping of similar test configurations
- Fully customizable charts with Plotly
- Interactive visualizations
- Export charts as HTML

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
streamlit run app.py
```

## Usage

1. Upload an Excel file containing force tester data
2. Each test sample should have two adjacent columns (x data, y data)
3. Customize the chart using the sidebar controls:
   - Chart titles and axis labels
   - Line types (solid, dash, dashdot)
   - Line colors per configuration group
   - Legend names
   - Scale min/max values

## Data Format

Your Excel file should have data organized as column pairs:
- Column 1: X data (e.g., Travel in mm) for Sample 1
- Column 2: Y data (e.g., Force in N) for Sample 1
- Column 3: X data for Sample 2
- Column 4: Y data for Sample 2
- And so on...

Samples with similar names (e.g., `syringe-30G-water`, `syringe-30G-water-1`) 
will be automatically grouped together for consistent styling.

# plots
