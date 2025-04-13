# Paint Retail Analytics

A Python-based business analytics application designed for paint retail businesses. This application allows users to analyze sales data, inventory, and customer information through an intuitive graphical interface.

## Features

- Import data from Excel (.xlsx, .xls) or CSV files
- Analyze sales trends and patterns
- Track product performance
- Monitor color popularity
- Analyze brand performance
- Interactive visualizations using Plotly
- Flexible date range filtering

## Required Data Format

Your spreadsheet should include the following columns:
- Date: Transaction date
- Product Name: Name of the paint product
- Category: Product category (e.g., interior, exterior)
- Brand: Paint brand
- Color: Paint color
- Quantity Sold: Number of units sold
- Unit Price: Price per unit
- Total Revenue: Total sale amount
- Cost Price: Cost per unit
- Store Location: Store identifier

## Installation

1. Create a virtual environment:
```bash
python -m venv venv
```

2. Activate the virtual environment:
```bash
# On Windows
venv\Scripts\activate
```

3. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the application:
```bash
python paint_analytics.py
```

2. Click "Upload Spreadsheet" to import your data
3. Select analysis options:
   - Choose date range
   - Select analysis type
4. Click "Run Analysis" to generate insights
5. View results in the application window and interactive charts in your web browser

## Support

For any issues or questions, please open an issue in the repository.
