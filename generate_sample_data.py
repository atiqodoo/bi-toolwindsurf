import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# Set random seed for reproducibility
np.random.seed(42)

# Sample data
products = [
    'Premium Interior Matte',
    'Premium Interior Satin',
    'Premium Exterior Flat',
    'Premium Exterior Semi-Gloss',
    'Economy Interior Matte',
    'Economy Exterior Flat',
    'Specialty Chalk Paint',
    'Specialty Metal Paint',
    'Designer Collection Matte',
    'Designer Collection Gloss'
]

brands = ['ColorMaster', 'PaintPro', 'ArtisanHue', 'EcoPaint', 'LuxuryCoat']
colors = ['White', 'Beige', 'Gray', 'Blue', 'Green', 'Red', 'Yellow', 'Brown', 'Black', 'Navy']

# Generate dates for the last year
end_date = datetime.now()
start_date = end_date - timedelta(days=365)
dates = pd.date_range(start=start_date, end=end_date, freq='D')

def format_date(date):
    # Convert numpy datetime64 to pandas Timestamp
    date = pd.Timestamp(date)
    # Randomly choose a date format
    formats = [
        '%Y-%m-%d',           # ISO format: 2025-04-11
        '%d/%m/%Y',           # British: 11/04/2025
        '%m/%d/%Y',           # American: 04/11/2025
        '%d-%b-%Y',           # Text format: 11-Apr-2025
        '%d %B %Y',           # Full month: 11 April 2025
    ]
    return date.strftime(np.random.choice(formats))

# Generate random data
num_records = 1000
random_dates = np.random.choice(dates, num_records)

# Base prices for different product categories
base_prices = {
    'Premium': {'cost': 25, 'price': 45},
    'Economy': {'cost': 15, 'price': 25},
    'Specialty': {'cost': 35, 'price': 65},
    'Designer': {'cost': 40, 'price': 75}
}

def get_product_prices(product_name):
    category = next((cat for cat in base_prices.keys() if cat in product_name), 'Economy')
    base = base_prices[category]
    # Add some random variation
    cost = base['cost'] * (1 + np.random.uniform(-0.1, 0.1))
    price = base['price'] * (1 + np.random.uniform(-0.1, 0.1))
    return cost, price

# Generate the data
data = []
for _ in range(num_records):
    product = np.random.choice(products)
    quantity = np.random.randint(1, 21)
    cost_price, unit_price = get_product_prices(product)
    
    record = {
        'Date': format_date(np.random.choice(dates)),
        'Product Name': product,
        'Category': 'Interior' if 'Interior' in product else 'Exterior' if 'Exterior' in product else 'Specialty',
        'Brand': np.random.choice(brands),
        'Color': np.random.choice(colors),
        'Quantity Sold': quantity,
        'Unit Price': round(unit_price, 2),
        'Cost Price': round(cost_price, 2),
        'Total Revenue': round(quantity * unit_price, 2),
        'Total Cost': round(quantity * cost_price, 2),
        'Profit': round(quantity * (unit_price - cost_price), 2)
    }
    data.append(record)

# Create DataFrame and save to Excel
df = pd.DataFrame(data)
df.to_excel('sample_paint_sales.xlsx', index=False)
print("Sample data saved to sample_paint_sales.xlsx")
