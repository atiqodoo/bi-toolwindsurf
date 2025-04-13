import pandas as pd
import plotly.express as px
import plotly.io as pio
from datetime import datetime

# Read the sample data
df = pd.read_excel('sample_paint_sales.xlsx')

# 1. Overall Performance
print("=== Overall Performance ===")
total_revenue = df['Total Revenue'].sum()
total_profit = df['Profit'].sum()
total_units = df['Quantity Sold'].sum()
profit_margin = (total_profit / total_revenue) * 100

print(f"Total Revenue: ${total_revenue:,.2f}")
print(f"Total Profit: ${total_profit:,.2f}")
print(f"Total Units Sold: {total_units:,}")
print(f"Overall Profit Margin: {profit_margin:.1f}%")

# 2. Top Products
print("\n=== Top 5 Products by Revenue ===")
top_products = df.groupby('Product Name').agg({
    'Total Revenue': 'sum',
    'Quantity Sold': 'sum',
    'Profit': 'sum'
}).sort_values('Total Revenue', ascending=False).head(5)

print(top_products)

# Create and save top products chart
fig_products = px.bar(top_products.reset_index(), 
                     x='Product Name', 
                     y='Total Revenue',
                     title="Top 5 Products by Revenue")
fig_products.write_html("top_products.html")

# 3. Color Analysis
print("\n=== Color Popularity ===")
color_analysis = df.groupby('Color').agg({
    'Total Revenue': 'sum',
    'Quantity Sold': 'sum'
}).sort_values('Quantity Sold', ascending=False)

print(color_analysis)

# Create and save color analysis chart
fig_colors = px.pie(color_analysis.reset_index(), 
                   values='Quantity Sold', 
                   names='Color',
                   title="Sales Distribution by Color")
fig_colors.write_html("color_analysis.html")

# 4. Monthly Trends
df['Month'] = pd.to_datetime(df['Date']).dt.to_period('M')
monthly_sales = df.groupby('Month').agg({
    'Total Revenue': 'sum',
    'Profit': 'sum'
}).reset_index()
monthly_sales['Month'] = monthly_sales['Month'].astype(str)

print("\n=== Monthly Performance ===")
print(monthly_sales)

# Create and save monthly trends chart
fig_trends = px.line(monthly_sales, 
                    x='Month', 
                    y=['Total Revenue', 'Profit'],
                    title="Monthly Revenue and Profit Trends")
fig_trends.write_html("monthly_trends.html")

# 5. Brand Performance
print("\n=== Brand Performance ===")
brand_analysis = df.groupby('Brand').agg({
    'Total Revenue': 'sum',
    'Profit': 'sum',
    'Quantity Sold': 'sum'
}).sort_values('Total Revenue', ascending=False)
brand_analysis['Profit Margin'] = (brand_analysis['Profit'] / brand_analysis['Total Revenue']) * 100

print(brand_analysis)

# Create and save brand analysis chart
fig_brands = px.bar(brand_analysis.reset_index(),
                   x='Brand',
                   y=['Total Revenue', 'Profit'],
                   title="Brand Performance Analysis")
fig_brands.write_html("brand_analysis.html")

print("\nAnalysis completed! Open the HTML files to view interactive charts.")
