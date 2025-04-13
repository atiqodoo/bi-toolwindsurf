import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import plotly.express as px
import plotly.io as pio
from PIL import Image, ImageTk
import webbrowser
import os
from datetime import datetime
import plotly.graph_objects as go

class PaintAnalyticsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Paint Retail Analytics Dashboard")
        self.root.geometry("1200x800")
        self.root.configure(bg='#f0f2f5')
        
        # Configure ttk style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('Dashboard.TFrame', background='#f0f2f5')
        self.style.configure('Card.TFrame', background='white', relief='solid')
        self.style.configure('Header.TLabel', 
                           background='#1a73e8', 
                           foreground='white', 
                           font=('Helvetica', 12, 'bold'),
                           padding=10)
        self.style.configure('Title.TLabel',
                           font=('Helvetica', 14, 'bold'),
                           background='white',
                           padding=5)
        self.style.configure('Metric.TLabel',
                           font=('Helvetica', 20, 'bold'),
                           background='white',
                           padding=5)
        
        # Create main container
        self.main_container = ttk.Frame(root, style='Dashboard.TFrame')
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Create header
        self.create_header()
        
        # Create dashboard layout
        self.create_dashboard_layout()
        
        # Initialize data
        self.df = None
        
    def create_header(self):
        """Create the dashboard header with controls."""
        header = ttk.Frame(self.main_container, style='Dashboard.TFrame')
        header.pack(fill=tk.X, pady=(0, 20))
        
        # Left side - Title and Upload
        left_header = ttk.Frame(header, style='Dashboard.TFrame')
        left_header.pack(side=tk.LEFT)
        
        title = ttk.Label(left_header, 
                         text="Paint Retail Analytics Dashboard",
                         style='Header.TLabel')
        title.pack(side=tk.LEFT, padx=5)
        
        upload_btn = ttk.Button(left_header, 
                              text=" Upload Data",
                              command=self.load_file)
        upload_btn.pack(side=tk.LEFT, padx=10)
        
        # Right side - Analysis Controls
        right_header = ttk.Frame(header, style='Dashboard.TFrame')
        right_header.pack(side=tk.RIGHT)
        
        # Date Range
        date_frame = ttk.Frame(right_header, style='Dashboard.TFrame')
        date_frame.pack(side=tk.LEFT, padx=10)
        
        ttk.Label(date_frame, text="Date Range:").pack(side=tk.LEFT)
        self.start_date = ttk.Entry(date_frame, width=10)
        self.start_date.pack(side=tk.LEFT, padx=5)
        ttk.Label(date_frame, text="to").pack(side=tk.LEFT, padx=2)
        self.end_date = ttk.Entry(date_frame, width=10)
        self.end_date.pack(side=tk.LEFT, padx=5)
        
        # Analysis Type
        analysis_frame = ttk.Frame(right_header, style='Dashboard.TFrame')
        analysis_frame.pack(side=tk.LEFT, padx=10)
        
        ttk.Label(analysis_frame, text="View:").pack(side=tk.LEFT)
        self.analysis_var = tk.StringVar(value="Sales Overview")
        analysis_menu = ttk.OptionMenu(analysis_frame, 
                                     self.analysis_var, 
                                     "Sales Overview",
                                     *self.get_analysis_options(),
                                     command=self.refresh_analysis)
        analysis_menu.pack(side=tk.LEFT, padx=5)
        
        # Refresh button
        refresh_btn = ttk.Button(right_header,
                               text=" Refresh",
                               command=self.refresh_analysis)
        refresh_btn.pack(side=tk.LEFT, padx=10)
        
    def create_dashboard_layout(self):
        """Create the main dashboard layout."""
        # Create left and right panels
        self.left_panel = ttk.Frame(self.main_container, style='Dashboard.TFrame')
        self.left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        self.right_panel = ttk.Frame(self.main_container, style='Dashboard.TFrame')
        self.right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))
        
        # Create metric cards
        self.create_metric_cards()
        
        # Create charts area
        self.create_charts_area()
        
        # Create details area
        self.create_details_area()
        
    def create_metric_cards(self):
        """Create metric cards for key performance indicators."""
        metrics_frame = ttk.Frame(self.left_panel, style='Dashboard.TFrame')
        metrics_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Create metric cards
        self.metric_cards = {}
        metrics = [
            ("Total Revenue", "", "#34a853"),
            ("Total Profit", "", "#4285f4"),
            ("Units Sold", "", "#fbbc05"),
            ("Profit Margin", "", "#ea4335")
        ]
        
        for i, (metric, icon, color) in enumerate(metrics):
            card = ttk.Frame(metrics_frame, style='Card.TFrame')
            card.grid(row=0, column=i, padx=5, sticky='nsew')
            
            ttk.Label(card, text=f"{icon} {metric}", style='Title.TLabel').pack(pady=5)
            value_label = ttk.Label(card, text="--", style='Metric.TLabel')
            value_label.pack(pady=5)
            
            self.metric_cards[metric] = value_label
            
        metrics_frame.grid_columnconfigure((0,1,2,3), weight=1)
        
    def create_charts_area(self):
        """Create area for charts and graphs."""
        charts_frame = ttk.Frame(self.left_panel, style='Card.TFrame')
        charts_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(charts_frame, text="Performance Trends", style='Title.TLabel').pack(pady=10)
        
        # Create tabs for different charts
        self.charts_notebook = ttk.Notebook(charts_frame)
        self.charts_notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Revenue Trend tab
        revenue_frame = ttk.Frame(self.charts_notebook)
        self.charts_notebook.add(revenue_frame, text="Revenue Trend")
        
        # Product Performance tab
        product_frame = ttk.Frame(self.charts_notebook)
        self.charts_notebook.add(product_frame, text="Product Performance")
        
        # Department Analysis tab
        dept_frame = ttk.Frame(self.charts_notebook)
        self.charts_notebook.add(dept_frame, text="Department Analysis")
        
    def create_details_area(self):
        """Create area for detailed analysis."""
        details_frame = ttk.Frame(self.right_panel, style='Card.TFrame')
        details_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(details_frame, text="Detailed Analysis", style='Title.TLabel').pack(pady=10)
        
        # Create text widget for details
        self.result_text = tk.Text(details_frame, wrap=tk.WORD, height=30)
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(details_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.config(yscrollcommand=scrollbar.set)
        
    def update_metrics(self, metrics):
        """Update the metric cards with new values."""
        if metrics:
            self.metric_cards["Total Revenue"].config(text=self.format_currency(metrics['Total Revenue']))
            self.metric_cards["Total Profit"].config(text=self.format_currency(metrics['Total Profit']))
            self.metric_cards["Units Sold"].config(text=f"{int(metrics['Total Units Sold']):,}")
            self.metric_cards["Profit Margin"].config(text=self.format_percent(metrics['Profit Margin %']))
    
    def create_trend_chart(self, df):
        """Create and display trend chart."""
        if 'Date' not in df.columns:
            return
            
        # Convert date and group by month
        df['Month'] = pd.to_datetime(df['Date']).dt.to_period('M')
        monthly = df.groupby('Month').agg({
            'Net Sales': 'sum',
            'Cost of Sale': 'sum'
        }).reset_index()
        
        # Calculate profit
        monthly['Profit'] = monthly['Net Sales'] - monthly['Cost of Sale']
        
        # Create figure
        fig = go.Figure()
        
        # Add traces
        fig.add_trace(go.Scatter(
            x=monthly['Month'].astype(str),
            y=monthly['Net Sales'],
            name='Revenue',
            line=dict(color='#4285f4', width=2)
        ))
        
        fig.add_trace(go.Scatter(
            x=monthly['Month'].astype(str),
            y=monthly['Profit'],
            name='Profit',
            line=dict(color='#34a853', width=2)
        ))
        
        # Update layout
        fig.update_layout(
            title='Monthly Revenue and Profit Trends',
            xaxis_title='Month',
            yaxis_title='Amount',
            template='plotly_white',
            height=400,
            margin=dict(l=40, r=40, t=40, b=40)
        )
        
        # Save and display
        fig.write_html("trend_chart.html")
        webbrowser.open("trend_chart.html")
        
    def load_file(self):
        try:
            file_path = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")]
            )
            
            if file_path:
                print(f"\nAttempting to load file: {file_path}")
                
                if file_path.lower().endswith(('.xlsx', '.xls')):
                    try:
                        # First try reading with no data conversion
                        print("Loading Excel file (initial read)...")
                        raw_df = pd.read_excel(file_path, engine='openpyxl')
                        
                        print("\nInitial data read successful")
                        print(f"Shape: {raw_df.shape}")
                        print("\nColumns found:", raw_df.columns.tolist())
                        
                        # Show sample of raw data
                        print("\nFirst few rows of raw data:")
                        print(raw_df.head())
                        
                        # Now try to convert numeric columns
                        print("\nAttempting numeric conversion...")
                        self.df = raw_df.copy()
                        
                        numeric_cols = ['Qty', 'CP incl VAT', 'SP incl VAT', 'Cost of Sale', 
                                      'Discounts', 'Net Sales', 'Nt. Sl. Ls Vt']
                        
                        for col in numeric_cols:
                            if col in self.df.columns:
                                print(f"\nProcessing column: {col}")
                                print("Original values (first 5):", self.df[col].head().tolist())
                                print("Data type:", self.df[col].dtype)
                                
                                try:
                                    # Try direct numeric conversion first
                                    self.df[col] = pd.to_numeric(self.df[col], errors='coerce')
                                    print("Converted values:", self.df[col].head().tolist())
                                    print(f"Sum: {self.df[col].sum()}")
                                except Exception as conv_err:
                                    print(f"Direct conversion failed: {str(conv_err)}")
                                    
                                    # Try cleaning and converting
                                    try:
                                        # Convert to string and clean
                                        cleaned = self.df[col].astype(str)
                                        cleaned = cleaned.str.replace('£', '', regex=False)
                                        cleaned = cleaned.str.replace('$', '', regex=False)
                                        cleaned = cleaned.str.replace(',', '', regex=False)
                                        cleaned = cleaned.str.replace(' ', '', regex=False)
                                        cleaned = cleaned.str.strip()
                                        
                                        print("Cleaned values:", cleaned.head().tolist())
                                        
                                        # Convert to numeric
                                        self.df[col] = pd.to_numeric(cleaned, errors='coerce')
                                        print("Final converted values:", self.df[col].head().tolist())
                                        print(f"Sum: {self.df[col].sum()}")
                                    except Exception as clean_err:
                                        print(f"Cleaning conversion failed: {str(clean_err)}")
                            else:
                                print(f"Warning: Column {col} not found")
                        
                        # Show final data info
                        print("\nFinal DataFrame Info:")
                        print(self.df.info())
                        
                    except Exception as excel_err:
                        print(f"Excel load error: {str(excel_err)}")
                        raise ValueError(f"Could not read Excel file. Error: {str(excel_err)}")
                else:
                    raise ValueError("Please use an Excel file (.xlsx or .xls)")
                
                # Clear previous results
                self.result_text.delete(1.0, tk.END)
                
                # Show data preview
                self.result_text.insert(tk.END, "Data Preview\n")
                self.result_text.insert(tk.END, "=" * 50 + "\n\n")
                self.result_text.insert(tk.END, f"Loaded {len(self.df)} rows and {len(self.df.columns)} columns\n\n")
                self.result_text.insert(tk.END, "Column Details:\n\n")
                
                for col in self.df.columns:
                    self.result_text.insert(tk.END, f"Column: {col}\n")
                    self.result_text.insert(tk.END, f"  Type: {self.df[col].dtype}\n")
                    self.result_text.insert(tk.END, f"  Non-null values: {self.df[col].count()}\n")
                    self.result_text.insert(tk.END, f"  Null values: {self.df[col].isna().sum()}\n")
                    
                    if col in numeric_cols:
                        self.result_text.insert(tk.END, f"  Sum: {self.df[col].sum()}\n")
                        
                    sample_vals = self.df[col].head(3).tolist()
                    self.result_text.insert(tk.END, f"  Sample values: {sample_vals}\n\n")
                
                # Initialize date fields
                date_col = self.get_date_column(self.df)
                if date_col:
                    try:
                        dates = pd.to_datetime(self.df[date_col])
                        min_date = dates.min().strftime('%Y-%m-%d')
                        max_date = dates.max().strftime('%Y-%m-%d')
                        self.start_date.delete(0, tk.END)
                        self.start_date.insert(0, min_date)
                        self.end_date.delete(0, tk.END)
                        self.end_date.insert(0, max_date)
                        print(f"\nSet date range: {min_date} to {max_date}")
                    except Exception as e:
                        print(f"Error setting date range: {str(e)}")
                
                # Run initial analysis
                self.run_analysis()
                
        except Exception as e:
            error_msg = f"Failed to load file: {str(e)}"
            print(f"Error: {error_msg}")
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, f"Error: {error_msg}\n\n")
            if hasattr(self, 'df') and self.df is not None:
                self.result_text.insert(tk.END, "\nPartial data info:\n")
                self.result_text.insert(tk.END, f"Shape: {self.df.shape}\n")
                self.result_text.insert(tk.END, "Columns: " + ", ".join(self.df.columns.tolist()))
            messagebox.showerror("Error", error_msg)
            
    def run_analysis(self):
        try:
            if self.df is None:
                raise ValueError("Please load a data file first")
            
            print("\nRunning analysis...")
            
            # Filter data by date if needed
            filtered_df = self.filter_data_by_date()
            
            # Get selected analysis type
            analysis_type = self.analysis_var.get()
            
            # Run the analysis
            self.analyze_data(filtered_df, analysis_type)
            
            # Create and show trend chart
            self.create_trend_chart(filtered_df)
            
        except Exception as e:
            error_msg = f"Analysis failed: {str(e)}"
            print(f"Error: {error_msg}")
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, f"Error: {error_msg}\n\n")
            if self.df is not None:
                self.result_text.insert(tk.END, "Available columns:\n")
                for col in self.df.columns:
                    self.result_text.insert(tk.END, f"- {col}\n")
            messagebox.showerror("Error", error_msg)
            
    def parse_date(self, date_series):
        """Try multiple date formats to parse the date column."""
        date_formats = [
            # ISO format
            '%Y-%m-%d',
            # British format
            '%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y',
            # American format
            '%m/%d/%Y', '%m-%d-%Y', '%m.%d.%Y',
            # Month name formats
            '%d-%b-%Y', '%d %b %Y', '%d-%B-%Y', '%d %B %Y',
            '%b-%d-%Y', '%b %d %Y', '%B-%d-%Y', '%B %d %Y',
            # Two digit years
            '%d/%m/%y', '%m/%d/%y',
            # With time
            '%Y-%m-%d %H:%M:%S', '%d/%m/%Y %H:%M:%S', '%m/%d/%Y %H:%M:%S'
        ]

        # First try pandas default parsing
        try:
            return pd.to_datetime(date_series, errors='raise')
        except:
            pass

        # Try each format
        for date_format in date_formats:
            try:
                return pd.to_datetime(date_series, format=date_format, errors='raise')
            except:
                continue

        # If no format works, try a more flexible parser
        try:
            return pd.to_datetime(date_series, format='mixed', errors='raise')
        except Exception as e:
            raise ValueError(f"Could not parse dates. Please ensure dates are in a standard format. Error: {str(e)}")

    def filter_data_by_date(self):
        # Print columns for debugging
        print("Available columns:", self.df.columns.tolist())
        
        # Find the date column (case-insensitive and strip whitespace)
        date_columns = [col for col in self.df.columns 
                       if str(col).lower().strip() == 'date']
        
        if not date_columns:
            # Try to find any column containing date-like values
            for col in self.df.columns:
                try:
                    # Try parsing first value
                    self.parse_date(pd.Series([self.df[col].iloc[0]]))
                    date_columns = [col]
                    print(f"Found date column: {col}")
                    break
                except:
                    continue
        
        if not date_columns:
            messagebox.showwarning("Warning", 
                f"Date column not found. Available columns: {', '.join(self.df.columns)}")
            return self.df
            
        date_col = date_columns[0]
        print(f"Using column '{date_col}' as date column")
        
        try:
            # Convert dates using flexible parser
            self.df[date_col] = self.parse_date(self.df[date_col])
            date_filter = self.start_date.get() + " to " + self.end_date.get()
            
            if date_filter == "  to ":
                return self.df
            
            start_date = pd.to_datetime(self.start_date.get())
            end_date = pd.to_datetime(self.end_date.get())
                
            return self.df[(self.df[date_col] >= start_date) & (self.df[date_col] <= end_date)]
        except Exception as e:
            messagebox.showerror("Error", 
                f"Error processing date column '{date_col}': {str(e)}\n"
                f"Sample values: {', '.join(map(str, self.df[date_col].head().tolist()))}")
            return self.df
        
    def calculate_financial_metrics(self, df):
        """Calculate key financial metrics."""
        try:
            print("\nCalculating financial metrics...")
            print("\nOriginal DataFrame Info:")
            print(df.info())
            
            # First, make a copy to avoid modifying original
            work_df = df.copy()
            
            # Define column mappings
            qty_col = 'Qty'
            revenue_col = 'Net Sales'
            cost_col = 'Cost of Sale'
            
            # Print raw data samples
            print("\nRaw data samples:")
            for col in [qty_col, revenue_col, cost_col]:
                if col in work_df.columns:
                    print(f"\n{col}:")
                    print("First 5 values:", work_df[col].head().tolist())
                    print("Data type:", work_df[col].dtype)
            
            # Clean numeric data
            def clean_numeric_column(df, col_name):
                if col_name not in df.columns:
                    print(f"Warning: Column {col_name} not found")
                    return
                
                print(f"\nCleaning {col_name}:")
                try:
                    # Convert to string first
                    df[col_name] = df[col_name].astype(str)
                    print("After string conversion:", df[col_name].head().tolist())
                    
                    # Remove any currency symbols, commas, and spaces
                    df[col_name] = df[col_name].str.replace('£', '', regex=False)
                    df[col_name] = df[col_name].str.replace('$', '', regex=False)
                    df[col_name] = df[col_name].str.replace(',', '', regex=False)
                    df[col_name] = df[col_name].str.strip()
                    print("After cleaning:", df[col_name].head().tolist())
                    
                    # Convert to numeric
                    df[col_name] = pd.to_numeric(df[col_name], errors='coerce')
                    print("After numeric conversion:", df[col_name].head().tolist())
                    print("Sum:", df[col_name].sum())
                    print("Non-null count:", df[col_name].count())
                    
                except Exception as e:
                    print(f"Error cleaning {col_name}: {str(e)}")
                    raise
            
            # Clean all numeric columns
            for col in [qty_col, revenue_col, cost_col]:
                clean_numeric_column(work_df, col)
            
            # Calculate totals
            total_quantity = work_df[qty_col].sum()
            total_revenue = work_df[revenue_col].sum()
            total_cost = work_df[cost_col].sum()
            
            print("\nCalculated totals:")
            print(f"Total quantity: {total_quantity}")
            print(f"Total revenue: {total_revenue}")
            print(f"Total cost: {total_cost}")
            
            if total_quantity == 0 or total_revenue == 0 or total_cost == 0:
                error_msg = "One or more totals are zero. Details:\n"
                error_msg += f"Quantity total: {total_quantity}\n"
                error_msg += f"Revenue total: {total_revenue}\n"
                error_msg += f"Cost total: {total_cost}\n\n"
                
                error_msg += "Column details:\n"
                for col in [qty_col, revenue_col, cost_col]:
                    if col in work_df.columns:
                        error_msg += f"\n{col}:\n"
                        error_msg += f"  Type: {work_df[col].dtype}\n"
                        error_msg += f"  Non-null count: {work_df[col].count()}\n"
                        error_msg += f"  Sample values (first 5): {work_df[col].head().tolist()}\n"
                        error_msg += f"  Sum: {work_df[col].sum()}\n"
                
                raise ValueError(error_msg)
            
            # Calculate metrics
            total_profit = total_revenue - total_cost
            profit_margin = (total_profit / total_revenue * 100) if total_revenue > 0 else 0
            markup = ((total_revenue - total_cost) / total_cost * 100) if total_cost > 0 else 0
            avg_profit_per_unit = total_profit / total_quantity if total_quantity > 0 else 0
            
            metrics = {
                'Total Revenue': total_revenue,
                'Total Cost': total_cost,
                'Total Profit': total_profit,
                'Total Units Sold': total_quantity,
                'Profit Margin (%)': profit_margin,
                'Average Profit per Unit': avg_profit_per_unit,
                'Markup (%)': markup
            }
            
            print("\nFinal metrics:")
            for key, value in metrics.items():
                print(f"{key}: {value}")
            
            return metrics
            
        except Exception as e:
            print(f"\nError in calculate_financial_metrics: {str(e)}")
            raise ValueError(f"Failed to calculate metrics: {str(e)}")

    def format_currency(self, value):
        """Format number as currency."""
        return f"${value:,.2f}"

    def format_percent(self, value):
        """Format number as percentage."""
        return f"{value:.1f}%"

    def analyze_sales(self, df):
        try:
            # Print available columns for debugging
            print("\nAnalyzing sales data...")
            print("Available columns:", [col.strip() for col in df.columns])
            
            # Calculate financial metrics
            metrics = self.calculate_financial_metrics(df)
            
            # Display results
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, "Financial Analysis\n")
            self.result_text.insert(tk.END, "=" * 50 + "\n\n")
            
            # Revenue and Profit
            self.result_text.insert(tk.END, "Revenue & Profit:\n")
            self.result_text.insert(tk.END, f"Total Revenue: {self.format_currency(metrics['Total Revenue'])}\n")
            self.result_text.insert(tk.END, f"Total Cost: {self.format_currency(metrics['Total Cost'])}\n")
            self.result_text.insert(tk.END, f"Total Profit: {self.format_currency(metrics['Total Profit'])}\n")
            
            # Margins and Markup
            self.result_text.insert(tk.END, "\nProfitability Metrics:\n")
            self.result_text.insert(tk.END, f"Profit Margin: {self.format_percent(metrics['Profit Margin (%)'])}\n")
            self.result_text.insert(tk.END, f"Markup: {self.format_percent(metrics['Markup (%)'])}\n")
            
            # Per Unit Analysis
            self.result_text.insert(tk.END, "\nPer Unit Analysis:\n")
            self.result_text.insert(tk.END, f"Total Units Sold: {metrics['Total Units Sold']:,}\n")
            self.result_text.insert(tk.END, f"Average Unit Price: {self.format_currency(metrics['Average Unit Price'])}\n")
            self.result_text.insert(tk.END, f"Average Cost Price: {self.format_currency(metrics['Average Cost Price'])}\n")
            self.result_text.insert(tk.END, f"Average Profit per Unit: {self.format_currency(metrics['Average Profit per Unit'])}\n")
            
            self.update_metrics(metrics)
            
        except Exception as e:
            import traceback
            error_msg = str(e)
            error_details = traceback.format_exc()
            print(f"Error details:\n{error_details}")
            
            # Show detailed error in the result text
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, "Error in Analysis\n")
            self.result_text.insert(tk.END, "=" * 50 + "\n\n")
            self.result_text.insert(tk.END, "The following error occurred:\n\n")
            self.result_text.insert(tk.END, f"{error_msg}\n\n")
            self.result_text.insert(tk.END, "Available Columns in Your Data:\n")
            for col in df.columns:
                self.result_text.insert(tk.END, f"- {col}\n")
                # Show sample data for debugging
                try:
                    sample = df[col].head()
                    self.result_text.insert(tk.END, f"  Sample values: {', '.join(map(str, sample))}\n")
                except:
                    pass
            
            messagebox.showerror("Error", "Analysis failed. Check the main window for details.")

    def analyze_products(self, df):
        try:
            # Find required columns (case-insensitive)
            required_columns = {
                'product': next((col for col in df.columns if 'product description' in col.lower().strip()), None),
                'department': next((col for col in df.columns if 'department' in col.lower().strip()), None),
                'quantity': next((col for col in df.columns if col.lower().strip() == 'qty'), None),
                'revenue': next((col for col in df.columns if col.lower().strip() in ['net sales', 'nt. sl. ls vt']), None)
            }
            
            # Check for missing columns
            missing = [name for name, col in required_columns.items() if col is None]
            if missing:
                raise ValueError(f"Missing columns for product analysis: {', '.join(missing)}")
            
            # Group by product and calculate metrics
            product_metrics = df.groupby(required_columns['product']).agg({
                required_columns['quantity']: 'sum',
                required_columns['revenue']: 'sum'
            }).reset_index()
            
            # Sort by revenue
            product_metrics = product_metrics.sort_values(by=required_columns['revenue'], ascending=False)
            
            # Display results
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, "Top Products by Revenue\n")
            self.result_text.insert(tk.END, "=" * 50 + "\n\n")
            
            # Display top 10 products
            for _, row in product_metrics.head(10).iterrows():
                self.result_text.insert(tk.END, f"Product: {row[required_columns['product']]}\n")
                self.result_text.insert(tk.END, f"Total Revenue: {self.format_currency(row[required_columns['revenue']])}\n")
                self.result_text.insert(tk.END, f"Units Sold: {int(row[required_columns['quantity']]):,}\n")
                self.result_text.insert(tk.END, "-" * 50 + "\n")
            
            # Department Analysis
            if required_columns['department']:
                dept_metrics = df.groupby(required_columns['department']).agg({
                    required_columns['quantity']: 'sum',
                    required_columns['revenue']: 'sum'
                }).reset_index()
                
                dept_metrics = dept_metrics.sort_values(by=required_columns['revenue'], ascending=False)
                
                self.result_text.insert(tk.END, "\nDepartment Performance\n")
                self.result_text.insert(tk.END, "=" * 50 + "\n\n")
                
                for _, row in dept_metrics.iterrows():
                    self.result_text.insert(tk.END, f"Department: {row[required_columns['department']]}\n")
                    self.result_text.insert(tk.END, f"Total Revenue: {self.format_currency(row[required_columns['revenue']])}\n")
                    self.result_text.insert(tk.END, f"Units Sold: {int(row[required_columns['quantity']]):,}\n")
                    self.result_text.insert(tk.END, "-" * 50 + "\n")
            
        except Exception as e:
            print(f"Error in analyze_products: {str(e)}")
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, "Error in Product Analysis\n")
            self.result_text.insert(tk.END, "=" * 50 + "\n\n")
            self.result_text.insert(tk.END, f"Error: {str(e)}\n\n")
            self.result_text.insert(tk.END, "Available columns:\n")
            for col in df.columns:
                self.result_text.insert(tk.END, f"- {col}\n")
            messagebox.showwarning("Warning", str(e))

    def get_analysis_options(self):
        """Return available analysis options."""
        return [
            "Sales Overview",
            "Product Analysis",
            "Department Performance"
        ]

    def analyze_data(self, df, analysis_type):
        """Analyze the data based on the selected analysis type."""
        try:
            if df.empty:
                raise ValueError("No data available for analysis")

            print("\nStarting analysis...")
            print(f"Analysis type: {analysis_type}")
            print(f"Data shape: {df.shape}")
            
            # Clear previous results
            self.result_text.delete(1.0, tk.END)
            
            # Calculate financial metrics
            try:
                metrics = self.calculate_financial_metrics(df)
                
                # Display metrics
                self.result_text.insert(tk.END, "Financial Metrics:\n")
                self.result_text.insert(tk.END, "=" * 50 + "\n\n")
                
                for key, value in metrics.items():
                    print(f"{key}: {value}")  # Debug print
                    self.result_text.insert(tk.END, f"{key}: {value:,.2f}\n")
                
                # Check if any metrics are zero
                zero_metrics = [k for k, v in metrics.items() if v == 0]
                if zero_metrics:
                    self.result_text.insert(tk.END, "\nWarning: The following metrics are zero:\n")
                    for metric in zero_metrics:
                        self.result_text.insert(tk.END, f"- {metric}\n")
                    
                    # Show column information for troubleshooting
                    self.result_text.insert(tk.END, "\nColumn Information for Debugging:\n")
                    self.result_text.insert(tk.END, "=" * 50 + "\n\n")
                    
                    numeric_cols = ['Qty', 'CP incl VAT', 'SP incl VAT', 'Cost of Sale', 
                                  'Discounts', 'Net Sales', 'Nt. Sl. Ls Vt']
                    
                    for col in numeric_cols:
                        if col in df.columns:
                            self.result_text.insert(tk.END, f"\n{col}:\n")
                            self.result_text.insert(tk.END, f"  Type: {df[col].dtype}\n")
                            self.result_text.insert(tk.END, f"  Non-null count: {df[col].count()}\n")
                            self.result_text.insert(tk.END, f"  Sum: {df[col].sum()}\n")
                            self.result_text.insert(tk.END, f"  Sample values: {df[col].head().tolist()}\n")
                            
                            # Check for string values that should be numeric
                            if df[col].dtype == 'object':
                                sample_strings = df[col].head().astype(str).tolist()
                                self.result_text.insert(tk.END, f"  Sample strings: {sample_strings}\n")
                        else:
                            self.result_text.insert(tk.END, f"\nWarning: Column '{col}' not found in data\n")
                    
                    raise ValueError("All totals are zero. See details above for debugging information.")
                
            except Exception as calc_error:
                print(f"Error in financial calculations: {str(calc_error)}")
                self.result_text.insert(tk.END, f"\nError in financial calculations: {str(calc_error)}\n")
                self.result_text.insert(tk.END, "\nAvailable columns:\n")
                for col in df.columns:
                    self.result_text.insert(tk.END, f"- {col}\n")
                    if col in df.columns:
                        sample = df[col].head()
                        self.result_text.insert(tk.END, f"  Sample values: {sample.tolist()}\n")
                raise calc_error
            
        except Exception as e:
            print(f"Analysis error: {str(e)}")
            error_msg = f"\nThe following error occurred:\n\n{str(e)}\n\n"
            error_msg += "Available Columns in Your Data:\n"
            
            for col in df.columns:
                error_msg += f"- {col}\n"
                if col in df.columns:
                    dtype = df[col].dtype
                    non_null = df[col].count()
                    sample = df[col].head(3)
                    error_msg += f"  Type: {dtype}\n"
                    error_msg += f"  Non-null count: {non_null}\n"
                    error_msg += f"  Sample values: {sample.tolist()}\n"
            
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, error_msg)
            raise ValueError("Analysis failed. Check the main window for details.")
    
    def save_and_show_plot(self, fig):
        # Save plot as HTML and open in browser
        fig.write_html("temp_chart.html")
        webbrowser.open("file://" + os.path.realpath("temp_chart.html"))

    def get_date_column(self, df):
        date_columns = [col for col in df.columns 
                       if str(col).lower().strip() == 'date']
        
        if not date_columns:
            # Try to find any column containing date-like values
            for col in df.columns:
                try:
                    # Try parsing first value
                    self.parse_date(pd.Series([df[col].iloc[0]]))
                    date_columns = [col]
                    print(f"Found date column: {col}")
                    break
                except:
                    continue
        
        if not date_columns:
            messagebox.showwarning("Warning", 
                f"Date column not found. Available columns: {', '.join(df.columns)}")
            return None
            
        return date_columns[0]

    def refresh_analysis(self, value=None):
        self.run_analysis()

def main():
    root = tk.Tk()
    app = PaintAnalyticsApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
