"""
Automated Sales Dashboard Generator

Use Case: Generate a comprehensive monthly sales dashboard from raw data
Author: Excel Automation Pipeline
Date: 2025-12-17

Takes raw sales data and creates a multi-sheet dashboard with:
- Executive summary
- Product performance
- Regional analysis
- Trends over time
- Top customers
"""

import pandas as pd
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path


def create_sample_sales_data():
    """Create sample sales data for demonstration"""

    import random

    # Generate 3 months of sales data
    start_date = datetime.now() - timedelta(days=90)
    dates = pd.date_range(start=start_date, periods=500, freq='H')

    products = ['Widget A', 'Widget B', 'Gadget Pro', 'Tool Set', 'Parts Kit',
                'Premium Package', 'Starter Kit', 'Deluxe Edition']
    regions = ['North', 'South', 'East', 'West', 'Central']
    sales_reps = ['Alice', 'Bob', 'Charlie', 'Diana', 'Edward', 'Fiona']
    customers = [f'Customer_{i:03d}' for i in range(1, 51)]

    data = {
        'Transaction_ID': [f'TXN{i:06d}' for i in range(1, 501)],
        'Date': [d.date() for d in dates],
        'Product': [random.choice(products) for _ in range(500)],
        'Region': [random.choice(regions) for _ in range(500)],
        'Sales_Rep': [random.choice(sales_reps) for _ in range(500)],
        'Customer': [random.choice(customers) for _ in range(500)],
        'Units': [random.randint(1, 20) for _ in range(500)],
        'Unit_Price': [round(random.uniform(10, 200), 2) for _ in range(500)],
        'Discount_%': [random.choice([0, 5, 10, 15, 20]) for _ in range(500)]
    }

    df = pd.DataFrame(data)
    df['Subtotal'] = df['Units'] * df['Unit_Price']
    df['Discount_Amount'] = df['Subtotal'] * (df['Discount_%'] / 100)
    df['Total_Sale'] = df['Subtotal'] - df['Discount_Amount']

    df.to_excel('raw_sales_data.xlsx', index=False)
    print("✓ Created sample raw_sales_data.xlsx")
    return df


def load_sales_data(filename='raw_sales_data.xlsx'):
    """Load sales data from Excel"""

    if not Path(filename).exists():
        print(f"⚠️  {filename} not found. Creating sample data...")
        return create_sample_sales_data()

    print(f"📂 Loading {filename}...")
    df = pd.read_excel(filename)
    df['Date'] = pd.to_datetime(df['Date'])
    print(f"✓ Loaded {len(df)} transactions")
    return df


def generate_executive_summary(df):
    """Generate high-level executive summary"""

    # Key metrics
    total_revenue = df['Total_Sale'].sum()
    total_transactions = len(df)
    avg_transaction = df['Total_Sale'].mean()
    total_units = df['Units'].sum()

    # Date range
    date_range = f"{df['Date'].min().strftime('%Y-%m-%d')} to {df['Date'].max().strftime('%Y-%m-%d')}"

    # Growth metrics (compare first half vs second half)
    mid_date = df['Date'].min() + (df['Date'].max() - df['Date'].min()) / 2
    first_half = df[df['Date'] < mid_date]['Total_Sale'].sum()
    second_half = df[df['Date'] >= mid_date]['Total_Sale'].sum()
    growth_rate = ((second_half - first_half) / first_half * 100) if first_half > 0 else 0

    summary = {
        'Metric': [
            'Reporting Period',
            'Total Revenue',
            'Total Transactions',
            'Average Transaction Value',
            'Total Units Sold',
            'Revenue Growth Rate',
            'Unique Products Sold',
            'Active Regions',
            'Active Sales Reps',
            'Unique Customers'
        ],
        'Value': [
            date_range,
            f'${total_revenue:,.2f}',
            f'{total_transactions:,}',
            f'${avg_transaction:,.2f}',
            f'{total_units:,}',
            f'{growth_rate:+.1f}%',
            df['Product'].nunique(),
            df['Region'].nunique(),
            df['Sales_Rep'].nunique(),
            df['Customer'].nunique()
        ]
    }

    return pd.DataFrame(summary)


def analyze_product_performance(df):
    """Analyze performance by product"""

    product_perf = df.groupby('Product').agg({
        'Transaction_ID': 'count',
        'Units': 'sum',
        'Total_Sale': 'sum',
        'Discount_Amount': 'sum'
    }).reset_index()

    product_perf.columns = ['Product', 'Transactions', 'Units_Sold', 'Revenue', 'Total_Discounts']
    product_perf['Avg_Transaction'] = product_perf['Revenue'] / product_perf['Transactions']
    product_perf['Discount_%'] = (
                product_perf['Total_Discounts'] / (product_perf['Revenue'] + product_perf['Total_Discounts']) * 100)

    # Add revenue share
    product_perf['Revenue_Share_%'] = (product_perf['Revenue'] / product_perf['Revenue'].sum() * 100)

    # Sort by revenue
    product_perf = product_perf.sort_values('Revenue', ascending=False)

    # Format currency columns
    for col in ['Revenue', 'Total_Discounts', 'Avg_Transaction']:
        product_perf[col] = product_perf[col].round(2)

    return product_perf


def analyze_regional_performance(df):
    """Analyze performance by region"""

    regional = df.groupby('Region').agg({
        'Transaction_ID': 'count',
        'Units': 'sum',
        'Total_Sale': 'sum'
    }).reset_index()

    regional.columns = ['Region', 'Transactions', 'Units_Sold', 'Revenue']
    regional['Avg_Transaction'] = regional['Revenue'] / regional['Transactions']
    regional['Revenue_Share_%'] = (regional['Revenue'] / regional['Revenue'].sum() * 100)

    regional = regional.sort_values('Revenue', ascending=False)

    return regional


def analyze_sales_rep_performance(df):
    """Analyze performance by sales representative"""

    rep_perf = df.groupby('Sales_Rep').agg({
        'Transaction_ID': 'count',
        'Total_Sale': 'sum',
        'Customer': 'nunique'
    }).reset_index()

    rep_perf.columns = ['Sales_Rep', 'Transactions', 'Revenue', 'Unique_Customers']
    rep_perf['Avg_Transaction'] = rep_perf['Revenue'] / rep_perf['Transactions']
    rep_perf['Revenue_per_Customer'] = rep_perf['Revenue'] / rep_perf['Unique_Customers']

    rep_perf = rep_perf.sort_values('Revenue', ascending=False)

    return rep_perf


def analyze_time_trends(df):
    """Analyze trends over time"""

    # Daily trends
    daily = df.groupby('Date').agg({
        'Transaction_ID': 'count',
        'Total_Sale': 'sum'
    }).reset_index()
    daily.columns = ['Date', 'Transactions', 'Revenue']

    # Weekly trends
    df['Week'] = df['Date'].dt.to_period('W')
    weekly = df.groupby('Week').agg({
        'Transaction_ID': 'count',
        'Total_Sale': 'sum'
    }).reset_index()
    weekly.columns = ['Week', 'Transactions', 'Revenue']
    weekly['Week'] = weekly['Week'].astype(str)

    # Monthly trends
    df['Month'] = df['Date'].dt.to_period('M')
    monthly = df.groupby('Month').agg({
        'Transaction_ID': 'count',
        'Total_Sale': 'sum',
        'Units': 'sum'
    }).reset_index()
    monthly.columns = ['Month', 'Transactions', 'Revenue', 'Units']
    monthly['Month'] = monthly['Month'].astype(str)

    return daily, weekly, monthly


def identify_top_customers(df, top_n=20):
    """Identify top customers by revenue"""

    customer_analysis = df.groupby('Customer').agg({
        'Transaction_ID': 'count',
        'Total_Sale': 'sum',
        'Units': 'sum'
    }).reset_index()

    customer_analysis.columns = ['Customer', 'Transactions', 'Total_Revenue', 'Total_Units']
    customer_analysis['Avg_Transaction'] = customer_analysis['Total_Revenue'] / customer_analysis['Transactions']

    customer_analysis = customer_analysis.sort_values('Total_Revenue', ascending=False).head(top_n)

    return customer_analysis


def generate_dashboard(df, output_file='sales_dashboard.xlsx'):
    """Generate comprehensive sales dashboard"""

    print("\n" + "=" * 60)
    print("GENERATING SALES DASHBOARD")
    print("=" * 60)

    print("\n📊 Analyzing data...")

    # Generate all analyses
    exec_summary = generate_executive_summary(df)
    print("  ✓ Executive summary")

    product_perf = analyze_product_performance(df)
    print("  ✓ Product performance")

    regional = analyze_regional_performance(df)
    print("  ✓ Regional analysis")

    rep_perf = analyze_sales_rep_performance(df)
    print("  ✓ Sales rep performance")

    daily, weekly, monthly = analyze_time_trends(df)
    print("  ✓ Time trends")

    top_customers = identify_top_customers(df)
    print("  ✓ Customer analysis")

    # Create dashboard Excel file
    print(f"\n💾 Writing dashboard to '{output_file}'...")

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

        # Sheet 1: Executive Summary
        exec_summary.to_excel(writer, sheet_name='Executive_Summary', index=False)

        # Sheet 2: Product Performance
        product_perf.to_excel(writer, sheet_name='Product_Performance', index=False)

        # Sheet 3: Regional Analysis
        regional.to_excel(writer, sheet_name='Regional_Analysis', index=False)

        # Sheet 4: Sales Rep Performance
        rep_perf.to_excel(writer, sheet_name='Sales_Rep_Performance', index=False)

        # Sheet 5: Top Customers
        top_customers.to_excel(writer, sheet_name='Top_Customers', index=False)

        # Sheet 6-8: Time Trends
        monthly.to_excel(writer, sheet_name='Monthly_Trends', index=False)
        weekly.to_excel(writer, sheet_name='Weekly_Trends', index=False)
        daily.to_excel(writer, sheet_name='Daily_Trends', index=False)

    print(f"✓ Dashboard saved with 8 sheets")

    # Print key insights
    print("\n" + "=" * 60)
    print("KEY INSIGHTS")
    print("=" * 60)

    total_revenue = df['Total_Sale'].sum()
    print(f"\n💰 Total Revenue: ${total_revenue:,.2f}")
    print(f"📊 Total Transactions: {len(df):,}")
    print(f"📦 Total Units Sold: {df['Units'].sum():,}")

    print(f"\n🏆 Top 3 Products by Revenue:")
    for i, row in product_perf.head(3).iterrows():
        print(f"   {i + 1}. {row['Product']}: ${row['Revenue']:,.2f} ({row['Revenue_Share_%']:.1f}%)")

    print(f"\n🌍 Top 3 Regions by Revenue:")
    for i, row in regional.head(3).iterrows():
        print(f"   {i + 1}. {row['Region']}: ${row['Revenue']:,.2f} ({row['Revenue_Share_%']:.1f}%)")

    print(f"\n⭐ Top Sales Rep:")
    top_rep = rep_perf.iloc[0]
    print(f"   {top_rep['Sales_Rep']}: ${top_rep['Revenue']:,.2f} from {top_rep['Transactions']} transactions")


def main():
    """Main execution"""

    print("=" * 60)
    print("AUTOMATED SALES DASHBOARD GENERATOR")
    print("=" * 60)

    # Load data
    df = load_sales_data()

    # Generate dashboard
    generate_dashboard(df)

    print("\n" + "=" * 60)
    print("✓ DASHBOARD GENERATION COMPLETE")
    print("=" * 60)
    print("\n📊 Open 'sales_dashboard.xlsx' to view your dashboard!")


if __name__ == "__main__":
    main()