"""
Basic Multi-File Consolidation Example

Use Case: Combine multiple monthly sales files into a single consolidated report
Author: Excel Automation Pipeline
Date: 2025-12-17

This is the most common use case - you receive multiple Excel files
(monthly exports, regional reports, etc.) and need to combine them into one.
"""

import pandas as pd
from pathlib import Path
from datetime import datetime


def consolidate_sales_files(file_pattern='sales_*.xlsx', output_file='consolidated_sales.xlsx'):
    """
    Consolidate multiple Excel files matching a pattern into one file.

    Args:
        file_pattern: Pattern to match files (e.g., 'sales_*.xlsx')
        output_file: Name of output consolidated file
    """

    print("=" * 60)
    print("EXCEL FILE CONSOLIDATION")
    print("=" * 60)

    # Step 1: Find all matching files
    print(f"\n[1] Searching for files matching '{file_pattern}'...")
    files = list(Path('.').glob(file_pattern))

    if not files:
        print(f"❌ No files found matching '{file_pattern}'")
        print("\nCreating sample files for demonstration...")
        create_sample_files()
        files = list(Path('.').glob(file_pattern))

    print(f"✓ Found {len(files)} files:")
    for f in files:
        print(f"  - {f.name}")

    # Step 2: Read all files
    print(f"\n[2] Reading all Excel files...")
    dataframes = []

    for file in files:
        try:
            df = pd.read_excel(file)
            # Add source file column to track origin
            df['Source_File'] = file.name
            dataframes.append(df)
            print(f"  ✓ {file.name}: {len(df)} rows")
        except Exception as e:
            print(f"  ❌ Error reading {file.name}: {e}")

    if not dataframes:
        print("❌ No data to consolidate!")
        return

    # Step 3: Combine all DataFrames
    print(f"\n[3] Consolidating data...")
    consolidated = pd.concat(dataframes, ignore_index=True)
    print(f"✓ Total rows: {len(consolidated)}")

    # Step 4: Basic data quality checks
    print(f"\n[4] Data Quality Summary:")
    print(f"  - Total records: {len(consolidated)}")
    print(f"  - Columns: {', '.join(consolidated.columns)}")
    print(f"  - Date range: {consolidated['Date'].min()} to {consolidated['Date'].max()}")
    print(f"  - Missing values: {consolidated.isnull().sum().sum()}")

    # Step 5: Optional - Sort by date
    if 'Date' in consolidated.columns:
        consolidated['Date'] = pd.to_datetime(consolidated['Date'])
        consolidated = consolidated.sort_values('Date')
        print("✓ Sorted by date")

    # Step 6: Save consolidated file
    print(f"\n[5] Saving to '{output_file}'...")
    consolidated.to_excel(output_file, index=False)
    print(f"✓ Saved {len(consolidated)} records")

    # Step 7: Generate summary report
    print(f"\n[6] Generating summary...")
    generate_summary(consolidated, 'consolidation_summary.xlsx')

    print("\n" + "=" * 60)
    print("✓ CONSOLIDATION COMPLETE!")
    print("=" * 60)
    print(f"\nOutput files created:")
    print(f"  📊 {output_file} - Full consolidated data")
    print(f"  📈 consolidation_summary.xlsx - Summary report")


def create_sample_files():
    """Create sample files for demonstration"""

    import random

    products = ['Widget A', 'Widget B', 'Gadget Pro', 'Tool Set', 'Parts Kit']
    regions = ['North', 'South', 'East', 'West']

    for month in ['January', 'February', 'March']:
        data = {
            'Date': pd.date_range(start=f'2024-{["January", "February", "March"].index(month) + 1}-01',
                                  periods=30, freq='D'),
            'Product': [random.choice(products) for _ in range(30)],
            'Region': [random.choice(regions) for _ in range(30)],
            'Units': [random.randint(5, 50) for _ in range(30)],
            'Price': [round(random.uniform(10, 100), 2) for _ in range(30)]
        }
        df = pd.DataFrame(data)
        df['Revenue'] = df['Units'] * df['Price']

        filename = f'sales_{month.lower()}.xlsx'
        df.to_excel(filename, index=False)
        print(f"  ✓ Created {filename}")


def generate_summary(df, output_file):
    """Generate a summary report from consolidated data"""

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

        # Sheet 1: Overall Summary
        summary_data = {
            'Metric': [
                'Total Records',
                'Total Revenue',
                'Average Revenue per Transaction',
                'Date Range',
                'Unique Products',
                'Unique Regions'
            ],
            'Value': [
                len(df),
                f"${df['Revenue'].sum():,.2f}" if 'Revenue' in df.columns else 'N/A',
                f"${df['Revenue'].mean():,.2f}" if 'Revenue' in df.columns else 'N/A',
                f"{df['Date'].min()} to {df['Date'].max()}" if 'Date' in df.columns else 'N/A',
                df['Product'].nunique() if 'Product' in df.columns else 'N/A',
                df['Region'].nunique() if 'Region' in df.columns else 'N/A'
            ]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)

        # Sheet 2: By Product
        if 'Product' in df.columns and 'Revenue' in df.columns:
            product_summary = df.groupby('Product').agg({
                'Units': 'sum',
                'Revenue': 'sum'
            }).sort_values('Revenue', ascending=False).reset_index()
            product_summary.to_excel(writer, sheet_name='By Product', index=False)

        # Sheet 3: By Region
        if 'Region' in df.columns and 'Revenue' in df.columns:
            region_summary = df.groupby('Region').agg({
                'Units': 'sum',
                'Revenue': 'sum'
            }).sort_values('Revenue', ascending=False).reset_index()
            region_summary.to_excel(writer, sheet_name='By Region', index=False)

        # Sheet 4: By Month
        if 'Date' in df.columns and 'Revenue' in df.columns:
            df['Month'] = pd.to_datetime(df['Date']).dt.to_period('M')
            monthly_summary = df.groupby('Month').agg({
                'Revenue': 'sum',
                'Units': 'sum'
            }).reset_index()
            monthly_summary['Month'] = monthly_summary['Month'].astype(str)
            monthly_summary.to_excel(writer, sheet_name='By Month', index=False)

    print(f"✓ Summary saved to '{output_file}'")


def main():
    """Main execution"""

    # Option 1: Basic consolidation with default settings
    consolidate_sales_files()

    # Option 2: Custom pattern and output
    # consolidate_sales_files(file_pattern='report_*.xlsx', output_file='my_consolidated_report.xlsx')


if __name__ == "__main__":
    main()