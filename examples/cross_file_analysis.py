"""
Cross-File Analysis with SQL

Use Case: Join and analyze data from multiple Excel files using SQL queries
Author: Excel Automation Pipeline
Date: 2024-12-17

This demonstrates the power of SQLite for analyzing data spread across
multiple Excel files - something that's tedious or impossible in vanilla Excel.
"""

import pandas as pd
import sqlite3
from pathlib import Path
from datetime import datetime
import random


def create_sample_data():
    """Create sample Excel files for demonstration"""

    print("📝 Creating sample data files...")

    # File 1: Sales transactions
    sales_data = {
        'Transaction_ID': [f'TXN{i:05d}' for i in range(1, 101)],
        'Date': pd.date_range(start='2024-01-01', periods=100, freq='D'),
        'Product_ID': [f'PROD{random.randint(1, 20):03d}' for _ in range(100)],
        'Customer_ID': [f'CUST{random.randint(1, 30):03d}' for _ in range(100)],
        'Quantity': [random.randint(1, 10) for _ in range(100)],
        'Sale_Amount': [round(random.uniform(50, 500), 2) for _ in range(100)]
    }
    pd.DataFrame(sales_data).to_excel('sales.xlsx', index=False)
    print("  ✓ sales.xlsx")

    # File 2: Product catalog
    product_data = {
        'Product_ID': [f'PROD{i:03d}' for i in range(1, 21)],
        'Product_Name': [f'Product {i}' for i in range(1, 21)],
        'Category': [random.choice(['Electronics', 'Furniture', 'Office', 'Tools']) for _ in range(20)],
        'Unit_Cost': [round(random.uniform(10, 200), 2) for _ in range(20)],
        'Unit_Price': [round(random.uniform(50, 400), 2) for _ in range(20)],
        'Supplier': [random.choice(['Supplier A', 'Supplier B', 'Supplier C']) for _ in range(20)]
    }
    pd.DataFrame(product_data).to_excel('products.xlsx', index=False)
    print("  ✓ products.xlsx")

    # File 3: Customer information
    customer_data = {
        'Customer_ID': [f'CUST{i:03d}' for i in range(1, 31)],
        'Customer_Name': [f'Customer {i}' for i in range(1, 31)],
        'Region': [random.choice(['North', 'South', 'East', 'West']) for _ in range(30)],
        'Customer_Type': [random.choice(['Retail', 'Wholesale', 'VIP']) for _ in range(30)],
        'Credit_Limit': [random.choice([5000, 10000, 25000, 50000]) for _ in range(30)]
    }
    pd.DataFrame(customer_data).to_excel('customers.xlsx', index=False)
    print("  ✓ customers.xlsx")

    # File 4: Inventory levels
    inventory_data = {
        'Product_ID': [f'PROD{i:03d}' for i in range(1, 21)],
        'Current_Stock': [random.randint(0, 100) for _ in range(20)],
        'Reorder_Point': [random.randint(10, 30) for _ in range(20)],
        'Warehouse': [random.choice(['Warehouse A', 'Warehouse B']) for _ in range(20)]
    }
    pd.DataFrame(inventory_data).to_excel('inventory.xlsx', index=False)
    print("  ✓ inventory.xlsx")


def load_files_to_database(db_name='analysis.db'):
    """Load all Excel files into SQLite database"""

    print("\n📂 Loading Excel files into SQLite database...")

    # Connect to database
    conn = sqlite3.connect(db_name)

    # Define files to load
    files_to_load = {
        'sales': 'sales.xlsx',
        'products': 'products.xlsx',
        'customers': 'customers.xlsx',
        'inventory': 'inventory.xlsx'
    }

    # Load each file
    for table_name, filename in files_to_load.items():
        if Path(filename).exists():
            df = pd.read_excel(filename)
            df.to_sql(table_name, conn, if_exists='replace', index=False)
            print(f"  ✓ {filename} → {table_name} table ({len(df)} rows)")
        else:
            print(f"  ❌ {filename} not found")

    return conn


def run_cross_file_analysis(conn):
    """Run SQL queries that join data from multiple files"""

    print("\n" + "=" * 70)
    print("CROSS-FILE ANALYSIS USING SQL")
    print("=" * 70)

    # Analysis 1: Sales with Product Details
    print("\n📊 Analysis 1: Sales Revenue by Product Category")
    query1 = """
    SELECT 
        p.Category,
        COUNT(DISTINCT s.Transaction_ID) as Total_Transactions,
        SUM(s.Quantity) as Total_Units_Sold,
        SUM(s.Sale_Amount) as Total_Revenue,
        ROUND(AVG(s.Sale_Amount), 2) as Avg_Transaction_Value
    FROM sales s
    JOIN products p ON s.Product_ID = p.Product_ID
    GROUP BY p.Category
    ORDER BY Total_Revenue DESC
    """
    result1 = pd.read_sql_query(query1, conn)
    print(result1.to_string(index=False))

    # Analysis 2: Customer Purchase Patterns
    print("\n📊 Analysis 2: Top Customers with Regional Breakdown")
    query2 = """
    SELECT 
        c.Customer_Name,
        c.Region,
        c.Customer_Type,
        COUNT(s.Transaction_ID) as Total_Purchases,
        SUM(s.Sale_Amount) as Total_Spent,
        ROUND(AVG(s.Sale_Amount), 2) as Avg_Purchase
    FROM sales s
    JOIN customers c ON s.Customer_ID = c.Customer_ID
    GROUP BY c.Customer_ID
    ORDER BY Total_Spent DESC
    LIMIT 10
    """
    result2 = pd.read_sql_query(query2, conn)
    print(result2.to_string(index=False))

    # Analysis 3: Inventory Status with Sales Performance
    print("\n📊 Analysis 3: Products Needing Reorder (with Sales Velocity)")
    query3 = """
    SELECT 
        p.Product_Name,
        p.Category,
        i.Current_Stock,
        i.Reorder_Point,
        i.Warehouse,
        COUNT(s.Transaction_ID) as Times_Sold,
        SUM(s.Quantity) as Total_Units_Sold,
        ROUND(SUM(s.Quantity) / 100.0, 2) as Avg_Daily_Sales,
        CASE 
            WHEN i.Current_Stock < i.Reorder_Point THEN 'REORDER NOW'
            WHEN i.Current_Stock < i.Reorder_Point * 1.5 THEN 'Monitor'
            ELSE 'OK'
        END as Status
    FROM inventory i
    JOIN products p ON i.Product_ID = p.Product_ID
    LEFT JOIN sales s ON p.Product_ID = s.Product_ID
    GROUP BY i.Product_ID
    HAVING Status != 'OK'
    ORDER BY Current_Stock ASC
    """
    result3 = pd.read_sql_query(query3, conn)
    if len(result3) > 0:
        print(result3.to_string(index=False))
    else:
        print("  ✅ All inventory levels are adequate!")

    # Analysis 4: Profitability Analysis
    print("\n📊 Analysis 4: Product Profitability Analysis")
    query4 = """
    SELECT 
        p.Product_Name,
        p.Category,
        p.Supplier,
        COUNT(s.Transaction_ID) as Units_Sold,
        p.Unit_Cost,
        p.Unit_Price,
        ROUND(p.Unit_Price - p.Unit_Cost, 2) as Profit_Per_Unit,
        ROUND((p.Unit_Price - p.Unit_Cost) / p.Unit_Price * 100, 2) as Margin_Percent,
        ROUND((p.Unit_Price - p.Unit_Cost) * COUNT(s.Transaction_ID), 2) as Total_Profit
    FROM products p
    LEFT JOIN sales s ON p.Product_ID = s.Product_ID
    GROUP BY p.Product_ID
    ORDER BY Total_Profit DESC
    LIMIT 15
    """
    result4 = pd.read_sql_query(query4, conn)
    print(result4.to_string(index=False))

    # Analysis 5: Regional Performance with Customer Mix
    print("\n📊 Analysis 5: Regional Performance by Customer Type")
    query5 = """
    SELECT 
        c.Region,
        c.Customer_Type,
        COUNT(DISTINCT c.Customer_ID) as Customer_Count,
        COUNT(s.Transaction_ID) as Total_Transactions,
        SUM(s.Sale_Amount) as Total_Revenue,
        ROUND(AVG(s.Sale_Amount), 2) as Avg_Transaction
    FROM customers c
    LEFT JOIN sales s ON c.Customer_ID = s.Customer_ID
    GROUP BY c.Region, c.Customer_Type
    ORDER BY c.Region, Total_Revenue DESC
    """
    result5 = pd.read_sql_query(query5, conn)
    print(result5.to_string(index=False))

    # Return results for export
    return {
        'revenue_by_category': result1,
        'top_customers': result2,
        'inventory_alerts': result3,
        'profitability': result4,
        'regional_performance': result5
    }


def export_analysis_results(results, output_file='cross_file_analysis_results.xlsx'):
    """Export all analysis results to a single Excel workbook"""

    print(f"\n💾 Exporting results to '{output_file}'...")

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write each analysis to a separate sheet
        for sheet_name, df in results.items():
            # Format sheet name (replace underscores, capitalize)
            formatted_name = sheet_name.replace('_', ' ').title()[:31]  # Excel sheet name limit
            df.to_excel(writer, sheet_name=formatted_name, index=False)
            print(f"  ✓ {formatted_name}")

    print(f"✓ Results exported successfully!")


def demonstrate_complex_query(conn):
    """Demonstrate a complex multi-join query"""

    print("\n" + "=" * 70)
    print("BONUS: Complex Multi-Table Join")
    print("=" * 70)
    print("\n🎯 Finding: Products with high sales but low inventory by region\n")

    complex_query = """
    SELECT 
        p.Product_Name,
        p.Category,
        c.Region,
        COUNT(DISTINCT s.Customer_ID) as Unique_Customers,
        SUM(s.Quantity) as Total_Units_Sold,
        i.Current_Stock,
        i.Reorder_Point,
        ROUND(SUM(s.Sale_Amount), 2) as Total_Revenue,
        ROUND(SUM(s.Sale_Amount) / SUM(s.Quantity), 2) as Avg_Price_Per_Unit,
        CASE 
            WHEN i.Current_Stock < i.Reorder_Point AND SUM(s.Quantity) > 10 
            THEN '🚨 HIGH PRIORITY'
            WHEN i.Current_Stock < i.Reorder_Point * 1.5 AND SUM(s.Quantity) > 5 
            THEN '⚠️  MONITOR CLOSELY'
            ELSE '✅ OK'
        END as Alert_Status
    FROM sales s
    JOIN products p ON s.Product_ID = p.Product_ID
    JOIN customers c ON s.Customer_ID = c.Customer_ID
    JOIN inventory i ON p.Product_ID = i.Product_ID
    GROUP BY p.Product_ID, c.Region
    HAVING Alert_Status != '✅ OK'
    ORDER BY Total_Revenue DESC
    """

    result = pd.read_sql_query(complex_query, conn)

    if len(result) > 0:
        print(result.to_string(index=False))
        print(f"\n⚠️  Found {len(result)} product-region combinations needing attention!")
    else:
        print("✅ No critical inventory issues found across regions!")


def main():
    """Main execution"""

    print("=" * 70)
    print("CROSS-FILE ANALYSIS DEMONSTRATION")
    print("=" * 70)

    # Check if sample files exist
    if not all(Path(f).exists() for f in ['sales.xlsx', 'products.xlsx', 'customers.xlsx', 'inventory.xlsx']):
        create_sample_data()

    # Load files into database
    conn = load_files_to_database()

    # Run analysis
    results = run_cross_file_analysis(conn)

    # Demonstrate complex query
    demonstrate_complex_query(conn)

    # Export results
    export_analysis_results(results)

    # Clean up
    conn.close()

    print("\n" + "=" * 70)
    print("✓ ANALYSIS COMPLETE")
    print("=" * 70)
    print("\n💡 Key Takeaway:")
    print("   SQL lets you analyze data across multiple Excel files")
    print("   in ways that would be impossible with VLOOKUP alone!")
    print("\n📊 Check 'cross_file_analysis_results.xlsx' for full results")


if __name__ == "__main__":
    main()