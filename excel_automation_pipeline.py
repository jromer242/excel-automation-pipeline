"""
Low-Tech to High-Impact: Excel Automation with Python & SQLite
A practical demonstration of using technical skills in traditional environments
"""

import pandas as pd
import sqlite3
from datetime import datetime, timedelta
import random
from pathlib import Path


class ExcelAnalyticsPipeline:
    """Automates analysis of Excel files using Python and SQLite"""

    def __init__(self, db_name='business_data.db'):
        self.db_name = db_name
        self.conn = sqlite3.connect(db_name)

    def create_sample_excel_files(self):
        """Generate realistic sample Excel files (simulating existing business files)"""

        # 1. Sales Data (monthly exports from POS system)
        dates = pd.date_range(start='2024-01-01', end='2024-12-15', freq='D')
        sales_data = {
            'Date': random.choices(dates, k=500),
            'Product_ID': [f'PROD{random.randint(100, 150)}' for _ in range(500)],
            'Product_Name': random.choices(['Widget A', 'Widget B', 'Gadget Pro',
                                            'Tool Set', 'Parts Kit'], k=500),
            'Quantity': [random.randint(1, 20) for _ in range(500)],
            'Unit_Price': [round(random.uniform(10, 200), 2) for _ in range(500)],
            'Customer_Type': random.choices(['Retail', 'Wholesale', 'Online'], k=500)
        }
        sales_df = pd.DataFrame(sales_data)
        sales_df['Total_Sale'] = sales_df['Quantity'] * sales_df['Unit_Price']
        sales_df.to_excel('monthly_sales.xlsx', index=False)
        print("✓ Created monthly_sales.xlsx")

        # 2. Inventory Data (manual updates from warehouse)
        inventory_data = {
            'Product_ID': [f'PROD{i}' for i in range(100, 151)],
            'Product_Name': ['Widget A'] * 10 + ['Widget B'] * 10 + ['Gadget Pro'] * 11 +
                            ['Tool Set'] * 10 + ['Parts Kit'] * 10,
            'Current_Stock': [random.randint(0, 200) for _ in range(51)],
            'Reorder_Point': [random.randint(20, 50) for _ in range(51)],
            'Supplier': random.choices(['Supplier X', 'Supplier Y', 'Supplier Z'], k=51)
        }
        inventory_df = pd.DataFrame(inventory_data)
        inventory_df.to_excel('current_inventory.xlsx', index=False)
        print("✓ Created current_inventory.xlsx")

        # 3. Customer Data (from manual CRM spreadsheet)
        customer_data = {
            'Customer_ID': [f'CUST{i:04d}' for i in range(1, 101)],
            'Company_Name': [f'Company {i}' for i in range(1, 101)],
            'Type': random.choices(['Retail', 'Wholesale', 'Online'], k=100),
            'Region': random.choices(['North', 'South', 'East', 'West'], k=100),
            'Credit_Limit': [random.choice([5000, 10000, 25000, 50000]) for _ in range(100)]
        }
        customer_df = pd.DataFrame(customer_data)
        customer_df.to_excel('customer_list.xlsx', index=False)
        print("✓ Created customer_list.xlsx")

        return sales_df, inventory_df, customer_df

    def load_excel_to_sqlite(self):
        """Load all Excel files into a SQLite database for analysis"""

        # Load each Excel file and create corresponding tables
        files_to_load = {
            'sales': 'monthly_sales.xlsx',
            'inventory': 'current_inventory.xlsx',
            'customers': 'customer_list.xlsx'
        }

        for table_name, file_path in files_to_load.items():
            if Path(file_path).exists():
                df = pd.read_excel(file_path)
                df.to_sql(table_name, self.conn, if_exists='replace', index=False)
                print(f"✓ Loaded {file_path} → '{table_name}' table ({len(df)} records)")

    def run_analytics(self):
        """Run SQL analytics on the consolidated data"""

        print("\n" + "=" * 60)
        print("ANALYTICS REPORT")
        print("=" * 60)

        # Analysis 1: Top Products by Revenue
        query1 = """
        SELECT 
            Product_Name,
            COUNT(*) as Total_Orders,
            SUM(Quantity) as Units_Sold,
            ROUND(SUM(Total_Sale), 2) as Total_Revenue
        FROM sales
        GROUP BY Product_Name
        ORDER BY Total_Revenue DESC
        """
        print("\n📊 TOP PRODUCTS BY REVENUE:")
        print(pd.read_sql_query(query1, self.conn).to_string(index=False))

        # Analysis 2: Inventory Alert (items below reorder point)
        query2 = """
        SELECT 
            i.Product_Name,
            i.Current_Stock,
            i.Reorder_Point,
            i.Supplier,
            (i.Reorder_Point - i.Current_Stock) as Units_Needed
        FROM inventory i
        WHERE i.Current_Stock < i.Reorder_Point
        ORDER BY Units_Needed DESC
        """
        print("\n⚠️  REORDER ALERTS:")
        reorder_df = pd.read_sql_query(query2, self.conn)
        if len(reorder_df) > 0:
            print(reorder_df.to_string(index=False))
        else:
            print("No items need reordering!")

        # Analysis 3: Sales by Customer Type and Month
        query3 = """
        SELECT 
            strftime('%Y-%m', Date) as Month,
            Customer_Type,
            COUNT(*) as Transactions,
            ROUND(SUM(Total_Sale), 2) as Revenue
        FROM sales
        GROUP BY Month, Customer_Type
        ORDER BY Month DESC, Revenue DESC
        LIMIT 10
        """
        print("\n📈 RECENT SALES BY CUSTOMER TYPE:")
        print(pd.read_sql_query(query3, self.conn).to_string(index=False))

        # Analysis 4: Product Performance vs Inventory
        query4 = """
        SELECT 
            s.Product_Name,
            ROUND(AVG(s.Quantity), 1) as Avg_Order_Size,
            i.Current_Stock,
            ROUND(i.Current_Stock / AVG(s.Quantity), 1) as Days_of_Stock
        FROM sales s
        JOIN inventory i ON s.Product_ID = i.Product_ID
        GROUP BY s.Product_Name
        ORDER BY Days_of_Stock ASC
        """
        print("\n📦 INVENTORY EFFICIENCY:")
        print(pd.read_sql_query(query4, self.conn).to_string(index=False))

    def export_report(self):
        """Export analysis results back to Excel for sharing with non-technical stakeholders"""

        with pd.ExcelWriter('automated_report.xlsx', engine='openpyxl') as writer:
            # Sheet 1: Summary Dashboard
            summary_query = """
            SELECT 
                COUNT(DISTINCT Product_ID) as Total_Products,
                COUNT(*) as Total_Transactions,
                ROUND(SUM(Total_Sale), 2) as Total_Revenue,
                ROUND(AVG(Total_Sale), 2) as Avg_Transaction_Value
            FROM sales
            """
            summary = pd.read_sql_query(summary_query, self.conn)
            summary.to_excel(writer, sheet_name='Summary', index=False)

            # Sheet 2: Product Performance
            product_perf = pd.read_sql_query("""
                SELECT Product_Name, COUNT(*) as Orders, 
                       SUM(Quantity) as Units, ROUND(SUM(Total_Sale), 2) as Revenue
                FROM sales GROUP BY Product_Name ORDER BY Revenue DESC
            """, self.conn)
            product_perf.to_excel(writer, sheet_name='Product Performance', index=False)

            # Sheet 3: Reorder List
            reorder = pd.read_sql_query("""
                SELECT Product_Name, Current_Stock, Reorder_Point, Supplier
                FROM inventory WHERE Current_Stock < Reorder_Point
            """, self.conn)
            reorder.to_excel(writer, sheet_name='Reorder Needed', index=False)

        print("\n✓ Exported automated_report.xlsx")

    def close(self):
        """Clean up database connection"""
        self.conn.close()


def main():
    """Main execution flow"""
    print("=" * 60)
    print("EXCEL AUTOMATION PIPELINE")
    print("Transforming Low-Tech Files with High-Impact Analytics")
    print("=" * 60)

    # Initialize pipeline
    pipeline = ExcelAnalyticsPipeline()

    # Step 1: Create sample files (simulating existing business Excel files)
    print("\n[1] Creating sample Excel files...")
    pipeline.create_sample_excel_files()

    # Step 2: Load into SQLite for analysis
    print("\n[2] Loading Excel files into SQLite database...")
    pipeline.load_excel_to_sqlite()

    # Step 3: Run analytics
    print("\n[3] Running automated analytics...")
    pipeline.run_analytics()

    # Step 4: Export results
    print("\n[4] Exporting report for stakeholders...")
    pipeline.export_report()

    # Cleanup
    pipeline.close()

    print("\n" + "=" * 60)
    print("✓ PIPELINE COMPLETE!")
    print("=" * 60)
    print("\nWhat this demonstrates:")
    print("• Python replaces manual Excel consolidation")
    print("• SQLite enables complex cross-file analysis")
    print("• Automation reduces hours of work to seconds")
    print("• Results exported back to Excel for non-technical users")
    print("\nPerfect for environments where:")
    print("✓ Teams use Excel but need better insights")
    print("✓ No budget for expensive BI tools")
    print("✓ Data lives in multiple spreadsheets")
    print("✓ Manual reporting is time-consuming")


if __name__ == "__main__":
    main()