# Tutorial: Excel Automation from Scratch

## Table of Contents
1. [Introduction](#introduction)
2. [Setup](#setup)
3. [Basic Concepts](#basic-concepts)
4. [Step-by-Step Examples](#step-by-step-examples)
5. [Advanced Techniques](#advanced-techniques)
6. [Best Practices](#best-practices)

## Introduction

This tutorial will guide you through automating Excel workflows using Python. No prior Python experience with Excel required!

### What You'll Learn
- Load Excel files into Python
- Perform analytics using SQL
- Edit Excel files safely
- Generate automated reports
- Build a complete data pipeline

### Time Required
- Basic: 30 minutes
- Complete: 2 hours

## Setup

### 1. Install Python
Download from [python.org](https://python.org) (version 3.8+)

### 2. Set Up Project
```bash
# Create project folder
mkdir excel-automation
cd excel-automation

# Create virtual environment
python -m venv venv

# Activate it
# Windows:
venv\Scripts\activate
# Mac/Linux:
source venv/bin/activate

# Install dependencies
pip install pandas openpyxl
```

### 3. Verify Installation
```python
import pandas as pd
import openpyxl
print("Setup complete!")
```

## Basic Concepts

### Excel Files in Python

**Reading Excel:**
```python
import pandas as pd

# Read a single sheet
df = pd.read_excel('data.xlsx', sheet_name='Sheet1')

# Read all sheets
all_sheets = pd.read_excel('data.xlsx', sheet_name=None)

# Read specific columns
df = pd.read_excel('data.xlsx', usecols=['Name', 'Sales'])
```

**Writing Excel:**
```python
# Write a DataFrame to Excel
df.to_excel('output.xlsx', index=False)

# Write multiple sheets
with pd.ExcelWriter('output.xlsx') as writer:
    df1.to_excel(writer, sheet_name='Sales')
    df2.to_excel(writer, sheet_name='Inventory')
```

### SQLite Basics

SQLite lets you run SQL queries on your data without a database server!

```python
import sqlite3
import pandas as pd

# Create database connection
conn = sqlite3.connect('my_data.db')

# Load Excel into SQLite
df = pd.read_excel('sales.xlsx')
df.to_sql('sales', conn, if_exists='replace')

# Query with SQL
result = pd.read_sql_query("""
    SELECT Product, SUM(Sales) as Total
    FROM sales
    GROUP BY Product
""", conn)

conn.close()
```

## Step-by-Step Examples

### Example 1: Consolidate Multiple Files

**Scenario:** You have 3 monthly sales files to combine.

```python
import pandas as pd

# Read multiple files
jan = pd.read_excel('sales_jan.xlsx')
feb = pd.read_excel('sales_feb.xlsx')
mar = pd.read_excel('sales_mar.xlsx')

# Combine them
all_sales = pd.concat([jan, feb, mar], ignore_index=True)

# Save consolidated file
all_sales.to_excel('Q1_sales.xlsx', index=False)

print(f"Consolidated {len(all_sales)} records!")
```

### Example 2: Add Calculations

**Scenario:** Add profit margin to sales data.

```python
import pandas as pd

# Read data
df = pd.read_excel('sales.xlsx')

# Add calculations
df['Revenue'] = df['Units'] * df['Price']
df['Cost'] = df['Units'] * df['Unit_Cost']
df['Profit'] = df['Revenue'] - df['Cost']
df['Margin_%'] = (df['Profit'] / df['Revenue'] * 100).round(2)

# Save with new columns
df.to_excel('sales_with_calculations.xlsx', index=False)
```

### Example 3: Filter and Summary

**Scenario:** Find top products by revenue.

```python
import pandas as pd

df = pd.read_excel('sales.xlsx')

# Calculate revenue
df['Revenue'] = df['Units'] * df['Price']

# Group by product
summary = df.groupby('Product').agg({
    'Units': 'sum',
    'Revenue': 'sum'
}).reset_index()

# Sort by revenue
summary = summary.sort_values('Revenue', ascending=False)

# Get top 10
top_10 = summary.head(10)

# Save
top_10.to_excel('top_products.xlsx', index=False)
```

### Example 4: Cross-File Analysis with SQLite

**Scenario:** Join sales data with inventory data.

```python
import pandas as pd
import sqlite3

# Create database
conn = sqlite3.connect('analysis.db')

# Load multiple files
sales = pd.read_excel('sales.xlsx')
inventory = pd.read_excel('inventory.xlsx')

sales.to_sql('sales', conn, if_exists='replace')
inventory.to_sql('inventory', conn, if_exists='replace')

# Run SQL query joining both
query = """
SELECT 
    s.Product,
    SUM(s.Units) as Units_Sold,
    i.Current_Stock,
    i.Reorder_Point,
    CASE 
        WHEN i.Current_Stock < i.Reorder_Point 
        THEN 'REORDER NEEDED'
        ELSE 'OK'
    END as Status
FROM sales s
JOIN inventory i ON s.Product_ID = i.Product_ID
GROUP BY s.Product
"""

result = pd.read_sql_query(query, conn)
result.to_excel('inventory_status.xlsx', index=False)

conn.close()
```

### Example 5: Safe Worksheet Editing

**Scenario:** Update one sheet without touching others.

```python
import pandas as pd

# Read the sheet you want to update
df = pd.read_excel('report.xlsx', sheet_name='Sales')

# Make changes
df['Updated_Date'] = '2024-12-17'
df['Tax'] = df['Revenue'] * 0.08

# Write back ONLY this sheet
with pd.ExcelWriter('report.xlsx', engine='openpyxl', 
                    mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='Sales', index=False)

print("Updated Sales sheet only!")
```

## Advanced Techniques

### 1. Automated Alerts

Send email alerts when inventory is low:

```python
import pandas as pd
import sqlite3

conn = sqlite3.connect('data.db')

# Load data
inventory = pd.read_excel('inventory.xlsx')
inventory.to_sql('inventory', conn, if_exists='replace')

# Find low stock items
low_stock = pd.read_sql_query("""
    SELECT * FROM inventory 
    WHERE Current_Stock < Reorder_Point
""", conn)

if len(low_stock) > 0:
    low_stock.to_excel('URGENT_reorder_needed.xlsx', index=False)
    print(f"⚠️  Alert: {len(low_stock)} items need reordering!")
else:
    print("✓ All inventory levels OK")

conn.close()
```

### 2. Time-Series Analysis

Track trends over time:

```python
import pandas as pd

# Read sales data
df = pd.read_excel('sales_history.xlsx')

# Convert to datetime
df['Date'] = pd.to_datetime(df['Date'])

# Group by month
monthly = df.groupby(df['Date'].dt.to_period('M')).agg({
    'Revenue': 'sum',
    'Units': 'sum'
}).reset_index()

# Calculate month-over-month growth
monthly['Revenue_Growth_%'] = monthly['Revenue'].pct_change() * 100

monthly.to_excel('monthly_trends.xlsx', index=False)
```

### 3. Pivot Tables in Python

Create Excel-style pivot tables:

```python
import pandas as pd

df = pd.read_excel('sales.xlsx')

# Create pivot table
pivot = pd.pivot_table(
    df,
    values='Revenue',
    index='Product',
    columns='Region',
    aggfunc='sum',
    fill_value=0
)

# Add totals
pivot['Total'] = pivot.sum(axis=1)

pivot.to_excel('sales_by_region.xlsx')
```

## Best Practices

### 1. Always Use Virtual Environments
```bash
python -m venv venv
source venv/bin/activate  # or venv\Scripts\activate on Windows
```

### 2. Handle Errors Gracefully
```python
import pandas as pd

try:
    df = pd.read_excel('sales.xlsx')
except FileNotFoundError:
    print("Error: File not found!")
    exit()
except Exception as e:
    print(f"Error: {e}")
    exit()
```

### 3. Close Database Connections
```python
conn = sqlite3.connect('data.db')
try:
    # Your code here
    pass
finally:
    conn.close()
```

### 4. Use Descriptive Variable Names
```python
# Bad
df1 = pd.read_excel('file1.xlsx')
df2 = pd.read_excel('file2.xlsx')

# Good
sales_data = pd.read_excel('sales.xlsx')
inventory_data = pd.read_excel('inventory.xlsx')
```

### 5. Comment Your Code
```python
# Load sales data from last quarter
sales = pd.read_excel('Q4_sales.xlsx')

# Calculate profit margin
sales['Margin'] = (sales['Profit'] / sales['Revenue']) * 100

# Filter for high-margin products (>30%)
high_margin = sales[sales['Margin'] > 30]
```

### 6. Test on Sample Data First
Always test your script on a copy of your data before running on the real thing!

```python
# Make a backup
import shutil
shutil.copy('important_data.xlsx', 'important_data_BACKUP.xlsx')
```

## Common Pitfalls to Avoid

❌ **Don't:** Overwrite files without backups
✅ **Do:** Create backups before modifying

❌ **Don't:** Use `mode='w'` when updating Excel (overwrites everything)
✅ **Do:** Use `mode='a'` with `if_sheet_exists='replace'`

❌ **Don't:** Keep Excel files open while running scripts
✅ **Do:** Close files before running automation

❌ **Don't:** Hard-code file paths
✅ **Do:** Use variables or configuration files

## Next Steps

1. Try the examples in this tutorial
2. Modify them for your own data
3. Build a complete pipeline
4. Share your results on LinkedIn!

## Getting Help

- Check the [API Reference](API_REFERENCE.md) for detailed function docs
- Browse [Examples](EXAMPLES.md) for more use cases
- Open an issue on GitHub for bugs
- Join discussions for questions

---

**You're ready to automate! 🚀**

Start with Example 1 and work your way up. Remember: every expert was once a beginner!