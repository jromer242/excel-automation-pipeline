# Quick Reference Guide

A cheat sheet for common Excel automation tasks.

## Table of Contents
- [Reading Excel](#reading-excel)
- [Writing Excel](#writing-excel)
- [Data Operations](#data-operations)
- [SQLite Queries](#sqlite-queries)
- [Common Patterns](#common-patterns)

---

## Reading Excel

### Basic Read
```python
import pandas as pd

# Single sheet
df = pd.read_excel('file.xlsx')

# Specific sheet
df = pd.read_excel('file.xlsx', sheet_name='Sales')

# All sheets
sheets = pd.read_excel('file.xlsx', sheet_name=None)
```

### Advanced Read
```python
# Specific columns
df = pd.read_excel('file.xlsx', usecols=['Name', 'Sales'])

# Skip rows
df = pd.read_excel('file.xlsx', skiprows=2)

# Specify header row
df = pd.read_excel('file.xlsx', header=1)

# No header
df = pd.read_excel('file.xlsx', header=None)
```

---

## Writing Excel

### Basic Write
```python
# Single sheet
df.to_excel('output.xlsx', index=False)

# Multiple sheets
with pd.ExcelWriter('output.xlsx') as writer:
    df1.to_excel(writer, sheet_name='Sheet1', index=False)
    df2.to_excel(writer, sheet_name='Sheet2', index=False)
```

### Safe Update (Preserve Other Sheets)
```python
# Update one sheet without touching others
with pd.ExcelWriter('file.xlsx', engine='openpyxl', 
                    mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='SheetName', index=False)
```

---

## Data Operations

### Filter Data
```python
# Single condition
high_sales = df[df['Sales'] > 1000]

# Multiple conditions
filtered = df[(df['Sales'] > 1000) & (df['Region'] == 'North')]

# Contains text
contains_widget = df[df['Product'].str.contains('Widget')]
```

### Sort Data
```python
# Ascending
sorted_df = df.sort_values('Sales')

# Descending
sorted_df = df.sort_values('Sales', ascending=False)

# Multiple columns
sorted_df = df.sort_values(['Region', 'Sales'], ascending=[True, False])
```

### Group and Aggregate
```python
# Simple groupby
summary = df.groupby('Product')['Sales'].sum()

# Multiple aggregations
summary = df.groupby('Product').agg({
    'Sales': 'sum',
    'Units': 'mean',
    'Orders': 'count'
})

# Reset index for clean output
summary = summary.reset_index()
```

### Add Calculated Columns
```python
# Simple calculation
df['Total'] = df['Units'] * df['Price']

# Conditional calculation
df['Category'] = df['Sales'].apply(lambda x: 'High' if x > 1000 else 'Low')

# Multiple conditions
df['Status'] = df.apply(
    lambda row: 'Reorder' if row['Stock'] < row['Min'] else 'OK',
    axis=1
)
```

---

## SQLite Queries

### Setup
```python
import sqlite3
conn = sqlite3.connect('data.db')

# Load data
df.to_sql('table_name', conn, if_exists='replace', index=False)
```

### Basic Queries
```python
# Select all
result = pd.read_sql_query("SELECT * FROM table_name", conn)

# Where clause
result = pd.read_sql_query(
    "SELECT * FROM sales WHERE Sales > 1000", 
    conn
)

# Group by
result = pd.read_sql_query("""
    SELECT Product, SUM(Sales) as Total
    FROM sales
    GROUP BY Product
    ORDER BY Total DESC
""", conn)
```

### Joins
```python
# Inner join
result = pd.read_sql_query("""
    SELECT s.*, i.Stock
    FROM sales s
    JOIN inventory i ON s.Product_ID = i.Product_ID
""", conn)

# Left join
result = pd.read_sql_query("""
    SELECT s.*, i.Stock
    FROM sales s
    LEFT JOIN inventory i ON s.Product_ID = i.Product_ID
""", conn)
```

### Cleanup
```python
conn.close()
```

---

## Common Patterns

### Pattern 1: Consolidate Multiple Files
```python
import pandas as pd
from glob import glob

# Get all Excel files
files = glob('sales_*.xlsx')

# Read and combine
dfs = [pd.read_excel(f) for f in files]
combined = pd.concat(dfs, ignore_index=True)

# Save
combined.to_excel('consolidated.xlsx', index=False)
```

### Pattern 2: Pivot Table
```python
pivot = pd.pivot_table(
    df,
    values='Sales',
    index='Product',
    columns='Region',
    aggfunc='sum',
    fill_value=0
)
```

### Pattern 3: Month-over-Month Growth
```python
df['Date'] = pd.to_datetime(df['Date'])
df['Month'] = df['Date'].dt.to_period('M')

monthly = df.groupby('Month')['Sales'].sum()
monthly['Growth_%'] = monthly.pct_change() * 100
```

### Pattern 4: Conditional Formatting (Manual)
```python
# Identify items needing attention
df['Alert'] = df.apply(
    lambda row: '⚠️' if row['Stock'] < row['Reorder_Point'] else '✓',
    axis=1
)
```

### Pattern 5: Export with Timestamp
```python
from datetime import datetime

timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
filename = f'report_{timestamp}.xlsx'
df.to_excel(filename, index=False)
```

---

## Troubleshooting

### Common Errors

**FileNotFoundError**
```python
from pathlib import Path

if Path('file.xlsx').exists():
    df = pd.read_excel('file.xlsx')
else:
    print("File not found!")
```

**PermissionError** (File is open)
```python
import time

try:
    df.to_excel('file.xlsx')
except PermissionError:
    print("Close the Excel file first!")
```

**KeyError** (Column doesn't exist)
```python
if 'ColumnName' in df.columns:
    result = df['ColumnName']
else:
    print("Column not found!")
```

---

## Performance Tips

### Large Files
```python
# Read in chunks
chunks = pd.read_excel('large_file.xlsx', chunksize=10000)
for chunk in chunks:
    process(chunk)

# Read only needed columns
df = pd.read_excel('file.xlsx', usecols=['A', 'B', 'C'])
```

### Speed Up Processing
```python
# Use categorical for repeated values
df['Category'] = df['Category'].astype('category')

# Vectorized operations (fast)
df['Total'] = df['Units'] * df['Price']

# Avoid loops (slow)
# for i, row in df.iterrows():  # DON'T DO THIS
#     df.at[i, 'Total'] = row['Units'] * row['Price']
```

---

## Resources

- [Pandas Docs](https://pandas.pydata.org/docs/)
- [SQLite Tutorial](https://www.sqlitetutorial.net/)
- [Python Excel](https://www.python-excel.org/)

---

**💡 Tip:** Bookmark this page for quick reference during development!