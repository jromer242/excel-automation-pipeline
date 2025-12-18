# Examples Directory

This directory contains practical examples demonstrating different use cases for the Excel Automation Pipeline.

## Available Examples

### 1. `basic_consolidation.py`
**Use Case:** Combine multiple monthly sales files into one
```bash
python examples/basic_consolidation.py
```

### 2. `inventory_alerts.py`
**Use Case:** Automated inventory monitoring with alerts
```bash
python examples/inventory_alerts.py
```

### 3. `sales_dashboard.py`
**Use Case:** Generate a monthly sales dashboard report
```bash
python examples/sales_dashboard.py
```

### 4. `cross_file_analysis.py`
**Use Case:** Join data from multiple Excel files using SQL
```bash
python examples/cross_file_analysis.py
```

### 5. `safe_sheet_update.py`
**Use Case:** Update one worksheet without affecting others
```bash
python examples/safe_sheet_update.py
```

## Creating Your Own Examples

1. Copy a similar example as a template
2. Modify for your use case
3. Add comments explaining each step
4. Test with sample data
5. Share via pull request!

## Example Template

```python
"""
[Example Name] - Brief description

Use Case: Explain what problem this solves
Author: Your name
Date: YYYY-MM-DD
"""

import pandas as pd
import sqlite3

def main():
    """Main function with clear steps"""
    
    # Step 1: Load data
    df = pd.read_excel('your_file.xlsx')
    
    # Step 2: Process data
    # Your logic here
    
    # Step 3: Save results
    df.to_excel('output.xlsx', index=False)
    
    print("✓ Complete!")

if __name__ == "__main__":
    main()
```

## Need Help?

Check the [Tutorial](../docs/TUTORIAL.md) or [API Reference](../docs/API_REFERENCE.md) for more guidance.