# Excel Automation Pipeline 📊

> Transform low-tech Excel workflows into high-impact analytics using Python, SQLite, and automation

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 🎯 Overview

This project demonstrates how technical data analytics skills can transform traditional Excel-based workflows in low-tech environments. Perfect for small businesses, teams without expensive BI tools, or anyone working with multiple spreadsheets.

**Key Features:**
- 🔄 Automate consolidation of multiple Excel files
- 📊 Run complex SQL analytics across flat files
- ⚡ Reduce hours of manual work to seconds
- 📈 Generate automated reports for stakeholders
- 🛠️ Safe worksheet editing without affecting other sheets

## 🚀 Quick Start

### Prerequisites
- Python 3.8 or higher
- pip (Python package manager)

### Installation

1. Clone the repository:
```bash
git clone https://github.com/jromer242/excel-automation-pipeline.git
cd excel-automation-pipeline
```

2. Create a virtual environment (recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

### Basic Usage

Run the main pipeline demo:
```bash
python excel_automation_pipeline.py
```

This will:
1. Generate sample Excel files (sales, inventory, customers)
2. Load them into SQLite for analysis
3. Run automated analytics
4. Export results to a consolidated report

## 📁 Project Structure

```
excel-automation-pipeline/
│
├── excel_automation_pipeline.py    # Main pipeline demonstration
├── safe_worksheet_editing.py       # Methods for editing Excel safely
├── requirements.txt                # Python dependencies
├── README.md                       # This file
├── LICENSE                         # MIT License
│
├── docs/                           # Documentation
│   ├── TUTORIAL.md                # Step-by-step tutorial
│   ├── API_REFERENCE.md           # Detailed API documentation
│   └── EXAMPLES.md                # Additional examples
│
├── examples/                       # Example scripts
│   ├── basic_consolidation.py     # Simple multi-file consolidation
│   ├── inventory_alerts.py        # Automated inventory monitoring
│   └── sales_dashboard.py         # Monthly sales reporting
│
├── tests/                          # Unit tests
│   ├── test_pipeline.py
│   └── test_worksheet_editing.py
│
└── sample_data/                    # Sample Excel files (generated)
    ├── monthly_sales.xlsx
    ├── current_inventory.xlsx
    └── customer_list.xlsx
```

## 💡 Use Cases

### 1. Multi-File Consolidation
Combine data from multiple Excel exports into a single analysis:
```python
from excel_automation_pipeline import ExcelAnalyticsPipeline

pipeline = ExcelAnalyticsPipeline()
pipeline.load_excel_to_sqlite()
pipeline.run_analytics()
```

### 2. Safe Worksheet Editing
Update one sheet without touching others:
```python
import pandas as pd

df = pd.read_excel('report.xlsx', sheet_name='Sales')
df['Revenue'] = df['Units'] * df['Price']

with pd.ExcelWriter('report.xlsx', engine='openpyxl', 
                    mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='Sales', index=False)
```

### 3. Automated Reporting
Generate weekly reports automatically:
```python
pipeline = ExcelAnalyticsPipeline()
pipeline.load_excel_to_sqlite()
pipeline.export_report()  # Creates automated_report.xlsx
```

## 🔧 Configuration

The pipeline is highly configurable. Edit the `ExcelAnalyticsPipeline` class to:

- Change database name: `ExcelAnalyticsPipeline(db_name='custom.db')`
- Add custom analytics queries
- Modify report output format
- Add new data sources

## 📚 Documentation

- **[Tutorial](docs/TUTORIAL.md)**: Step-by-step guide for beginners
- **[API Reference](docs/API_REFERENCE.md)**: Detailed function documentation
- **[Examples](docs/EXAMPLES.md)**: Real-world usage scenarios

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes:

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 🐛 Troubleshooting

### Common Issues

**Issue: "openpyxl not found"**
```bash
pip install openpyxl
```

**Issue: Excel file is locked**
- Close the Excel file before running the script
- Check if another process is using the file

**Issue: Permission denied when saving**
- Ensure you have write permissions in the directory
- Run as administrator/sudo if needed

## 🎓 Learning Resources

- [Pandas Documentation](https://pandas.pydata.org/docs/)
- [SQLite Tutorial](https://www.sqlitetutorial.net/)
- [openpyxl Documentation](https://openpyxl.readthedocs.io/)

## 📊 Real-World Impact

This approach has helped:
- **Small businesses** eliminate 10+ hours/week of manual data entry
- **Analysts** consolidate 20+ Excel files into actionable insights
- **Teams** transition from manual reporting to automated dashboards
- **Organizations** leverage technical skills without expensive tools

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👤 Author

**Your Name**
- LinkedIn: [jylesromer](https://linkedin.com/in/jylesromer)
- GitHub: [@jromer242](https://github.com/jromer242)
- Blog: [Your Blog](https://yourblog.com)

## 🌟 Acknowledgments

- Inspired by real-world challenges in low-tech environments
- Built to demonstrate that technical skills create impact anywhere
- Thanks to the Python data science community

## 📧 Contact

Have questions or suggestions? Feel free to:
- Open an issue
- Start a discussion
- Reach out on LinkedIn

---

**⭐ If this project helped you, please consider giving it a star!**

*Built with ❤️ to bridge the gap between technical skills and traditional workflows*