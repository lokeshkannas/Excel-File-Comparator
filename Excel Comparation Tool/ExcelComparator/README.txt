# 📦 Excel Workbook Comparator

A modern, dark-themed GUI tool to compare two Excel workbooks (SSRS vs Power BI) for:
- Structure (sheets, row/column counts, column names/order)
- Data types
- Cell-by-cell values (with numeric tolerance)

## ▶ Quick Start
- Windows: Double-click `run.bat`
- macOS / Linux: `chmod +x run.sh && ./run.sh`

Reports are saved to `~/Documents/Excel_Comparisons/<your_report_name>.xlsx`

## Report Contents
- **Summary**: Overall counts and pass/fail
- **Structure_Issues**: Sheet/row/column/order/name differences
- **Dtype_Issues**: Data type inconsistencies
- **Value_Mismatches**: Cell-level mismatches with row/column details

## Tips
- Close Excel files before comparing.
- Use descriptive report names (e.g., `Oct_Sales_Check.xlsx`).
- Numeric tolerance can be adjusted in `config.json`.
