# gj-excel-cleaner
Automated Excel data cleaner built with Python &amp; Pandas. Instantly removes duplicates, trims whitespace, and auto-formats columns for data analysis.

# Excel Data Cleaner

**Automated Excel data cleaner built with Python & Pandas. Instantly removes duplicates, trims whitespace, and auto-formats columns for data analysis.**

This tool was designed to solve the common "dirty data" problem faced by analysts and accountants. It takes raw, chaotic Excel reports (often exported from legacy ERP systems) and transforms them into clean, standardized datasets ready for Pivot Tables, Power BI, or database ingestion.

##  Key Features

* **Smart Cleaning:** Automatically detects and removes duplicate rows and 100% empty lines (garbage collection).
* **Whitespace Trimming:** The "Secret Sauce". Fixes invisible errors by recursively stripping leading/trailing spaces from all text columns.
* **Auto-Formatting:** Intelligent column width adjustment based on content size for perfect readability immediately after opening.
* **Safe Mode:** Uses the `openpyxl` engine to handle modern `.xlsx` files securely without data corruption.

##  Installation

1.  Clone this repository:
    ```bash
    git clone [https://github.com/GustavoSJ7/excel-cleaner.git](https://github.com/GustavoSJ7/excel-cleaner.git)
    ```
2.  Install the required dependencies:
    ```bash
    pip install -r requirements.txt
    ```

##  Usage

You can use the script directly via terminal to clean any file:

```python
from cleaner import clean_excel_file

# Usage: clean_excel_file(input_path, output_path)
clean_excel_file("raw_sales_report.xlsx", "clean_sales_report.xlsx")
