# pyexcelops

> Professional Excel reporting for pandas, made simple.

`pyexcelops` is a lightweight Python utility that makes it easy to export one or multiple pandas DataFrames to a professionally formatted Excel workbook using `openpyxl`.

Instead of repeatedly writing Excel formatting code, you can use a single function to create clean, readable, and presentation-ready Excel reports.

## Features

- Export one or multiple pandas DataFrames
- Create multiple worksheets in a single workbook
- Automatically sanitize invalid Excel sheet names
- Handle duplicate sheet names
- Format headers and data
- Auto-adjust column widths
- Automatically adjust row heights for wrapped text
- Freeze header rows
- Enable Excel filters
- Convert DataFrames into Excel Tables
- Highlight selected columns
- Hide Excel grid-lines
- Wrap long cell content
- Handle empty DataFrames gracefully
- Support custom worksheet names
- Support large datasets with optimized column-width calculation

## Requirements

- Python 3.9+
- pandas
- openpyxl

Install the dependencies:

```bash
pip install pandas openpyxl
```

## Usage

### Basic Example

```python
import pandas as pd
from pyexcelops import save_formatted_excel

df = pd.DataFrame({
    "Name": ["John", "Jane", "Alex"],
    "Age": [25, 30, 28],
    "Department": ["IT", "HR", "Finance"]
})

save_formatted_excel(
    dfs=[df],
    output_path="report.xlsx"
)
```

This creates:

```text
report.xlsx
```

with the DataFrame exported to an appropriately formatted Excel worksheet.

---

## Multiple Worksheets

You can export multiple DataFrames into separate worksheets within the same workbook.

```python
import pandas as pd
from pyexcelops import save_formatted_excel

sales_df = pd.DataFrame({
    "Employee": ["John", "Jane"],
    "Sales": [12000, 15000]
})

employee_df = pd.DataFrame({
    "Employee": ["John", "Jane"],
    "Department": ["Sales", "Finance"]
})

save_formatted_excel(
    dfs=[sales_df, employee_df],
    output_path="company_report.xlsx",
    sheet_names=["Sales", "Employees"]
)
```

The resulting workbook will contain:

```text
company_report.xlsx
├── Sales
└── Employees
```

## Formatting Options

The function supports several options for controlling the appearance of the generated workbook.

Example:

```python
save_formatted_excel(
    dfs=[sales_df, employee_df],
    output_path="company_report.xlsx",
    sheet_names=["Sales", "Employees"],
    freeze_header=True,
    add_filters=True,
    as_table=True,
    show_grids=False,
    wrap_text=True,
    auto_fit_row_heights=True
)
```

### Common Options

| Option                 | Description                                  |
| ---------------------- | -------------------------------------------- |
| `sheet_names`          | Names of the worksheets                      |
| `freeze_header`        | Freezes the header row                       |
| `add_filters`          | Adds Excel filters                           |
| `as_table`             | Formats the data as an Excel Table           |
| `show_grids`       | Controls worksheet grids                         |
| `wrap_text`            | Wraps long cell content                      |
| `auto_fit_row_heights` | Adjusts row heights based on wrapped content |
| `highlight_cols_list`  | Highlights selected columns                  |

## Highlighting Columns

Specific columns can be highlighted when required.

```python
save_formatted_excel(
    dfs=[df],
    output_path="report.xlsx",
    sheet_names=["Report"],
    highlight_cols_list=[
        ["Sales", "Department"]
    ]
)
```

For multiple worksheets, provide the corresponding column lists for each DataFrame:

```python
save_formatted_excel(
    dfs=[sales_df, employee_df],
    output_path="report.xlsx",
    sheet_names=["Sales", "Employees"],
    highlight_cols_list=[
        ["Sales"],
        ["Department"]
    ]
)
```

## Excel Sheet Name Handling

Excel has restrictions on worksheet names. `pyexcelops` automatically handles common issues such as:

- Names longer than Excel's 31-character limit
- Invalid characters
- Duplicate worksheet names
- Empty worksheet names

For example:

```python
sheet_names=[
    "Very Long Worksheet Name That Exceeds Excel Limit",
    "Sales/Report",
    "Sales/Report"
]
```

will be sanitized and made unique automatically.

## Empty DataFrames

Empty DataFrames are handled without causing the export process to fail.

```python
empty_df = pd.DataFrame()

save_formatted_excel(
    dfs=[empty_df],
    output_path="empty_report.xlsx",
    sheet_names=["Empty Report"]
)
```

## Large DataFrames

For large datasets, column-width calculation does not need to inspect every value in a column. The implementation samples data to avoid unnecessary processing overhead.

This helps keep Excel generation reasonably efficient while still producing readable column widths.

## Project Structure

```text
pyexcelops/
│
├── pyexcelops.py
├── requirements.txt
├── README.md
└── LICENSE
```

## Why pyexcelops?

Creating a formatted Excel report with pandas often requires additional `openpyxl` code for:

- Workbook formatting
- Header styling
- Column sizing
- Filters
- Tables
- Freeze panes
- Text wrapping
- Row heights
- Worksheet cleanup

`pyexcelops` brings these repetitive operations together behind a simple function so that you can focus on the data rather than Excel formatting.

## Example Use Cases

`pyexcelops` can be useful for:

- Data analysis reports
- Operational reports
- Automated Excel deliverables
- Data quality reports
- Business reports
- Clinical/research data exports
- Recurring reporting workflows
- Multi-sheet Excel outputs

## License

This project is licensed under the MIT License.

See the `LICENSE` file for details.

## Author

**Srinibash Samal**

GitHub: [srinibashsamal](https://github.com/srinibashsamal)
