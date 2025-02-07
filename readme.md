# Excel Sales Automation

## Project Overview
This project automates the process of analyzing sales data from an Excel file and generating a summarized report with a pivot table and a bar chart. The script reads raw sales data, aggregates sales by gender and product line, and outputs a new Excel file containing the summary report with a visualization.

## Features
- Reads sales data from an Excel file (`sales_january.xlsx`).
- Creates a pivot table summarizing total sales by `Gender` and `Product Line`.
- Adds a total row at the bottom of the pivot table.
- Writes the summarized data into a new Excel file (`report_month.xlsx`).
- Generates a bar chart visualizing sales distribution and inserts it into the output Excel file.
- Uses `pandas` for data manipulation and `xlsxwriter` for Excel file creation and chart generation.

## Requirements
- Python 3.x
- Required Python libraries:
  - `pandas`
  - `xlsxwriter`

## Installation
To install the necessary dependencies, run:
```bash
pip install pandas xlsxwriter
```

## Usage
1. Place the sales data Excel file (`sales_january.xlsx`) in the project directory.
2. Run the script:
   ```bash
   python app.py
   ```
3. The script will generate a new report file (`report_month.xlsx`) with a pivot table and a bar chart.
4. Open the generated file to view the summarized sales report.

## File Structure
```
Excel_automation/
│── app.py                      # Main script for data processing
│── automate_excel.ipynb        # Jupyter Notebook for interactive analysis
│── sales_january.xlsx          # Input sales data (example)
│── report_month.xlsx           # Generated report (output)
│── README.md                   # Project documentation
```

## Explanation of the Script
The script follows these steps:
1. **Read Sales Data**: Loads data from `sales_january.xlsx` using `pandas.read_excel()`.
2. **Create Pivot Table**: Groups total sales by `Gender` and `Product Line` using `pd.pivot_table()`.
3. **Compute Totals**: Appends a row with column-wise total sales.
4. **Write to Excel**: Saves the pivot table to a new Excel file (`report_month.xlsx`).
5. **Generate Chart**: Uses `xlsxwriter` to create a bar chart and inserts it into the Excel report.
6. **Save and Exit**: Writes everything to the output file and completes execution.

## Example Output
The generated Excel file (`report_month.xlsx`) contains:
- A pivot table summarizing total sales by gender and product line.
- A bar chart visualizing the summarized sales data.

## Future Enhancements
- Add support for multiple input files (process multiple months).
- Include more visualization types (e.g., pie charts, trend analysis).
- Implement a GUI for user-friendly interaction.
- Automate email notifications with the generated report.

## Author
Alejandro

## License
This project is open-source and available under the MIT License.

