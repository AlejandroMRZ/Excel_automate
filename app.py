import pandas as pd
from xlsxwriter.utility import xl_range

def main():
    # -------------- PARAMETERS --------------
    # Input Excel file name and sheet
    input_file = 'sales_january.xlsx'
    sheet_name = 'Sheet1'
    
    # Column names in the Excel file
    gender_col = 'Gender'
    product_line_col = 'Product line'
    total_col = 'Total'  # Updated to use 'Total' instead of 'Sales'
    
    # Output file name
    output_file = 'report_month.xlsx'
    
    # Chart placement on the worksheet (e.g., cell H2)
    chart_position = 'H2'
    
    # -------------- READ DATA --------------
    try:
        df = pd.read_excel(input_file, sheet_name=sheet_name)
    except Exception as e:
        print(f"Error reading '{input_file}': {e}")
        return

    # -------------- CREATE THE PIVOT TABLE --------------
    # Create a pivot table that sums totals by Gender (rows) and Product line (columns)
    pivot = pd.pivot_table(
        df,
        index=gender_col,
        columns=product_line_col,
        values=total_col,
        aggfunc='sum',
        fill_value=0
    )
    
    # -------------- COMPUTE THE COLUMN TOTALS --------------
    # Append a new row with column totals (sums for each product line)
    pivot.loc['Total'] = pivot.sum(numeric_only=True)
    
    # -------------- WRITE TO A NEW EXCEL FILE WITH A BAR CHART --------------
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Write the pivot table to a sheet named 'Pivot'
        pivot.to_excel(writer, sheet_name='Pivot')
        
        # Access the workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Pivot']
        
        # Determine the dimensions of the pivot table as written in Excel.
        # Note: Data is written with the header in row 0 and index in column A.
        n_rows, n_cols = pivot.shape  # n_rows includes the totals row
        
        header_row = 0                  # Header is in the first row (row 0)
        data_start_row = header_row + 1 # Data starts at row 1
        # Exclude the totals row from the chart data (last row of actual data)
        data_end_row = data_start_row + n_rows - 2  
        
        index_col = 0                 # The index (Gender labels) is in column A (0)
        data_start_col = index_col + 1 # Data for product lines starts in column B (1)
        
        # -------------- CREATE THE BAR CHART --------------
        chart = workbook.add_chart({'type': 'column'})
        
        # Loop over each product line (each column in the pivot table)
        for i, product in enumerate(pivot.columns):
            chart.add_series({
                'name':       ['Pivot', header_row, data_start_col + i],
                'categories': ['Pivot', data_start_row, index_col, data_end_row, index_col],
                'values':     ['Pivot', data_start_row, data_start_col + i, data_end_row, data_start_col + i],
            })
        
        # Configure the chart with a title and axis labels.
        chart.set_title({'name': 'Total Sales by Gender and Product Line'})
        chart.set_x_axis({'name': 'Gender'})
        chart.set_y_axis({'name': 'Total'})
        
        # Insert the chart into the worksheet at the specified location.
        worksheet.insert_chart(chart_position, chart)
    
    print(f"Report generated successfully: '{output_file}'.")

if __name__ == '__main__':
    main()
