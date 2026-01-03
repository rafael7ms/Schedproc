import pandas as pd
from roster_parser import process_sheet, find_data_start

# Read the test Excel file
file_path = 'test_empty_lines.xlsx'
excel_data = pd.read_excel(file_path, sheet_name=None)

print("Examining Sheet1...")
sheet_data = excel_data['Sheet1']
print("Sheet1 data:")
print(sheet_data.head(10))
print("\nColumn names:")
print(list(sheet_data.columns))

# Test the updated process_sheet function
print("\nTesting updated process_sheet function:")
process_sheet(sheet_data, 'Attrition')
