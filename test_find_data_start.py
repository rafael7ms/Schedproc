import pandas as pd
from roster_parser import find_data_start

# Read the Excel file
excel_file = "Roster - December  2025 updated TC.xlsx"

# Read the Attrition sheet
df = pd.read_excel(excel_file, sheet_name='Attrition')

print("Attrition sheet data:")
print(df.head(10))

print("\nColumn names:")
print(df.columns.tolist())

print("\nTesting find_data_start function:")
data_start_row = find_data_start(df)
print(f"find_data_start returned: {data_start_row}")

# Show what the function is looking at
print("\nAnalyzing rows:")
for idx, row in df.head(10).iterrows():
    non_null_count = row.count()
    row_values = [str(val).lower() for val in row if pd.notnull(val)]
    common_columns = ['name', 'agent', 'id', 'role', 'shift', 'schedule', 'department']
    has_common_columns = any(col in ' '.join(row_values) for col in common_columns)
    
    print(f"Row {idx}: non_null_count={non_null_count}, has_common_columns={has_common_columns}")
    print(f"  Values: {row_values[:5]}...")  # Show first 5 values
