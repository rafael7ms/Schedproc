import pandas as pd
import numpy as np

# Let's examine what's actually in our test file
df = pd.read_excel('test_empty_lines.xlsx', sheet_name='Attrition', header=None)
print("Raw data from Excel file:")
print(df)
print(f"Shape: {df.shape}")

# Let's also check what happens when we read with header detection
df_with_header = pd.read_excel('test_empty_lines.xlsx', sheet_name='Attrition', header=0)
print("\nData with automatic header detection:")
print(df_with_header)
print(f"Shape: {df_with_header.shape}")
print(f"Columns: {df_with_header.columns.tolist()}")

# Let's check what our find_data_start function would identify
def find_data_start(df):
    """Find the first row that contains actual column headers"""
    for idx, row in df.iterrows():
        # Count non-null values in the row
        non_null_count = row.count()
        # If row has substantial data and contains typical column names
        if non_null_count > 5:  # At least 5 non-null values
            # Check if this row contains common column names
            row_values = [str(val).lower() for val in row if pd.notnull(val)]
            common_columns = ['name', 'agent', 'id', 'role', 'shift', 'schedule', 'department']
            if any(col in ' '.join(row_values) for col in common_columns):
                return idx
    return 0  # Default to first row if no header found

data_start_row = find_data_start(df)
print(f"\nfind_data_start identifies row {data_start_row} as the header row")

# Let's see what's in that row
if data_start_row < len(df):
    header_row = df.iloc[data_start_row]
    print(f"Header row {data_start_row}: {header_row.tolist()}")

# Let's see what's in the rows after that
if data_start_row + 1 < len(df):
    data_rows = df.iloc[data_start_row+1:]
    print(f"Data rows after header:")
    print(data_rows)
