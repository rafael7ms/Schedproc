import pandas as pd
import numpy as np

# Let's examine what's actually in our test file
df = pd.read_excel('test_empty_lines.xlsx', sheet_name='Attrition', header=None)
print("Raw data from Excel file:")
print(df)
print(f"Shape: {df.shape}")

# Improved find_data_start function
def find_data_start(df):
    """Find the first row that contains actual column headers"""
    for idx, row in df.iterrows():
        # Count non-null values in the row
        non_null_count = row.count()
        print(f"Row {idx}: {non_null_count} non-null values - {row.tolist()}")
        
        # If row has substantial data (at least 3 non-null values)
        if non_null_count >= 3:
            # Check if this row contains common column names
            row_values = [str(val).lower() if pd.notnull(val) else '' for val in row]
            common_columns = ['name', 'agent', 'id', 'role', 'shift', 'schedule', 'department']
            
            # Count how many common column names are in this row
            matching_columns = sum(1 for val in row_values if any(col in val for col in common_columns))
            
            print(f"  Matching columns: {matching_columns}")
            
            # If we have at least 2 matching column names, consider this the header
            if matching_columns >= 2:
                return idx
    return 0  # Default to first row if no header found

data_start_row = find_data_start(df)
print(f"\nfind_data_start identifies row {data_start_row} as the header row")

# Let's see what's in that row
if data_start_row < len(df):
    header_row = df.iloc[data_start_row]
    print(f"Header row {data_start_row}: {header_row.tolist()}")
    
    # Set column names
    df.columns = header_row
    
    # Data rows start after header
    data_rows = df.iloc[data_start_row+1:]
    print(f"Data rows after header:")
    print(data_rows)
    
    # Clean data rows (remove rows that are all NaN)
    data_rows = data_rows.dropna(how='all')
    print(f"Cleaned data rows:")
    print(data_rows)
