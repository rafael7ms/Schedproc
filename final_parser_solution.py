import pandas as pd
import numpy as np

def find_data_start(df):
    """Find the first row that contains actual column headers"""
    for idx, row in df.iterrows():
        # Count non-null values in the row
        non_null_count = row.count()
        
        # If row has substantial data (at least 3 non-null values)
        if non_null_count >= 3:
            # Check if this row contains common column names
            row_values = [str(val).lower() if pd.notnull(val) else '' for val in row]
            common_columns = ['name', 'agent', 'id', 'role', 'shift', 'schedule', 'department']
            
            # Count how many common column names are in this row
            matching_columns = sum(1 for val in row_values if any(col in val for col in common_columns))
            
            # If we have at least 2 matching column names, consider this the header
            if matching_columns >= 2:
                return idx
    return 0  # Default to first row if no header found

def parse_attrition_sheet(file_path):
    """Parse the Attrition sheet handling empty lines before the data"""
    # Read the Excel file without headers first
    df = pd.read_excel(file_path, sheet_name='Attrition', header=None)
    print(f"Raw data shape: {df.shape}")
    
    # Find where the actual data starts
    header_row_index = find_data_start(df)
    print(f"Header row identified at index: {header_row_index}")
    
    # Get the header row
    header_row = df.iloc[header_row_index]
    print(f"Header row: {header_row.tolist()}")
    
    # Get data rows (everything after the header row)
    data_rows = df.iloc[header_row_index + 1:]
    print(f"Data rows shape: {data_rows.shape}")
    
    # Set the column names
    data_rows.columns = header_row
    
    # Remove rows that are completely empty
    data_rows = data_rows.dropna(how='all')
    
    # Remove columns that are completely empty
    data_rows = data_rows.dropna(axis=1, how='all')
    
    print(f"Final data shape: {data_rows.shape}")
    print("Final data:")
    print(data_rows)
    
    return data_rows

# Test our function
if __name__ == "__main__":
    df = parse_attrition_sheet('test_empty_lines.xlsx')
