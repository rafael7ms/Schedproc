import pandas as pd
import psycopg2
import os
from datetime import datetime
import sys

def find_data_start(df):
    """Find the first row that contains actual column headers"""
    for idx, row in df.iterrows():
        # Count non-null values in the row
        non_null_count = row.count()
        print(f"Row {idx}: non-null count = {non_null_count}")
        # If row has substantial data and contains typical column names
        if non_null_count > 5:  # At least 5 non-null values
            # Check if this row contains common column names
            row_values = [str(val).lower() for val in row if pd.notnull(val)]
            common_columns = ['name', 'agent', 'id', 'role', 'shift', 'schedule', 'department']
            print(f"Row {idx} values: {row_values}")
            if any(col in ' '.join(row_values) for col in common_columns):
                print(f"Found header row at index {idx}")
                return idx
    print("No header row found, returning 0")
    return 0  # Default to first row if no header found

def process_sheet_debug(df, sheet_name):
    """Process a dataframe from an Excel sheet with debug output"""
    print(f"Processing sheet: {sheet_name}")
    print(f"Original number of rows: {len(df)}")
    print("First 5 rows of original data:")
    print(df.head())
    
    # Find the actual start of data (skip empty rows)
    data_start_row = find_data_start(df)
    print(f"Data start row: {data_start_row}")
    
    if data_start_row > 0:
        print(f"Found data starting at row {data_start_row}")
        # Adjust DataFrame to start from the correct row
        new_header = df.iloc[data_start_row]  # Use the row as column names
        df = df[data_start_row+1:]  # Take data after the header row
        df.columns = new_header  # Set the new column names
        print(f"After adjusting for header, DataFrame shape: {df.shape}")
    else:
        print("No header adjustment needed")
    
    # Clean column names
    df.columns = [str(col).strip() if pd.notnull(col) else '' for col in df.columns]
    
    print(f"Adjusted number of rows: {len(df)}")
    print("First 3 rows after processing:")
    print(df.head(3))
    print(f"Column names: {list(df.columns)}")

def main():
    """Main function to parse Excel file and populate database"""
    excel_file = "Roster - December  2025 updated TC.xlsx"
    sheet_name = "Attrition"
    
    # Check if file exists
    if not os.path.exists(excel_file):
        print(f"Excel file not found: {excel_file}")
        return
    
    try:
        # Read the Excel file
        print(f"Reading Excel file: {excel_file}, sheet: {sheet_name}")
        
        # Read specific sheet
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        
        # Process the sheet
        process_sheet_debug(df, sheet_name)
        
        print("Excel file processing complete!")
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
