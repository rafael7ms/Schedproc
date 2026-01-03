import pandas as pd

def examine_attrition_sheet():
    # Load the Excel file
    file_path = 'Roster - December  2025 updated TC.xlsx'
    
    # Read the attrition sheet
    attrition_df = pd.read_excel(file_path, sheet_name='Attrition', header=None)
    
    print("Attrition sheet structure:")
    print(f"Total rows: {len(attrition_df)}")
    print(f"Total columns: {len(attrition_df.columns)}")
    print("\nFirst 20 rows:")
    print(attrition_df.head(20))
    
    # Find the first row with actual data
    print("\nFinding first row with actual data...")
    for i, row in attrition_df.iterrows():
        # Check if row has any non-null values
        if row.notna().any():
            print(f"Row {i} has data: {row.values}")
            # Check if it's a header row
            if any(str(cell).lower() in ['name', 'agent', 'id', 'role'] for cell in row if pd.notna(cell)):
                print(f"Row {i} appears to be a header row")
            else:
                print(f"Row {i} appears to be data row")
        if i >= 30:  # Limit to first 30 rows
            break
    
    # Check for completely empty rows
    print("\nChecking for empty rows...")
    empty_rows = []
    for i, row in attrition_df.iterrows():
        if not row.notna().any():
            empty_rows.append(i)
        if i >= 50:  # Limit to first 50 rows
            break
    
    print(f"Empty rows in first 50 rows: {empty_rows}")
    
    # Try to find where actual data starts
    print("\nTrying to find where actual data starts...")
    data_start_row = None
    for i, row in attrition_df.iterrows():
        # Skip if row is completely empty
        if not row.notna().any():
            continue
        
        # Check if this looks like a header row
        non_null_values = [str(cell) for cell in row if pd.notna(cell)]
        if any(val.lower() in ['name', 'agent', 'id', 'role', 'agent id', 'agent_id'] for val in non_null_values):
            print(f"Found header row at index {i}: {non_null_values}")
            # Data should start after header
            data_start_row = i + 1
            break
        elif len(non_null_values) >= 2:  # If row has at least 2 non-null values, might be data
            print(f"Potential data row at index {i}: {non_null_values}")
            data_start_row = i
            break
    
    if data_start_row is not None:
        print(f"\nSuggesting data starts at row {data_start_row}")
        print("First few data rows:")
        for j in range(data_start_row, min(data_start_row + 5, len(attrition_df))):
            row_data = attrition_df.iloc[j]
            non_null_values = [cell for cell in row_data if pd.notna(cell)]
            print(f"Row {j}: {non_null_values}")
    else:
        print("Could not determine where data starts")

if __name__ == "__main__":
    examine_attrition_sheet()
