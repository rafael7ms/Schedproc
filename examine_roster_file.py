import pandas as pd

# Read the Excel file
file_path = "Roster - December  2025 updated TC.xlsx"
print(f"Reading Excel file: {file_path}")

try:
    # Read all sheets to see which one contains the attrition data
    excel_file = pd.ExcelFile(file_path)
    print(f"Sheet names: {excel_file.sheet_names}")
    
    # Check each sheet
    for sheet_name in excel_file.sheet_names:
        print(f"\n--- Sheet: {sheet_name} ---")
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print(f"Shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        print("First few rows:")
        print(df.head())
        
        # Check if this might be the attrition sheet
        if 'attrition' in sheet_name.lower() or 'attrition' in [str(col).lower() for col in df.columns]:
            print("*** This sheet might contain attrition data ***")
            
except Exception as e:
    print(f"Error reading Excel file: {e}")
