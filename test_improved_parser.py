import pandas as pd
from roster_parser_improved import process_sheet

# Read the Excel file
excel_file = 'Roster - December  2025 updated TC.xlsxX'  # Replace with your actual file name
excel_data = pd.read_excel(excel_file, sheet_name=None)

# Process the attrition sheet specifically
if 'Attrition' in excel_data:
    attrition_df = excel_data['Attrition']
    print("Attrition sheet before processing:")
    print(attrition_df.head(10))
    print("\nColumn names before processing:")
    print(list(attrition_df.columns))
    
    # Test the process_sheet function
    print("\n" + "="*50)
    print("Testing improved parser...")
    print("="*50)
    
    # We'll modify the process_sheet function to just print what it would do
    # instead of actually connecting to the database for this test
    print("Processing sheet: Attrition")
    print(f"Original number of rows: {len(attrition_df)}")
    
    # Check if column names are 'Unnamed' (indicating pandas couldn't find proper headers)
    unnamed_columns = [col for col in attrition_df.columns if str(col).startswith('Unnamed')]
    
    if len(unnamed_columns) > len(attrition_df.columns) // 2:  # If more than half are 'Unnamed'
        print("Detected 'Unnamed' columns, adjusting data structure...")
        # Use the first row as column names
        new_header = attrition_df.iloc[0]  # First row as header
        attrition_df = attrition_df[1:]  # Remove the first row from data
        attrition_df.columns = new_header  # Set the new column names
        print("Adjusted column names using first row of data")
    else:
        print("No adjustment needed for column names")
    
    # Clean column names
    attrition_df.columns = [str(col).strip() if pd.notnull(col) else '' for col in attrition_df.columns]
    
    print(f"Adjusted number of rows: {len(attrition_df)}")
    print(f"Column names after processing: {list(attrition_df.columns)}")
    
    print("\nFirst few rows after processing:")
    print(attrition_df.head())
else:
    print("Attrition sheet not found in the Excel file")
