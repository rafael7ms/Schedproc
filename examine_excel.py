import pandas as pd

# Read the Excel file
excel_file = 'Roster - December  2025 updated TC.xlsx'

# Get all sheet names
xls = pd.ExcelFile(excel_file)
print("Available sheets:")
for sheet in xls.sheet_names:
    print(f"  - '{sheet}'")  # Added quotes to see exact names

print("\n" + "="*50)
print("Examining Attrition sheet:")
df_attrition = pd.read_excel(excel_file, sheet_name='Attrition')
print(f"Number of rows: {len(df_attrition)}")

# Display column names for Attrition
print("Attrition sheet columns:")
for i, col in enumerate(df_attrition.columns):
    print(f"  {i}: '{col}' -> {df_attrition[col].iloc[0]}")

print("\n" + "="*50)
print("Examining Data sheet (main roster):")

# Load the Data sheet which seems to contain the main roster
# Try different header options to find the real data
for header_row in [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, None]:
    try:
        df_main = pd.read_excel(excel_file, sheet_name='Data', header=header_row)
        print(f"\nTrying header row {header_row}:")
        print(f"Number of rows: {len(df_main)}")
        
        # Display column names
        print("Columns:")
        for i, col in enumerate(df_main.columns):
            print(f"  {i}: '{col}'")
            
        # Check if this looks like the real data
        if len(df_main.columns) > 5 and 'Name' in str(df_main.columns):
            print("This might be the correct header!")
            print("\nFirst 5 rows:")
            print(df_main.head(5))
            break
    except Exception as e:
        print(f"Error with header row {header_row}: {e}")
        continue

# Also examine the sheet that might have the correct name with space
print("\n" + "="*50)
print("Examining '7MS Main Roster ' sheet (with trailing space):")
try:
    df_main_roster = pd.read_excel(excel_file, sheet_name='7MS Main Roster ')
    print("Successfully loaded '7MS Main Roster ' sheet")
    print(f"Number of rows: {len(df_main_roster)}")
    print("Columns:")
    for col in df_main_roster.columns:
        print(f"  - '{col}'")
        
    print("\nFirst 10 rows:")
    print(df_main_roster.head(10))
    
    if 'Role' in df_main_roster.columns:
        print("\nUnique roles in main roster:")
        print(df_main_roster['Role'].unique())
except Exception as e:
    print(f"Error loading '7MS Main Roster ' sheet: {e}")

# Try to find the actual data by skipping empty rows
print("\n" + "="*50)
print("Trying to read Data sheet by skipping rows:")

try:
    # Read without header first to see raw data
    df_raw = pd.read_excel(excel_file, sheet_name='Data', header=None)
    print(f"Raw data shape: {df_raw.shape}")
    
    # Find first row with substantial data
    for idx, row in df_raw.iterrows():
        non_null_count = row.count()
        if non_null_count > 5:  # At least 5 non-null values
            print(f"Row {idx} has {non_null_count} non-null values:")
            print(row)
            # Check if this row contains column names
            if any('name' in str(val).lower() for val in row if pd.notnull(val)):
                print(f"Row {idx} seems to contain column names")
                # Try reading with this row as header
                df_with_header = pd.read_excel(excel_file, sheet_name='Data', header=idx)
                print("Data with this header:")
                print(df_with_header.head())
                print("Columns:")
                for col in df_with_header.columns:
                    print(f"  - '{col}'")
                break
except Exception as e:
    print(f"Error reading raw data: {e}")
