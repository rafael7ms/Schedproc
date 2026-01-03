import pandas as pd

# Read the Excel file
excel_file = "Roster - December  2025 updated TC.xlsx"

# Read the Attrition sheet
df = pd.read_excel(excel_file, sheet_name='Attrition')

# Print the first 10 rows to see the structure
print("First 10 rows of Attrition sheet:")
print(df.head(10))

# Print column names
print("\nColumn names:")
print(df.columns.tolist())

# Check for empty rows at the beginning
print("\nChecking for empty rows at the beginning:")
for idx, row in df.head(10).iterrows():
    print(f"Row {idx}: {row.tolist()}")
    # Check if row is empty
    if row.isnull().all() or (row.astype(str).str.strip() == '').all():
        print(f"  -> Row {idx} is empty")
    else:
        print(f"  -> Row {idx} has data")
