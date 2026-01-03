import pandas as pd
from roster_parser_improved import find_data_start, process_sheet

# Read the Excel file
excel_file = 'Roster - December  2025 updated TC.xlsx'
df = pd.read_excel(excel_file, sheet_name='Attrition', header=None)

print("Testing the improved parser with the Attrition sheet...")
print(f"Total rows in raw data: {len(df)}")
print(f"First row is empty: {df.iloc[0].isnull().all()}")
print(f"Second row (should be headers): {df.iloc[1].tolist()}")

# Test the find_data_start function
data_start_row = find_data_start(df)
print(f"find_data_start returned: {data_start_row}")

# Test processing the sheet
print("\nTesting process_sheet function:")
# We'll create a mock process_sheet function that just shows what it would do
# since we don't want to actually connect to the database in this test

# Simulate what process_sheet would do
if data_start_row > 0:
    print(f"Would adjust DataFrame to start at row {data_start_row}")
    new_header = df.iloc[data_start_row]  # Use the row as column names
    test_df = df[data_start_row+1:]  # Take data after the header row
    test_df.columns = new_header  # Set the new column names
elif data_start_row == 0 and (df.iloc[0].isnull().all() or (df.iloc[0].astype(str).str.strip() == '').all()):
    print("First row is empty, checking for header row...")
    # Look for the first non-empty row that could be headers
    for idx in range(1, min(10, len(df))):  # Check up to 10 rows
        row = df.iloc[idx]
        if not (row.isnull().all() or (row.astype(str).str.strip() == '').all()):
            # Check if this row contains common column names
            non_null_count = row.count()
            if non_null_count > 5:  # At least 5 non-null values
                row_values = [str(val).lower() for val in row if pd.notnull(val)]
                common_columns = ['name', 'agent', 'id', 'role', 'shift', 'schedule', 'department']
                if any(col in ' '.join(row_values) for col in common_columns):
                    print(f"Found header row at index {idx}")
                    new_header = df.iloc[idx]  # Use the row as column names
                    test_df = df[idx+1:]  # Take data after the header row
                    test_df.columns = new_header  # Set the new column names
                    break

# Show the column names we would get
print(f"Column names would be: {list(test_df.columns)}")
print(f"Number of data rows: {len(test_df)}")
print("First few rows of data:")
print(test_df.head())
