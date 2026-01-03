import pandas as pd
import numpy as np

# Create a test DataFrame that simulates the issue with empty lines before data
# This simulates what might happen with the attrition sheet
test_data = [
    ['', '', '', '', ''],  # Empty row 1
    ['', '', '', '', ''],  # Empty row 2
    ['', '', '', '', ''],  # Empty row 3
    ['Name', 'Agent ID', 'Role', 'Department', 'Shift'],  # Header row
    ['John Doe', '12345', 'Associate', 'Sales', 'Morning'],
    ['Jane Smith', '12346', 'Supervisor', 'Sales', 'Morning'],
    ['Bob Johnson', '12347', 'Associate', 'Support', 'Evening']
]

# Create DataFrame
df = pd.DataFrame(test_data)

print("Original DataFrame:")
print(df)
print(f"Original shape: {df.shape}")

# Apply our fix - find the data start
def find_data_start(df):
    """Find the first row that contains actual column headers"""
    for idx, row in df.iterrows():
        # Count non-null values in the row
        non_null_count = row.count()
        # If row has substantial data and contains typical column names
        if non_null_count > 3:  # At least 3 non-null values (lowered for test)
            # Check if this row contains common column names
            row_values = [str(val).lower() for val in row if pd.notnull(val)]
            common_columns = ['name', 'agent', 'id', 'role', 'shift', 'department']
            if any(col in ' '.join(row_values) for col in common_columns):
                return idx
    return 0  # Default to first row if no header found

# Find the actual start of data (skip empty rows)
data_start_row = find_data_start(df)
print(f"\nData starts at row: {data_start_row}")

if data_start_row > 0:
    print(f"Found data starting at row {data_start_row}")
    # Adjust DataFrame to start from the correct row
    new_header = df.iloc[data_start_row]  # Use the row as column names
    df_adjusted = df[data_start_row+1:]  # Take data after the header row
    df_adjusted.columns = new_header  # Set the new column names
else:
    df_adjusted = df.copy()

# Clean column names
df_adjusted.columns = [str(col).strip() if pd.notnull(col) else '' for col in df_adjusted.columns]

print("\nAdjusted DataFrame:")
print(df_adjusted)
print(f"Adjusted shape: {df_adjusted.shape}")

# Verify the data is correctly parsed
print("\nColumn names:")
print(list(df_adjusted.columns))

print("\nData rows:")
for idx, row in df_adjusted.iterrows():
    print(f"Row {idx}: {dict(row)}")
