import pandas as pd
from roster_parser_improved import find_data_start

# Read the test file
df = pd.read_excel('test_empty_lines.xlsx', header=None)

print("Original DataFrame:")
print(df)
print(f"Original shape: {df.shape}")

# Apply the same logic as in our improved parser
print("\n--- Parser Analysis ---")
data_start_row = find_data_start(df)
print(f"Found data starting at row {data_start_row}")

if data_start_row > 0:
    new_header = df.iloc[data_start_row]  # Use the row as column names
    df_data = df[data_start_row+1:]  # Take data after the header row
    df_data.columns = new_header  # Set the new column names
else:
    # Fallback mechanism
    df_data = df.dropna(how='all').reset_index(drop=True)
    if len(df_data) > 0:
        new_header = df_data.iloc[0]  # Use the first row as column names
        df_data = df_data[1:]  # Take data after the header row
        df_data.columns = new_header  # Set the new column names

# Clean column names
df_data.columns = [str(col).strip() if pd.notnull(col) else '' for col in df_data.columns]

# Remove any remaining completely empty rows
df_data = df_data.dropna(how='all').reset_index(drop=True)

print("\nProcessed DataFrame:")
print(df_data)
print(f"Processed shape: {df_data.shape}")
print(f"Column names: {list(df_data.columns)}")

# Show first row of data
if len(df_data) > 0:
    print(f"First row of data: {df_data.iloc[0].to_dict()}")
