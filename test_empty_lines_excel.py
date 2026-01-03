import pandas as pd
import numpy as np

# Create a test DataFrame with empty lines before the actual data
data = {
    'Name': ['John Doe', 'Jane Smith', 'Bob Johnson'],
    'Agent ID': [12345, 12346, 12347],
    'Role': ['Associate', 'Supervisor', 'Associate'],
    'Department': ['Sales', 'Sales', 'Support'],
    'Shift': ['Morning', 'Morning', 'Evening']
}

# Create DataFrame with empty rows at the beginning
df_with_empty = pd.DataFrame(index=range(4))  # 3 empty rows + 1 for header
df_data = pd.DataFrame(data)
df_combined = pd.concat([df_with_empty, df_data], ignore_index=True)

# Set the column names in the correct row
df_combined.iloc[3] = ['Name', 'Agent ID', 'Role', 'Department', 'Shift']

# Save to Excel with the correct sheet name
with pd.ExcelWriter('test_empty_lines.xlsx', engine='openpyxl') as writer:
    df_combined.to_excel(writer, sheet_name='Attrition', index=False, header=False)

print("Saved test_empty_lines.xlsx with 'Attrition' sheet name")
print("Original DataFrame:")
print(df_combined)
print(f"Original shape: {df_combined.shape}")
