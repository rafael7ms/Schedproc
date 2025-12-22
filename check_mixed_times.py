import pandas as pd
df = pd.read_excel('Agent Schedules (Dec 1st - 21st)_20251204_071622_clean.xlsx')

print('Mixed Shift agents:')
mixed = df[df['Shift'] == 'Mixed Shift']
print(f'Total: {len(mixed)}')
print()
print('Sample Mixed Shift agents with start/stop times:')
for idx, row in mixed.head(10).iterrows():
    print(f'  {row["Name"]:25} {row["Start"]:>8} - {row["Stop"]:>8} ({row["Date"]})')
