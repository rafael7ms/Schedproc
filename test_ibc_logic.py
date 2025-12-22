import pandas as pd
import seater2

df = pd.read_excel('Agent Schedules (Dec 1st - 21st)_20251204_071622_clean.xlsx')
processed_df = seater2.process_schedules(df)
seated_df = seater2.assign_seats(processed_df, 'Agent Schedules (Dec 1st - 21st)_20251204_071622_clean.xlsx')

# Check a date with < 13 IBC agents
test_date = '2025-12-14'  # Has 6 IBC agents
ibc_on_date = seated_df[(seated_df['Queue'] == 'IBC Support') & (seated_df['Date'] == test_date)]
print(f'Date: {test_date} (IBC agents: {len(ibc_on_date)})')
area_counts = ibc_on_date['Area'].value_counts().to_dict()
print(f'IBC agents by area: {area_counts}')
print()

# Check a date with exactly at threshold
test_date2 = '2025-12-01'  # Has 12 IBC agents
ibc_on_date2 = seated_df[(seated_df['Queue'] == 'IBC Support') & (seated_df['Date'] == test_date2)]
print(f'Date: {test_date2} (IBC agents: {len(ibc_on_date2)})')
area_counts2 = ibc_on_date2['Area'].value_counts().to_dict()
print(f'IBC agents by area: {area_counts2}')
