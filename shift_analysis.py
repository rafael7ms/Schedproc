import pandas as pd
import numpy as np
from datetime import datetime, timedelta

def parse_time(time_str):
    # Parse time strings into datetime objects
    try:
        return pd.to_datetime(time_str, format='%I:%M %p')
    except:
        return pd.to_datetime(time_str)

# Read the Excel file from TimeReport directory
file_path = '../TimeReport/Login Logout - Dec 1 - 21 raw.xls'
df = pd.read_excel(file_path)

# Filter for PAN agents (assuming 'Group' column exists)
pan_df = df[df['Group'] == 'PAN']

# Sort by Name and Time
pan_df = pan_df.sort_values(by=['Name', 'Time'])

# Process each agent
results = []
for name, group in pan_df.groupby('Name'):
    # Sort events by time
    events = group.sort_values(by='Time')
    times = events['Time'].tolist()
    statuses = events['Status'].tolist()
    
    # Initialize variables
    shift_start = None
    shift_end = None
    first_break_start = None
    first_break_end = None
    lunch_start = None
    lunch_end = None
    second_break_start = None
    second_break_end = None
    
    # Find shift start (first login) and shift end (last logout)
    if times:
        shift_start = times[0]
        shift_end = times[-1]
    
    # Process breaks and lunch
    break_durations = []
    lunch_durations = []
    for i in range(len(times) - 1):
        if statuses[i] == 'Logout' and statuses[i+1] == 'Login':
            duration = (pd.to_datetime(times[i+1]) - pd.to_datetime(times[i])).total_seconds() / 60
            if 10 <= duration <= 20:  # Break (15 mins)
                break_durations.append((times[i], times[i+1], duration))
            elif 50 <= duration <= 70:  # Lunch (60 mins)
                lunch_durations.append((times[i], times[i+1], duration))
    
    # Assign break and lunch details
    if break_durations:
        first_break_start, first_break_end, first_break_duration = break_durations[0]
        if len(break_durations) > 1:
            second_break_start, second_break_end, second_break_duration = break_durations[1]
        else:
            second_break_start = second_break_end = second_break_duration = None
    else:
        first_break_start = first_break_end = first_break_duration = None
        second_break_start = second_break_end = second_break_duration = None
    
    if lunch_durations:
        lunch_start, lunch_end, lunch_duration = lunch_durations[0]
    else:
        lunch_start = lunch_end = lunch_duration = None
    
    # Calculate online duration
    total_shift_duration = (pd.to_datetime(shift_end) - pd.to_datetime(shift_start)).total_seconds() / 60
    total_break_time = sum([d[2] for d in break_durations]) if break_durations else 0
    total_lunch_time = sum([d[2] for d in lunch_durations]) if lunch_durations else 0
    online_duration = total_shift_duration - total_break_time - total_lunch_time
    
    results.append({
        'Name': name,
        'Shift Start': shift_start,
        '1st Break Start': first_break_start,
        '1st Break Duration': first_break_duration,
        'Lunch Start': lunch_start,
        'Lunch Duration': lunch_duration,
        '2nd Break Start': second_break_start,
        '2nd Break Duration': second_break_duration,
        'Shift End': shift_end,
        'Online Duration': online_duration
    })

# Create output DataFrame and save to Excel
output_df = pd.DataFrame(results)
output_df.to_excel('shift_analysis_output.xlsx', index=False)

print("Analysis complete. Results saved to shift_analysis_output.xlsx")
