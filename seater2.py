import argparse
import os
import sys
from datetime import datetime, time
import pandas as pd
from collections import defaultdict
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import peak_xtract as pk # Import for peak seating analysis

# Define seat areas and ranges
SEAT_AREAS = {
    'OPS1-A': list(range(1, 13)),
    'OPS1-B': list(range(13, 19)),
    'OPS1-C': list(range(19, 25)),
    'OPS1-D': list(range(25, 40)),
    'OPS1-E': list(range(40, 48)),
    'OPS2': list(range(48, 61)),  # Seat 61 will be reserved
    'OPS3': list(range(62, 70)),
    'TRN': list(range(71, 101))
}

# Define shift categories
MORNING_SHIFTS = ['5AM', '6AM', '7AM']
NIGHT_SHIFTS = ['3PM', '3:30PM', '4PM', '5PM']
MIXED_SHIFTS = ['10AM', '11AM', '2PM']

# Define shift time ranges
SHIFT_TIMES = {
    '5AM': (time(5, 0), time(14, 0)),
    '6AM': (time(6, 0), time(15, 0)),
    '7AM': (time(7, 0), time(16, 0)),
    '10AM': (time(10, 0), time(19, 0)),
    '11AM': (time(11, 0), time(20, 0)),
    '2PM': (time(14, 0), time(22, 0)),
    '3PM': (time(15, 0), time(23, 0)),
    '3:30PM': (time(15, 30), time(23, 30)),
    '4PM': (time(16, 0), time(0, 0)),  # Next day
    '5PM': (time(17, 0), time(1, 0))   # Next day
}

def generate_output_filename(input_path):
    """Generate output filename with prefix and timestamp."""
    basename = os.path.basename(input_path)
    prefix = basename.split('_')[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{prefix}_{timestamp}_seating_arrangement.xlsx"


def parse_time(time_str):
    """Parse time values, handling different formats and special cases."""
    if pd.isna(time_str) or time_str == "OFF":
        return None
    
    time_str = str(time_str).strip()
    
    try:
        if len(time_str) == 5:  # HH:MM
            return datetime.strptime(time_str, "%H:%M").time()
        elif len(time_str) == 8:  # HH:MM:SS
            return datetime.strptime(time_str, "%H:%M:%S").time()
        else:
            return None
    except ValueError:
        return None


def should_ignore_agent(status, start_time):
    """Determine if an agent should be ignored based on status or start time."""
    if status in ["Vacation", "Training"]:
        return True
    if start_time == "OFF" or pd.isna(start_time):
        return True
    return False


def process_schedules(df):
    """Process agent schedules and add parsed time columns."""
    df = df.copy()
    
    # Parse start and stop times
    df['Start_Time'] = df['Start'].apply(parse_time)
    df['Stop_Time'] = df['Stop'].apply(parse_time)
    
    # Filter out ignored agents
    df = df[~df.apply(lambda row: should_ignore_agent(row['Status'], row['Start']), axis=1)]
    
    return df


def is_overlapping(start1, end1, start2, end2):
    """Check if two time ranges overlap, handling overnight shifts.
    
    Returns True if ranges overlap, False if they don't.
    Examples:
    - 5AM-2PM (05:00-14:00) doesn't overlap with 3PM-11PM (15:00-23:00) -> False
    - 5AM-2PM (05:00-14:00) overlaps with 1PM-9PM (13:00-21:00) -> True
    - 4PM-12AM (16:00-00:00) overlaps with 11PM-7AM (23:00-07:00) -> True
    """
    # Handle case where both are overnight shifts
    if end1 < start1 and end2 < start2:
        # Both overnight: check if ranges overlap across midnight
        return not (end2 < start1 and start2 > end1)
    # Handle case where only first shift is overnight
    elif end1 < start1:
        # First is overnight (e.g., 4PM-12AM)
        # Check if start2 is in range [start1, 24:00) or in range [00:00, end1)
        return start2 >= start1 or start2 < end1 or end2 >= start1 or end2 < end1
    # Handle case where only second shift is overnight
    elif end2 < start2:
        # Second is overnight (e.g., 4PM-12AM)
        return start1 >= start2 or start1 < end2 or end1 >= start2 or end1 < end2
    # Handle normal day shifts
    else:
        return start1 < end2 and start2 < end1


def categorize_shift(start_time, stop_time):
    """Categorize shift as morning or night based on start time."""
    
    # Morning shifts: 5am to 11am
    if time(5, 0) <= start_time <= time(11, 0):
        return 'morning'
    # Night shifts: 2pm onward
    elif start_time >= time(14, 0):
        return 'night'
    else:
        return 'other'

def get_batch_priority(batch):
    """Get priority for batch (lower number = higher priority)."""
    if batch == 'DH':
        return 0  # Highest priority
    try:
        return int(batch[1:])  # B6 -> 6
    except:
        return 999  # Lowest priority for invalid batches


def assign_seats_single_day(date_agents: pd.DataFrame, date: str, nesting_locations: dict,
                          input_file: str, peak_agents: int, peak_hour: str) -> pd.DataFrame:
    """Assign seats to agents for a single date following all the rules."""
    date_agents = date_agents.copy()
    date_agents['Seat'] = None
    date_agents['Area'] = None
    date_agents['Shift_Category'] = date_agents.apply(
        lambda row: categorize_shift(row['Start_Time'], row['Stop_Time']), axis=1)
    date_agents['Batch_Priority'] = date_agents['Batch'].apply(get_batch_priority)

    # Create a list of all seats (excluding reserved seats 61 and 70)
    all_seats = []
    for area, seats in SEAT_AREAS.items():
        for seat in seats:
            if seat not in [61, 70]:  # Exclude non-existent seats
                all_seats.append((area, seat))

    # Track seat assignments for this date
    seat_assignments = {}  # {(seat): [(shift_cat, agent_id, is_reusable), ...]}

    # Adjust available seats based on total agents
    total_agents = len(date_agents)
    available_seats = all_seats.copy()
    if total_agents <= 66:
        available_seats = [s for s in available_seats if s[1] not in [46, 47]]
    allow_trn = total_agents > 68 and not any(date_agents['Status'] == 'Training')

    # Define preferred areas for IBC Support
    ibc_agents = date_agents[date_agents['Queue'] == 'IBC Support']
    ibc_agents_count = len(ibc_agents)

    # Calculate how many IBC agents should go to OPS3
    ibc_ops3_allocation = 0
    if ibc_agents_count > 13:
        if ibc_agents_count <= 16:
            ibc_ops3_allocation = max(3, ibc_agents_count - 13)
        else:
            ibc_ops3_allocation = ibc_agents_count - 13
        preferred_ibc = ['OPS2', 'OPS3']
    else:
        preferred_ibc = ['OPS2']

    available_seats_ibc = [s for s in available_seats if s[0] in preferred_ibc]
    available_seats_ibc = sorted(available_seats_ibc, key=lambda x: preferred_ibc.index(x[0]))

    # Separate OPS2 and OPS3 seats for IBC allocation
    ibc_ops2_seats = [s for s in available_seats_ibc if s[0] == 'OPS2']
    ibc_ops3_seats = [s for s in available_seats_ibc if s[0] == 'OPS3']

    # Separate nesting and regular agents
    if 'Nesting' in date_agents.columns:
        nesting_agents = date_agents[date_agents['Nesting'] == True]
        regular_agents = date_agents[date_agents['Nesting'] != True]
    else:
        nesting_agents = pd.DataFrame()
        regular_agents = date_agents

    # Track IBC agents assigned to OPS3
    ibc_ops3_assigned = 0

    # Assign nesting agents first (if any)
    for idx, agent in nesting_agents.iterrows():
        agent_id = agent['ID']
        if agent_id in nesting_locations:
            area, seat = nesting_locations[agent_id]
            date_agents.at[idx, 'Seat'] = seat
            date_agents.at[idx, 'Area'] = area
            if seat not in seat_assignments:
                seat_assignments[seat] = []
            # Mark nesting seats as non-reusable
            seat_assignments[seat].append((agent['Shift_Category'], agent_id, False))

    # Assign regular agents by queue priority and shift
    queue_priority = ['IBC Support', 'BNS', 'Customer Support']
    shift_priority = ['morning', 'night']

    for queue in queue_priority:
        queue_agents = regular_agents[regular_agents['Queue'] == queue]

        for shift_category in shift_priority:
            shift_agents = queue_agents[queue_agents['Shift_Category'] == shift_category]

            for idx, agent in shift_agents.iterrows():
                if pd.notna(date_agents.at[idx, 'Seat']):
                    continue  # Already assigned

                start_time = agent['Start_Time']
                stop_time = agent['Stop_Time']
                queue = agent['Queue']

                # Determine which seats to check
                if queue == 'IBC Support':
                    if ibc_ops3_assigned < ibc_ops3_allocation:
                        seats_to_check = ibc_ops3_seats + ibc_ops2_seats
                    else:
                        seats_to_check = ibc_ops2_seats + ibc_ops3_seats
                else:
                    seats_to_check = available_seats
                    if ibc_agents_count > 13 and ibc_ops3_assigned >= ibc_ops3_allocation:
                        overflow_ops3 = ibc_ops3_seats
                        seats_to_check = [s for s in seats_to_check if s[0] != 'OPS3'] + overflow_ops3

                assigned = False
                for area, seat in seats_to_check:
                    can_use = False

                    if seat not in seat_assignments:
                        # Seat is empty, can use it if allowed by area constraints
                        if queue == 'IBC Support':
                            if area == 'OPS2' and ibc_ops3_assigned < ibc_ops3_allocation and ibc_ops3_allocation > 0:
                                can_use = False
                            elif shift_category == 'night':
                                can_use = True
                            elif shift_category == 'morning':
                                can_use = True
                        elif area != 'TRN' or (area == 'TRN' and allow_trn):
                            can_use = True
                    else:
                        # Seat is occupied, check if current agent's shift doesn't overlap with existing
                        existing_entries = seat_assignments[seat]
                        can_reuse = True
                        for existing_shift_cat, existing_agent_id, is_reusable in existing_entries:
                            existing_agent = date_agents[date_agents['ID'] == existing_agent_id].iloc[0]
                            existing_start = existing_agent['Start_Time']
                            existing_stop = existing_agent['Stop_Time']

                            # Check if shifts overlap
                            if is_overlapping(start_time, stop_time, existing_start, existing_stop):
                                can_reuse = False
                                break

                            # Additional IBC constraint
                            existing_queue = existing_agent['Queue']
                            if queue == 'IBC Support' and existing_queue != 'IBC Support':
                                if is_overlapping(start_time, stop_time, existing_start, existing_stop):
                                    can_reuse = False
                                    break

                        # Special handling for night shifts reusing morning seats
                        if can_reuse and shift_category == 'night':
                            # Only allow reuse if morning shift ends before 4PM
                            for existing_shift_cat, existing_agent_id, is_reusable in existing_entries:
                                if existing_shift_cat == 'morning':
                                    existing_agent = date_agents[date_agents['ID'] == existing_agent_id].iloc[0]
                                    if existing_agent['Stop_Time'] >= time(16, 0):  # 4PM
                                        can_reuse = False
                                        break

                        if can_reuse and queue == 'IBC Support' and ibc_agents_count > 13:
                            if area == 'OPS2' and ibc_ops3_assigned < ibc_ops3_allocation:
                                can_reuse = False

                        if can_reuse:
                            if shift_category == 'morning':
                                can_use = True
                            elif shift_category == 'night':
                                can_use = True

                    if can_use:
                        date_agents.at[idx, 'Seat'] = seat
                        date_agents.at[idx, 'Area'] = area
                        if seat not in seat_assignments:
                            seat_assignments[seat] = []

                        # Mark if morning seat can be reused at night (ends before 4PM)
                        is_reusable = (shift_category == 'morning' and stop_time < time(16, 0))
                        seat_assignments[seat].append((shift_category, agent['ID'], is_reusable))

                        if queue == 'IBC Support' and area == 'OPS3':
                            ibc_ops3_assigned += 1

                        assigned = True
                        break

                # If not assigned to preferred area, try other areas
                if not assigned and queue != 'IBC Support':
                    remaining_seats = [s for s in available_seats if s not in seats_to_check]
                    for area, seat in remaining_seats:
                        can_use_fallback = False

                        if seat not in seat_assignments:
                            if area != 'TRN' or (area == 'TRN' and allow_trn):
                                can_use_fallback = True
                        else:
                            existing_entries = seat_assignments[seat]
                            can_reuse = True

                            for existing_shift_cat, existing_agent_id, is_reusable in existing_entries:
                                existing_agent = date_agents[date_agents['ID'] == existing_agent_id].iloc[0]
                                existing_start = existing_agent['Start_Time']
                                existing_stop = existing_agent['Stop_Time']

                                if is_overlapping(start_time, stop_time, existing_start, existing_stop):
                                    can_reuse = False
                                    break

                            if can_reuse and (area != 'TRN' or (area == 'TRN' and allow_trn)):
                                can_use_fallback = True

                        if can_use_fallback:
                            date_agents.at[idx, 'Seat'] = seat
                            date_agents.at[idx, 'Area'] = area
                            if seat not in seat_assignments:
                                seat_assignments[seat] = []
                            is_reusable = (shift_category == 'morning' and stop_time < time(16, 0))
                            seat_assignments[seat].append((shift_category, agent['ID'], is_reusable))
                            assigned = True
                            break

    return date_agents

def assign_seats(df: pd.DataFrame, input_file: str) -> pd.DataFrame:
    """Assign seats to agents by processing each date separately and aggregating results."""
    df = df.copy()

    # Create Nesting column if it doesn't exist
    if 'Nesting' not in df.columns:
        df['Nesting'] = df['Status'] == 'Nesting'

    # First, assign nesting agents across all dates for consistency
    nesting_locations = {}
    if 'Nesting' in df.columns:
        nesting_agents_all = df[df['Nesting'] == True]
        if not nesting_agents_all.empty:
            total_nesting = len(nesting_agents_all['ID'].unique())
            nesting_area = 'OPS1-A' if total_nesting <= 12 else 'OPS1-D'
            nesting_seats = SEAT_AREAS[nesting_area].copy()
            nesting_seats = [seat for seat in nesting_seats if seat not in [61, 70]]
            nesting_agent_ids = list(nesting_agents_all['ID'].unique())
            for i, agent_id in enumerate(nesting_agent_ids):
                if i < len(nesting_seats):
                    nesting_locations[agent_id] = (nesting_area, nesting_seats[i])

    # Get peak seating data
    peak_agents, peak_hour = pk.calculate_peak_seating_with_training(input_file)

    # Process each date separately
    daily_results = []
    for date in df['Date'].unique():
        date_agents = df[df['Date'] == date].copy()
        date_result = assign_seats_single_day(
            date_agents, str(date), nesting_locations, input_file, peak_agents, peak_hour)
        daily_results.append(date_result)

    # Combine all daily results
    final_df = pd.concat(daily_results, ignore_index=True)

    return final_df

def create_worksheet_data(df, sheet_type, morning_df=None, night_df=None):
    """Create data for specific worksheet type."""
    # Flatten all seat numbers (excluding reserved seats)
    all_seats = []
    for area, seats in SEAT_AREAS.items():
        for seat in seats:
            if seat not in [61, 70]:
                all_seats.append((area, seat))

    # Create DataFrame with Area and Station columns
    result_data = []
    for area, seat in all_seats:
        result_data.append({
            'Area': area,
            'Station': seat
        })
    result_df = pd.DataFrame(result_data)

    # Add date columns - ensure proper sorting
    if 'Date' in df.columns:
        # Convert dates to datetime for proper sorting
        dates_series = pd.to_datetime(df['Date'], errors='coerce')
        dates = sorted(dates_series.unique())
        dates_str = [str(d.date()) if pd.notna(d) else str(d) for d in dates]

        for date_str in dates_str:
            result_df[date_str] = ''

        if sheet_type == 'Full':
            # Build from morning and night data
            if morning_df is not None and night_df is not None:
                for _, row in result_df.iterrows():
                    seat = row['Station']
                    seat_idx = result_df[result_df['Station'] == seat].index[0]
                    for date_str in dates_str:
                        morning_vals = morning_df[morning_df['Station'] == seat][date_str]
                        night_vals = night_df[morning_df['Station'] == seat][date_str]

                        morning_name = morning_vals.values[0] if len(morning_vals) > 0 and morning_vals.values[0] != '' else ''
                        night_name = night_vals.values[0] if len(night_vals) > 0 and night_vals.values[0] != '' else ''

                        # Combine in priority: morning, night
                        parts = []
                        if morning_name:
                            parts.append(morning_name)
                        if night_name:
                            parts.append(night_name)

                        if parts:
                            result_df.loc[seat_idx, date_str] = ' / '.join(parts)
            else:
                # Fallback to building from full df if not provided
                morning_filtered = df[df['Shift_Category'] == 'morning']
                night_filtered = df[df['Shift_Category'] == 'night']

                for _, agent in morning_filtered.iterrows():
                    if pd.notna(agent['Seat']):
                        date = str(agent['Date'])
                        seat = agent['Seat']
                        seat_idx = result_df[result_df['Station'] == seat].index
                        if len(seat_idx) > 0 and date in result_df.columns:
                            if result_df.loc[seat_idx[0], date] == '':
                                result_df.loc[seat_idx[0], date] = agent['Name']
                            else:
                                result_df.loc[seat_idx[0], date] = f"{result_df.loc[seat_idx[0], date]} / {agent['Name']}"

                for _, agent in night_filtered.iterrows():
                    if pd.notna(agent['Seat']):
                        date = str(agent['Date'])
                        seat = agent['Seat']
                        seat_idx = result_df[result_df['Station'] == seat].index
                        if len(seat_idx) > 0 and date in result_df.columns:
                            current = result_df.loc[seat_idx[0], date]
                            if current == '':
                                result_df.loc[seat_idx[0], date] = agent['Name']
                            else:
                                result_df.loc[seat_idx[0], date] = f"{current} / {agent['Name']}"

        elif sheet_type == 'Morning':
            # Show all morning shifts
            filtered_df = df[df['Shift_Category'] == 'morning']
            for _, agent in filtered_df.iterrows():
                if pd.notna(agent['Seat']):
                    date = str(agent['Date'])
                    seat = agent['Seat']
                    seat_idx = result_df[result_df['Station'] == seat].index
                    if len(seat_idx) > 0 and date in result_df.columns:
                        if result_df.loc[seat_idx[0], date] == '':
                            result_df.loc[seat_idx[0], date] = agent['Name']
                        else:
                            result_df.loc[seat_idx[0], date] = f"{result_df.loc[seat_idx[0], date]} / {agent['Name']}"

        elif sheet_type == 'Night':
            # Show only night shifts
            filtered_df = df[df['Shift_Category'] == 'night']
            for _, agent in filtered_df.iterrows():
                if pd.notna(agent['Seat']):
                    date = str(agent['Date'])
                    seat = agent['Seat']
                    seat_idx = result_df[result_df['Station'] == seat].index
                    if len(seat_idx) > 0 and date in result_df.columns:
                        if result_df.loc[seat_idx[0], date] == '':
                            result_df.loc[seat_idx[0], date] = agent['Name']
                        else:
                            result_df.loc[seat_idx[0], date] = f"{result_df.loc[seat_idx[0], date]} / {agent['Name']}"

    return result_df

def save_output(df, output_path):
    """Save processed DataFrame to Excel with improved formatting and color coding."""
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Create worksheet data
            morning_df = create_worksheet_data(df, 'Morning')
            night_df = create_worksheet_data(df, 'Night')
            full_df = create_worksheet_data(df, 'Full', morning_df, night_df)

            # Write to Excel sheets
            morning_df.to_excel(writer, sheet_name='Morning Seating', index=False)
            night_df.to_excel(writer, sheet_name='Night Seating', index=False)
            full_df.to_excel(writer, sheet_name='Full Seating', index=False)

            # Add legend sheet with statistics
            legend_data = {
                'Queue Colors': ['IBC Support', 'Customer Support', 'BNS'],
                'Supervisor Colors': sorted([s for s in df['Supervisor'].unique() if pd.notna(s)]) if 'Supervisor' in df.columns else [],
                'Nesting Status': ['Nesting', 'Non-Nesting'] if 'Nesting' in df.columns else []
            }
            legend_df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in legend_data.items()]))
            legend_df.to_excel(writer, sheet_name='Legend', index=False)

            # Ensure all sheets are visible
            workbook = writer.book
            for sheet in workbook.worksheets:
                sheet.sheet_state = 'visible'  # Make all sheets visible

            workbook.save(output_path)  # [1]
    except Exception as e:
        print(f"Error saving Excel file: {e}")
        raise

def verify_agent_counts(original_df, processed_df, seated_df):
    """Verify that agent counts match between input and output by date and queue.
    
    Shows scheduled agents vs assigned agents with area breakdown to identify discrepancies.
    """
    print("\n" + "="*100)
    print("VERIFICATION REPORT: Agent Counts by Date, Queue, and Area")
    print("="*100)
    
    # Convert dates to datetime
    original_df = original_df.copy()
    original_df['Date'] = pd.to_datetime(original_df['Date'], errors='coerce')
    seated_df = seated_df.copy()
    seated_df['Date'] = pd.to_datetime(seated_df['Date'], errors='coerce')
    
    # Get unique dates
    dates = sorted(original_df['Date'].unique())
    dates = [d for d in dates if pd.notna(d)]
    
    queues = ['IBC Support', 'Customer Support', 'BNS']
    
    print(f"\n{'Date':<12} {'Queue':<20} {'Scheduled':<12} {'Assigned':<12} {'Diff':<8} {'Areas':<40}")
    print("-" * 100)
    
    total_by_queue_scheduled = {q: 0 for q in queues}
    total_by_queue_assigned = {q: 0 for q in queues}
    unassigned_agents = []
    
    for date in dates:
        date_str = str(date.date())
        
        # Get scheduled agents for this date from ORIGINAL input
        scheduled_on_date = original_df[original_df['Date'] == date]
        # Remove OFF and ignored agents
        scheduled_on_date = scheduled_on_date[
            ~scheduled_on_date.apply(lambda row: should_ignore_agent(row['Status'], row['Start']), axis=1)
        ]
        
        # Get assigned agents for this date from SEATED output
        assigned_on_date = seated_df[seated_df['Date'] == date]
        
        for queue in queues:
            scheduled_count = len(scheduled_on_date[scheduled_on_date['Queue'] == queue])
            assigned_count = len(assigned_on_date[assigned_on_date['Queue'] == queue])
            
            # Get unassigned agents in this queue for this date
            unassigned = assigned_on_date[
                (assigned_on_date['Queue'] == queue) & 
                (assigned_on_date['Seat'].isna())
            ]
            if not unassigned.empty:
                for _, agent in unassigned.iterrows():
                    unassigned_agents.append({
                        'Date': date_str,
                        'Name': agent.get('Name', 'Unknown'),
                        'Queue': queue,
                        'Supervisor': agent.get('Supervisor', 'Unknown'),
                        'Batch': agent.get('Batch', 'Unknown')
                    })
            
            difference = assigned_count - scheduled_count
            total_by_queue_scheduled[queue] += scheduled_count
            total_by_queue_assigned[queue] += assigned_count
            
            # Get area breakdown for assigned agents
            assigned_with_seats = assigned_on_date[
                (assigned_on_date['Queue'] == queue) & 
                (assigned_on_date['Seat'].notna())
            ]
            area_counts = assigned_with_seats['Area'].value_counts().to_dict() if not assigned_with_seats.empty else {}
            
            # Format areas string
            areas_str = ', '.join([f"{area}({count})" for area, count in sorted(area_counts.items())])
            if unassigned.empty or assigned_count == 0:
                areas_display = areas_str if areas_str else 'None'
            else:
                areas_display = f"{areas_str} [+{len(unassigned)} unassigned]"
            
            diff_str = f"{difference:+d}" if difference != 0 else "0"
            diff_marker = " *" if difference != 0 else ""
            
            print(f"{date_str:<12} {queue:<20} {scheduled_count:<12} {assigned_count:<12} {diff_str:<8}{diff_marker} {areas_display:<40}")
        
        print("-" * 100)
    
    # Print totals
    print(f"\n{'TOTAL':<12} {'Queue':<20} {'Scheduled':<12} {'Assigned':<12} {'Diff':<8} {'Areas':<40}")
    print("-" * 100)
    for queue in queues:
        difference = total_by_queue_assigned[queue] - total_by_queue_scheduled[queue]
        diff_str = f"{difference:+d}" if difference != 0 else "0"
        diff_marker = " *" if difference != 0 else ""
        
        # Get overall area breakdown
        all_assigned = seated_df[seated_df['Queue'] == queue]
        assigned_with_seats = all_assigned[all_assigned['Seat'].notna()]
        area_counts = assigned_with_seats['Area'].value_counts().to_dict() if not assigned_with_seats.empty else {}
        areas_str = ', '.join([f"{area}({count})" for area, count in sorted(area_counts.items())])
        
        print(f"{'TOTAL':<12} {queue:<20} {total_by_queue_scheduled[queue]:<12} {total_by_queue_assigned[queue]:<12} {diff_str:<8}{diff_marker} {areas_str:<40}")
    
    # Report unassigned agents if any
    if unassigned_agents:
        print("\n" + "="*130)
        print("UNASSIGNED AGENTS (These agents have no seat assigned):")
        print("="*130)
        print(f"{'Date':<12} {'Name':<25} {'Queue':<20} {'Area':<10} {'Supervisor':<20} {'Batch':<8} {'Start':<8} {'Stop':<8}")
        print("-" * 130)
        for agent in unassigned_agents:
            # Get additional info from seated_df
            agent_detail = seated_df[
                (seated_df['Name'] == agent['Name']) & 
                (seated_df['Date'] == agent['Date'])
            ]
            area = agent_detail['Area'].values[0] if not agent_detail.empty and pd.notna(agent_detail['Area'].values[0]) else 'None'
            start = agent_detail['Start'].values[0] if not agent_detail.empty else 'N/A'
            stop = agent_detail['Stop'].values[0] if not agent_detail.empty else 'N/A'
            
            print(f"{agent['Date']:<12} {agent['Name']:<25} {agent['Queue']:<20} {area:<10} {agent['Supervisor']:<20} {agent['Batch']:<8} {str(start):<8} {str(stop):<8}")
    
    print("="*130 + "\n")



def main():
    parser = argparse.ArgumentParser(description='Generate seating arrangement for contact center agents.')
    parser.add_argument('input_file', help='Path to the input Excel file containing agent schedules.')
    
    args = parser.parse_args()
    
    try:
        # Validate input file existence
        if not os.path.exists(args.input_file):
            print(f"Error: Input file '{args.input_file}' does not exist.")
            sys.exit(1)
        
        # Load input Excel file
        try:
            df = pd.read_excel(args.input_file)
        except Exception as e:
            print(f"Error reading input file: {e}")
            sys.exit(1)
        
        # Validate required columns
        required_columns = ['Date', 'Start', 'Stop', 'Status', 'Queue', 'Supervisor', 'Name', 'ID', 'Batch']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"Error: Missing required columns: {', '.join(missing_columns)}")
            sys.exit(1)
        
        # Create Nesting column based on Status column
        df['Nesting'] = df['Status'] == 'Nesting'
        nesting_count = (df['Nesting'] == True).sum()
        print(f"Number of Nesting agents: {nesting_count}")
        
        # Process schedules
        try:
            processed_df = process_schedules(df)
        except Exception as e:
            print(f"Error processing schedules: {e}")
            sys.exit(1)
        
        # Assign seats
        try:
            seated_df = assign_seats(processed_df, args.input_file)
        except Exception as e:
            print(f"Error assigning seats: {e}")
            sys.exit(1)
        
        # Verify agent counts
        try:
            verify_agent_counts(df, processed_df, seated_df)
        except Exception as e:
            print(f"Warning: Could not verify agent counts: {e}")
        
        # Generate output filename
        output_filename = generate_output_filename(args.input_file)
        
        # Save output
        try:
            save_output(seated_df, output_filename)
            print(f"Seating arrangement saved to '{output_filename}'")
        except Exception as e:
            print(f"Error saving output: {e}")
            sys.exit(1)
    
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        sys.exit(0)
    except Exception as e:
        print(f"Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()