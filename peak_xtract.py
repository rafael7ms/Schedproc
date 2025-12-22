import pandas as pd
import sys
import os
from datetime import datetime

def calculate_peak_seating_with_training(input_file):
    """
    Calculate peak seating analysis for schedules with detailed floor occupancy calculation.
    Only includes shifts from the shift_schedule list in the main table.
    """
    try:
        # Read the input file
        if input_file.endswith('.xlsx'):
            df = pd.read_excel(input_file)
        elif input_file.endswith('.csv'):
            df = pd.read_csv(input_file)
        else:
            raise ValueError("Input file must be .xlsx or .csv format")

        # Ensure required columns exist
        required_columns = ['Date', 'Start', 'Stop', 'Status']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Required column '{col}' not found in input file")

        # Define the shift schedule - only these shifts will appear in the main table
        shift_schedule = {
            '05:00': '14:00',
            '06:00': '15:00',
            '07:00': '16:00',
            '10:00': '19:00',
            '11:00': '20:00',
            '14:00': '22:00',
            '15:00': '23:00',
            '15:30': '23:30',
            '16:00': '00:00',
            '17:00': '01:00'
        }

        # Get only the shifts that are in our shift_schedule
        valid_shifts = list(shift_schedule.keys())

        # Ensure Date column is properly formatted as MM/DD/YYYY
        if 'Date' in df.columns:
            # Convert to datetime first to ensure proper parsing
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            # Format as YYYY-MM-DD string
            df['Date_Formatted'] = df['Date'].dt.strftime('%Y-%m-%d')
        else:
            print("Error: No 'Date' column found in the input data")
            return None, None

        # Get unique dates
        unique_dates = sorted(df['Date_Formatted'].dropna().unique())

        # Create peak seating DataFrame - only include shifts from our shift_schedule
        peak_seating_data = []

        # For each valid start time, count how many agents are scheduled on each date
        for start_time in valid_shifts:
            start_row = {'Shift': start_time}
            for date in unique_dates:
                # Count agents for this start time on this date (excluding errors and training)
                base_filter = (df['Date_Formatted'] == date) & (df['Start'] == start_time)

                # Further filter out agents with 'Error' or 'Training' status if Status column exists
                if 'Status' in df.columns:
                    base_filter = base_filter & (df['Status'] != 'Error') & (df['Status'].str.lower() != 'training')

                # Get base count
                count = len(df[base_filter])
                start_row[date] = count

            # Only add row if it has non-zero values
            date_values = [start_row.get(date, 0) for date in unique_dates]
            if sum(date_values) > 0:
                peak_seating_data.append(start_row)

        # Create DataFrame for shifts
        peak_seating_df = pd.DataFrame(peak_seating_data)

        # Ensure all date columns are present
        for date in unique_dates:
            if date not in peak_seating_df.columns:
                peak_seating_df[date] = 0

        # Calculate total associates scheduled each day (sum of all valid shifts)
        total_scheduled_row = {'Shift': 'Total Associates Scheduled'}
        for date in unique_dates:
            total_scheduled_row[date] = peak_seating_df[date].sum()    

        # Add total scheduled row after shift rows
        peak_seating_data.append(total_scheduled_row)

        # Create training row (sum of agents in training each day)
        training_row = {'Shift': 'Agents_In_Training'}
        for date in unique_dates:
            training_count = 0
            # Count agents with 'Training' status for this date
            try:
                # Filter for Training status and date
                training_filter = (
                    (df['Date_Formatted'] == date) &
                    (df['Status'].str.lower() == 'training')
                )
                training_count = len(df[training_filter])

            except Exception as e:
                print(f"Error processing training data for {date}: {e}")

            # Explicitly set 0 for days with no training
            training_row[date] = training_count if training_count > 0 else 0
            
             
        # Add training row to data
        peak_seating_data.append(training_row)

        # Create vacation row (sum of agents on vacation each day)
        vacation_row = {'Shift': 'Agents_On_Vacation'}
        for date in unique_dates:
            vacation_count = 0
            # Count agents with 'Vacation' status for this date (from ALL shifts)
            try:
                date_dt = pd.to_datetime(date, format='%Y-%m-%d')
                # Filter for Vacation status and date (from all shifts)
                vacation_filter = (
                    (df['Date_Formatted'] == date) &
                    (df['Status'].str.lower() == 'vacation')
                )
                vacation_count = len(df[vacation_filter])  
            except Exception as e:
                print(f"Error processing vacation data for {date}: {e}")

            # Explicitly set 0 for days with no vacation
            vacation_row[date] = vacation_count if vacation_count > 0 else 0

        # Add vacation row to data
        peak_seating_data.append(vacation_row)

        # Calculate detailed floor occupancy with 30-minute accuracy and find maximum occupancy per day
        occupancy_results = calculate_floor_occupancy_detailed(peak_seating_df, unique_dates, shift_to_row_mapping(peak_seating_df))

        # Create peak seating shift row with maximum occupancy times
        peak_shift_row = {'Shift': 'Peak Seating Shift'}
        peak_seating_numbers_row = {'Shift': 'Peak Seating'}
        for date in unique_dates:
            date_str = str(date)
            if date_str in occupancy_results:
                max_occupancy = occupancy_results[date_str]['max_occupancy']
                peak_time = occupancy_results[date_str]['peak_time']
                peak_shift_row[date] = peak_time
                peak_seating_numbers_row[date] = max_occupancy
            else:
                peak_shift_row[date] = "00:00"
                peak_seating_numbers_row[date] = 0

        # Calculate associates off each day (total scheduled - peak seating - training - vacation)
        associates_off_row = {'Shift': 'Associates Off'}
        for date in unique_dates:
            off_count = 0
            try:
                # Filter for Training status and date
    
                off_filter = (
                    (df['Date_Formatted'] == date) &
                    (df['Start'].str.lower() == 'off')
                )
                off_count = len(df[off_filter])
            except Exception as e:
                print(f"Error processing OFF data for {date}: {e}") 
            associates_off_row[date] = off_count if off_count > 0 else 0 

        # Add associates off row to data
        peak_seating_data.append(associates_off_row)

        # Add peak seating shift row and peak seating numbers row to data
        peak_seating_data.append(peak_shift_row)
        peak_seating_data.append(peak_seating_numbers_row)

        # Create final DataFrame
        peak_seating_df = pd.DataFrame(peak_seating_data)

        # Reorder columns
        cols = ['Shift'] + [str(date) for date in unique_dates]
        peak_seating_df = peak_seating_df[cols]

        # Generate output filename with timestamp
        base_name = os.path.splitext(os.path.basename(input_file))[0].split('_')[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"{base_name}_{timestamp}_peak_seating.xlsx"

        return peak_seating_df, output_file

    except Exception as e:
        print(f"Error in calculate_peak_seating_with_training: {e}")
        raise

def shift_to_row_mapping(peak_seating_df):
    """Create mapping from shift times to row indices"""
    shift_mapping = {}
    for idx, shift_name in enumerate(peak_seating_df['Shift']):
        shift_mapping[str(shift_name)] = idx
    return shift_mapping

def time_to_minutes(time_str):
    """Convert time string to minutes from midnight"""
    if ':' not in time_str:
        return 0
    hours, minutes = map(int, time_str.split(':'))
    return hours * 60 + minutes

def is_shift_active_detailed(start_time, end_time, check_time_minutes):
    """Check if a shift is active at a specific time in minutes"""
    start_minutes = time_to_minutes(start_time)
    end_minutes = time_to_minutes(end_time)

    # Handle overnight shifts
    if end_minutes < start_minutes:
        end_minutes += 24 * 60  # Add 24 hours
        if check_time_minutes < start_minutes:
            check_time_minutes += 24 * 60  # Adjust check time if it's next day

    # Apply 30-minute overlap rule
    # If checking within 30 minutes after shift end, still count it
    if check_time_minutes <= end_minutes + 30 and check_time_minutes >= start_minutes:
        return True
    return False

def calculate_floor_occupancy_detailed(peak_seating_df, unique_dates, shift_mapping):
    """
    Calculate detailed floor occupancy with 30-minute accuracy
    """
    # Define shifts with their start and end times
    shift_schedule = {
        '05:00': '14:00',
        '06:00': '15:00',
        '07:00': '16:00',
        '10:00': '19:00',
        '11:00': '20:00',
        '14:00': '22:00',
        '15:00': '23:00',
        '15:30': '23:30',
        '16:00': '00:00',
        '17:00': '01:00'
    }

    results = {}

    # For each date, calculate occupancy at 30-minute intervals
    for date in unique_dates:
        date_str = str(date)
        occupancy_by_time = {}

        # Create time slots (every 30 minutes from 5:00 to 23:30)
        time_slots = []
        for hour in range(5, 24):
            time_slots.append(f"{hour:02d}:00")
            if hour < 23:  # Don't add 30 min for 23:30 as it's already in shifts
                time_slots.append(f"{hour:02d}:30")
        time_slots.append("23:30")

        # Calculate occupancy at each time slot
        for time_slot in time_slots:
            total_occupancy = 0
            time_minutes = time_to_minutes(time_slot)

            # Check each shift for overlap
            for shift_start, shift_end in shift_schedule.items():
                # Check if this shift exists in our data and has agents
                if shift_start in shift_mapping:
                    row_idx = shift_mapping[shift_start]
                    if row_idx < len(peak_seating_df) and date in peak_seating_df.columns:
                        shift_count = peak_seating_df.iloc[row_idx][date]
                        if shift_count > 0:
                            # Check if shift is active at this time
                            if is_shift_active_detailed(shift_start, shift_end, time_minutes):
                                total_occupancy += shift_count
            occupancy_by_time[time_slot] = total_occupancy

        # Find maximum occupancy for this date
        if occupancy_by_time:
            max_occupancy = max(occupancy_by_time.values())
            # Find the first time when maximum occupancy occurs
            peak_time = None
            for time_slot, occupancy in occupancy_by_time.items():
                if occupancy == max_occupancy:
                    peak_time = time_slot
                    break
            results[date_str] = {
                'max_occupancy': max_occupancy,
                'peak_time': peak_time if peak_time else "00:00",
                'occupancy_timeline': occupancy_by_time
            }
        else:
            results[date_str] = {
                'max_occupancy': 0,
                'peak_time': "00:00",
                'occupancy_timeline': {}
            }

    return results

def create_styled_excel_report(peak_seating_df, filename='peak_seating_report.xlsx'):
    """
    Create a styled Excel report with colors, borders, and formatting as requested
    """
    try:
        import xlsxwriter
        # Create workbook and worksheet with xlsxwriter
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet('Peak Seating Report')

        # Define formats
        header_format = workbook.add_format({
            'bg_color': '#366092',
            'font_color': 'white',
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        # Red background with white text for Peak Seating row
        red_format = workbook.add_format({
            'bg_color': '#FF0000',
            'font_color': 'white',
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        # Blue background with white text for Training row
        blue_format = workbook.add_format({
            'bg_color': '#0000FF',
            'font_color': 'white',
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        # Purple background with white text for Vacation row
        purple_format = workbook.add_format({
            'bg_color': '#800080',
            'font_color': 'white',
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        # Orange background with white text for Peak Seating Shift row
        orange_format = workbook.add_format({
            'bg_color': '#FFA500',
            'font_color': 'white',
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        # Green background with white text for Associates Off row
        green_format = workbook.add_format({
            'bg_color': '#008000',
            'font_color': 'white',
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        bold_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        # Format for first cell of regular rows (keeping color)
        first_cell_format = workbook.add_format({
            'bg_color': '#D9E1F2',  # Light blue background for first cell
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        center_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        # Write header row
        headers = list(peak_seating_df.columns)
        for col_num, header in enumerate(headers):
            worksheet.write(0, col_num, header, header_format)

        # Write data rows
        for row_num, row_data in enumerate(peak_seating_df.values, start=1):
            shift_value = row_data[0]

            # Determine format for first cell based on row content
            if shift_value == "Peak Seating":
                first_cell_fmt = red_format
                worksheet.write(row_num, 0, shift_value, first_cell_fmt)
                # Write remaining columns with red background and white text
                for col_num, cell_value in enumerate(row_data[1:], start=1):
                    worksheet.write(row_num, col_num, cell_value, red_format)
            elif shift_value == "Peak Seating Shift":
                first_cell_fmt = orange_format
                worksheet.write(row_num, 0, shift_value, first_cell_fmt)
                # Write remaining columns with orange background and white text
                for col_num, cell_value in enumerate(row_data[1:], start=1):
                    worksheet.write(row_num, col_num, cell_value, orange_format)
            elif shift_value == "Agents_In_Training":
                first_cell_fmt = blue_format
                # Change label to "In Training"
                worksheet.write(row_num, 0, "In Training", first_cell_fmt)
                # Write remaining columns with blue background and white text
                for col_num, cell_value in enumerate(row_data[1:], start=1):
                    worksheet.write(row_num, col_num, cell_value, blue_format)
            elif shift_value == "Agents_On_Vacation":
                first_cell_fmt = purple_format
                # Change label to "On Vacation"
                worksheet.write(row_num, 0, "On Vacation", first_cell_fmt)
                # Write remaining columns with purple background and white text
                for col_num, cell_value in enumerate(row_data[1:], start=1):
                    worksheet.write(row_num, col_num, cell_value, purple_format)
            elif shift_value == "Associates Off":
                first_cell_fmt = green_format
                worksheet.write(row_num, 0, shift_value, first_cell_fmt)
                # Write remaining columns with green background and white text
                for col_num, cell_value in enumerate(row_data[1:], start=1):
                    worksheet.write(row_num, col_num, cell_value, green_format)
            elif shift_value == "Total Associates Scheduled":
                first_cell_fmt = bold_format
                worksheet.write(row_num, 0, shift_value, first_cell_fmt)
                # Write remaining columns with center format
                for col_num, cell_value in enumerate(row_data[1:], start=1):
                    worksheet.write(row_num, col_num, cell_value, center_format)
            elif shift_value in ["05:00", "06:00", "07:00", "10:00", "11:00", "14:00", "15:00", "15:30", "16:00", "17:00"]:
                first_cell_fmt = first_cell_format  # Keep color for shift rows
                worksheet.write(row_num, 0, shift_value, first_cell_fmt)
                # Write remaining columns with center format
                for col_num, cell_value in enumerate(row_data[1:], start=1):
                    worksheet.write(row_num, col_num, cell_value, center_format)
            else:
                first_cell_fmt = first_cell_format  # Keep color for other rows
                worksheet.write(row_num, 0, shift_value, first_cell_fmt)
                # Write remaining columns with center format
                for col_num, cell_value in enumerate(row_data[1:], start=1):
                    worksheet.write(row_num, col_num, cell_value, center_format)

        # Auto-fit column widths
        for col_num, header in enumerate(headers):
            max_length = len(str(header))
            for row_data in peak_seating_df.values:
                cell_value = row_data[col_num]
                if len(str(cell_value)) > max_length:
                    max_length = len(str(cell_value))
            worksheet.set_column(col_num, col_num, max_length + 2)

        # Save the workbook
        workbook.close()
        print(f"Styled report with borders and colored first cells saved as {filename}")
        return filename

    except ImportError:
        print("xlsxwriter not installed. Saving as regular Excel file without styling.")
        peak_seating_df.to_excel(filename, index=False)
        print(f"Report saved as {filename}")
        return filename
    except Exception as e:
        print(f"Error creating styled report: {e}")
        # Fallback to regular Excel save
        peak_seating_df.to_excel(filename, index=False)
        print(f"Report saved as {filename}")
        return filename

if __name__ == "__main__":
    # Check if file path is provided as command line argument
    if len(sys.argv) != 2:
        print("Usage: python peak_seating_calculator.py <input_file_path>")
        print("Example: python peak_seating_calculator.py 'schedule_data.xlsx'")
        sys.exit(1)

    input_file = sys.argv[1]

    # Check if input file exists
    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' not found.")
        sys.exit(1)

    try:
        # Calculate peak seating
        df_peak, output_file = calculate_peak_seating_with_training(input_file)

        # Create styled Excel report
        styled_file = create_styled_excel_report(df_peak, output_file)

        print(f"Peak seating analysis completed. Report saved as: {styled_file}")
    except Exception as e:
        print(f"Peak seating calculation failed: {e}")
        sys.exit(1)