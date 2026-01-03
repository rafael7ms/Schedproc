import pandas as pd
import psycopg2
import os
from datetime import datetime
import sys

def connect_to_db():
    """Establish connection to the PostgreSQL database"""
    try:
        conn = psycopg2.connect(
            host=os.environ.get('DB_HOST', 'localhost'),
            database=os.environ.get('DB_NAME', 'roster_db'),
            user=os.environ.get('DB_USER', 'roster_user'),
            password=os.environ.get('DB_PASSWORD', 'roster_password')
        )
        return conn
    except Exception as e:
        print(f"Error connecting to database: {e}")
        return None

def parse_date(date_str):
    """Parse date string to proper date format"""
    if pd.isna(date_str) or date_str == '':
        return None
    try:
        # Handle different date formats
        if isinstance(date_str, str):
            # Try common date formats
            for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d'):
                try:
                    return datetime.strptime(str(date_str), fmt).date()
                except ValueError:
                    continue
        elif isinstance(date_str, (int, float)):
            # Handle Excel serial date numbers
            return datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(date_str) - 2).date()
        return None
    except Exception:
        return None

def check_duplicate(conn, table_name, agent_id):
    """Check if a record with the given agent_id already exists in the table"""
    if pd.isna(agent_id) or agent_id == '':
        return False
    
    try:
        cur = conn.cursor()
        cur.execute(f"SELECT 1 FROM {table_name} WHERE agent_id = %s", (str(agent_id),))
        result = cur.fetchone()
        cur.close()
        return result is not None
    except Exception as e:
        print(f"Error checking duplicate in {table_name}: {e}")
        return False

def insert_record(conn, table_name, record):
    """Insert a record into the specified table"""
    try:
        cur = conn.cursor()
        
        # Prepare the INSERT statement based on table type
        if table_name == 'attrition':
            columns = [
                'name', 'last_name', 'first_name', 'batch', 'agent_id', 'odoo_id',
                'bo_user', 'axonify', 'supervisor', 'manager', 'tier', 'shift',
                'schedule', 'department', 'role', 'type_of_attrition', 'term_date',
                'gds_ticket', 'wfm_ticket'
            ]
            
            values = [
                record.get('Name', ''),
                record.get('Last Name', ''),
                record.get('First Name', ''),
                record.get('Batch', ''),
                record.get('Agent ID', ''),
                record.get('Odoo ID', ''),
                record.get('BO User', ''),
                record.get('Axonify', ''),
                record.get('Supervisor', ''),
                record.get('Manager', ''),
                record.get('Tier', ''),
                record.get('Shift', ''),
                record.get('Schedule', ''),
                record.get('Department', ''),
                record.get('Role', ''),
                record.get('Type of Attrition', ''),
                parse_date(record.get('Term Date')),
                record.get('GDS Ticket', ''),
                record.get('WFM Ticket', '')
            ]
        else:  # For all other tables (agents, supervisors, etc.)
            columns = [
                'name', 'last_name', 'first_name', 'batch', 'agent_id', 'odoo_id',
                'bo_user', 'axonify', 'supervisor', 'manager', 'tier', 'shift',
                'schedule', 'department', 'role', 'phase_1_date', 'phase_2_date',
                'phase_3_date', 'hire_date'
            ]
            
            values = [
                record.get('Name', ''),
                record.get('Last Name', ''),
                record.get('First Name', ''),
                record.get('Batch', ''),
                record.get('Agent ID', ''),
                record.get('Odoo ID', ''),
                record.get('BO User', ''),
                record.get('Axonify', ''),
                record.get('Supervisor', ''),
                record.get('Manager', ''),
                record.get('Tier', ''),
                record.get('Shift', ''),
                record.get('Schedule', ''),
                record.get('Department', ''),
                record.get('Role', ''),
                parse_date(record.get('Phase 1 Date')),
                parse_date(record.get('Phase 2 Date')),
                parse_date(record.get('Phase 3 Date')),
                parse_date(record.get('Hire Date'))
            ]
        
        placeholders = ', '.join(['%s'] * len(columns))
        columns_str = ', '.join(columns)
        
        # Remove None values from agent_id check
        agent_id = record.get('Agent ID', '')
        if pd.isna(agent_id):
            agent_id = ''
        
        # Check for duplicates before inserting
        if check_duplicate(conn, table_name, agent_id):
            print(f"Skipping duplicate record for Agent ID: {agent_id} in {table_name}")
            cur.close()
            return False
        
        # Execute the INSERT statement
        query = f"INSERT INTO {table_name} ({columns_str}) VALUES ({placeholders})"
        cur.execute(query, values)
        conn.commit()
        cur.close()
        print(f"Inserted record for Agent ID: {agent_id} into {table_name}")
        return True
    except Exception as e:
        print(f"Error inserting record into {table_name}: {e}")
        conn.rollback()
        return False

def find_data_start(df):
    """Find the first row that contains actual column headers"""
    for idx, row in df.iterrows():
        # Count non-null values in the row
        non_null_count = row.count()
        # If row has substantial data and contains typical column names
        if non_null_count > 5:  # At least 5 non-null values
            # Check if this row contains common column names
            row_values = [str(val).lower() for val in row if pd.notnull(val)]
            common_columns = ['name', 'agent', 'id', 'role', 'shift', 'schedule', 'department']
            if any(col in ' '.join(row_values) for col in common_columns):
                return idx
    return 0  # Default to first row if no header found

def process_sheet(df, sheet_name):
    """Process a dataframe from an Excel sheet"""
    print(f"Processing sheet: {sheet_name}")
    print(f"Original number of rows: {len(df)}")
    
    # Check if column names are 'Unnamed' (indicating pandas couldn't find proper headers)
    unnamed_columns = [col for col in df.columns if str(col).startswith('Unnamed')]
    
    if len(unnamed_columns) > len(df.columns) // 2:  # If more than half are 'Unnamed'
        print("Detected 'Unnamed' columns, adjusting data structure...")
        # Use the first row as column names
        new_header = df.iloc[0]  # First row as header
        df = df[1:]  # Remove the first row from data
        df.columns = new_header  # Set the new column names
        print("Adjusted column names using first row of data")
    else:
        # Find the actual start of data (skip empty rows)
        data_start_row = find_data_start(df)
        if data_start_row > 0:
            print(f"Found data starting at row {data_start_row}")
            # Adjust DataFrame to start from the correct row
            new_header = df.iloc[data_start_row]  # Use the row as column names
            df = df[data_start_row+1:]  # Take data after the header row
            df.columns = new_header  # Set the new column names
    
    # Clean column names
    df.columns = [str(col).strip() if pd.notnull(col) else '' for col in df.columns]
    
    print(f"Adjusted number of rows: {len(df)}")
    print(f"Column names: {list(df.columns)}")
    
    # Connect to database
    conn = connect_to_db()
    if not conn:
        print("Failed to connect to database")
        return
    
    try:
        inserted_count = 0
        skipped_count = 0
        
        # Process each row
        for index, row in df.iterrows():
            # Convert row to dictionary
            record = row.to_dict()
            
            # Determine which table to insert into based on role
            role = str(record.get('Role', '')).strip()
            
            if sheet_name == 'Attrition':
                table_name = 'attrition'
                if insert_record(conn, table_name, record):
                    inserted_count += 1
                else:
                    skipped_count += 1
            else:  # Main roster sheet
                if role == 'Associate':
                    table_name = 'agents'
                elif role == 'Supervisor':
                    table_name = 'supervisors'
                elif role == 'Trainer':
                    table_name = 'trainers'
                elif role == 'Analyst':
                    table_name = 'quality_analysts'
                elif role == 'OM':
                    table_name = 'operations_managers'
                else:
                    print(f"Skipping record with unknown role: {role}")
                    skipped_count += 1
                    continue
                
                if insert_record(conn, table_name, record):
                    inserted_count += 1
                else:
                    skipped_count += 1
        
        print(f"Sheet '{sheet_name}' processing complete:")
        print(f"  - Inserted: {inserted_count} records")
        print(f"  - Skipped: {skipped_count} records (duplicates)")
        
    except Exception as e:
        print(f"Error processing sheet {sheet_name}: {e}")
    finally:
        conn.close()

def main():
    """Main function to parse Excel file and populate database"""
    # Check if Excel file is provided as argument
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    else:
        # Look for Excel files in current directory
        excel_files = [f for f in os.listdir('.') if f.endswith(('.xlsx', '.xls'))]
        if not excel_files:
            print("No Excel files found in current directory")
            return
        
        # Use the first Excel file found
        excel_file = excel_files[0]
        print(f"Using Excel file: {excel_file}")
    
    # Check if file exists
    if not os.path.exists(excel_file):
        print(f"Excel file not found: {excel_file}")
        return
    
    try:
        # Read the Excel file
        print(f"Reading Excel file: {excel_file}")
        
        # Read all sheets
        excel_data = pd.read_excel(excel_file, sheet_name=None)
        
        # Process each sheet
        for sheet_name, df in excel_data.items():
            if not df.empty:
                process_sheet(df, sheet_name)
            else:
                print(f"Sheet '{sheet_name}' is empty, skipping...")
        
        print("Excel file processing complete!")
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")

if __name__ == "__main__":
    main()
