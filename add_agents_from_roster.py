#!/usr/bin/env python3
"""
Script to read the 7MS Main Roster sheet from the Roster Excel file
and add agents to the roster_db database.

This script specifically targets the "7MS Main Roster " sheet and processes
only agents with Role = 'Associate' to be added to the agents table.
"""

import pandas as pd
import psycopg2
import os
from datetime import datetime
import sys
import argparse

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

def check_duplicate(conn, agent_id, role):
    """Check if a record with the given agent_id already exists in the appropriate table"""
    if pd.isna(agent_id) or agent_id == '':
        return False

    try:
        cur = conn.cursor()

        # Check in the appropriate table based on role
        if role == 'Associate':
            cur.execute("SELECT 1 FROM agents WHERE agent_id = %s", (str(agent_id),))
        elif role == 'Supervisor':
            cur.execute("SELECT 1 FROM supervisors WHERE agent_id = %s", (str(agent_id),))
        elif role == 'Trainer':
            cur.execute("SELECT 1 FROM trainers WHERE agent_id = %s", (str(agent_id),))
        elif role == 'QA':
            cur.execute("SELECT 1 FROM quality_analysts WHERE agent_id = %s", (str(agent_id),))
        elif role == 'OM':
            cur.execute("SELECT 1 FROM operations_managers WHERE agent_id = %s", (str(agent_id),))
        else:
            # For unknown roles, check all tables
            cur.execute("""
                SELECT 1 FROM agents WHERE agent_id = %s
                UNION ALL SELECT 1 FROM supervisors WHERE agent_id = %s
                UNION ALL SELECT 1 FROM trainers WHERE agent_id = %s
                UNION ALL SELECT 1 FROM quality_analysts WHERE agent_id = %s
                UNION ALL SELECT 1 FROM operations_managers WHERE agent_id = %s
            """, (str(agent_id), str(agent_id), str(agent_id), str(agent_id), str(agent_id)))

        result = cur.fetchone()
        cur.close()
        return result is not None
    except Exception as e:
        print(f"Error checking duplicate: {e}")
        return False

def insert_agent(conn, record, role):
    """Insert an agent record into the appropriate table based on role"""
    try:
        cur = conn.cursor()

        # Prepare the INSERT statement based on role
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
            role,  # Use the standardized role
            parse_date(record.get('Phase 1 Date')),
            parse_date(record.get('Phase 2 Date')),
            parse_date(record.get('Phase 3 Date')),
            parse_date(record.get('Hire Date'))
        ]

        placeholders = ', '.join(['%s'] * len(columns))

        # Determine table and standardized role based on input role
        if role == 'Associate':
            table_name = 'agents'
            columns_str = ', '.join(columns)
            query = f"INSERT INTO {table_name} ({columns_str}) VALUES ({placeholders})"
        elif role == 'Supervisor':
            table_name = 'supervisors'
            columns_str = ', '.join(columns)
            query = f"INSERT INTO {table_name} ({columns_str}) VALUES ({placeholders})"
        elif role == 'Trainer':
            table_name = 'trainers'
            columns_str = ', '.join(columns)
            query = f"INSERT INTO {table_name} ({columns_str}) VALUES ({placeholders})"
        elif role == 'QA':
            table_name = 'quality_analysts'
            columns_str = ', '.join(columns)
            query = f"INSERT INTO {table_name} ({columns_str}) VALUES ({placeholders})"
        elif role == 'OM':
            table_name = 'operations_managers'
            columns_str = ', '.join(columns)
            query = f"INSERT INTO {table_name} ({columns_str}) VALUES ({placeholders})"
        else:
            print(f"Unknown role: {role}, skipping record")
            cur.close()
            return False

        # Check for duplicates before inserting
        agent_id = record.get('Agent ID', '')
        if pd.isna(agent_id):
            agent_id = ''

        if check_duplicate(conn, agent_id, role):
            print(f"Skipping duplicate {role}: {agent_id}")
            cur.close()
            return False

        # Execute the INSERT statement
        cur.execute(query, values)
        conn.commit()
        cur.close()
        print(f"Inserted {role}: {agent_id}")
        return True
    except Exception as e:
        print(f"Error inserting {role}: {e}")
        conn.rollback()
        return False

def find_data_start(df):
    """Find the first row that contains actual column headers"""
    for idx, row in df.iterrows():
        # Count non-null values in the row
        non_null_count = row.count()

        # If row has substantial data (at least 3 non-null values)
        if non_null_count >= 3:
            # Check if this row contains common column names
            row_values = [str(val).lower() if pd.notnull(val) else '' for val in row]
            common_columns = ['name', 'agent', 'id', 'role', 'shift', 'schedule', 'department']

            # Count how many common column names are in this row
            matching_columns = sum(1 for val in row_values if any(col in val for col in common_columns))

            # If we have at least 2 matching column names, consider this the header
            if matching_columns >= 2:
                return idx
    return 0  # Default to first row if no header found

def map_role_to_standard(role):
    """Map various role names to standardized role names"""
    if pd.isna(role) or role == '':
        return None

    role = str(role).strip().lower()

    # Map various role names to standardized roles
    if role in ['associate', 'agent']:
        return 'Associate'
    elif role in ['supervisor', 'sup']:
        return 'Supervisor'
    elif role in ['trainer', 'train']:
        return 'Trainer'
    elif role in ['qa', 'analyst', 'quality analyst', 'quality assurance']:
        return 'QA'
    elif role in ['om', 'operations manager', 'operations']:
        return 'OM'
    elif role in ['receptionist', 'reception']:
        return None  # Skip receptionists
    else:
        print(f"Unknown role: '{role}', skipping")
        return None

def process_roster_sheet(df, sheet_name):
    """Process the 7MS Main Roster sheet and add agents to database"""
    print(f"Processing sheet: {sheet_name}")
    print(f"Original number of rows: {len(df)}")

    # Find the actual start of data (skip empty rows)
    data_start_row = find_data_start(df)
    print(f"Found data starting at row {data_start_row}")

    # Adjust DataFrame to start from the correct row
    if data_start_row < len(df):
        new_header = df.iloc[data_start_row]  # Use the row as column names
        df = df[data_start_row+1:]  # Take data after the header row
        df.columns = new_header  # Set the new column names

        # Reset index to ensure proper handling
        df = df.reset_index(drop=True)

        # Clean column names
        df.columns = [str(col).strip() if pd.notnull(col) else '' for col in df.columns]

        print(f"Adjusted number of rows: {len(df)}")
    else:
        print("No data found after header row, skipping sheet...")
        return 0, 0, 0, 0, 0, 0

    # Connect to database
    conn = connect_to_db()
    if not conn:
        print("Failed to connect to database")
        return 0, 0, 0, 0, 0, 0

    try:
        # Initialize counters for each role
        inserted_counts = {
            'Associate': 0,
            'Supervisor': 0,
            'Trainer': 0,
            'QA': 0,
            'OM': 0
        }
        skipped_counts = {
            'Associate': 0,
            'Supervisor': 0,
            'Trainer': 0,
            'QA': 0,
            'OM': 0
        }
        total_processed = 0

        # Process each row
        for index, row in df.iterrows():
            # Convert row to dictionary
            record = row.to_dict()

            # Get and map role
            role = str(record.get('Role', '')).strip()
            standardized_role = map_role_to_standard(role)

            # Skip if no valid role
            if not standardized_role:
                continue

            total_processed += 1

            if insert_agent(conn, record, standardized_role):
                inserted_counts[standardized_role] += 1
            else:
                skipped_counts[standardized_role] += 1

        print(f"Sheet '{sheet_name}' processing complete:")
        print(f"  - Total processed: {total_processed}")
        print(f"  - Associates: {inserted_counts['Associate']} inserted, {skipped_counts['Associate']} skipped")
        print(f"  - Supervisors: {inserted_counts['Supervisor']} inserted, {skipped_counts['Supervisor']} skipped")
        print(f"  - Trainers: {inserted_counts['Trainer']} inserted, {skipped_counts['Trainer']} skipped")
        print(f"  - QA/Analysts: {inserted_counts['QA']} inserted, {skipped_counts['QA']} skipped")
        print(f"  - OMs: {inserted_counts['OM']} inserted, {skipped_counts['OM']} skipped")

        return (
            inserted_counts['Associate'],
            inserted_counts['Supervisor'],
            inserted_counts['Trainer'],
            inserted_counts['QA'],
            inserted_counts['OM'],
            sum(skipped_counts.values())
        )

    except Exception as e:
        print(f"Error processing sheet {sheet_name}: {e}")
        return 0, 0, 0, 0, 0, 0
    finally:
        conn.close()

def main():
    """Main function to parse the 7MS Main Roster sheet and add all roles to database"""
    parser = argparse.ArgumentParser(description='Add agents from 7MS Main Roster to database')
    parser.add_argument('--file', help='Path to the Excel roster file', default=None)
    parser.add_argument('--dry-run', action='store_true', help='Dry run - show what would be inserted without actually inserting')
    args = parser.parse_args()

    # Determine the Excel file to use
    if args.file:
        excel_file = args.file
    else:
        # Look for Excel files in current directory
        excel_files = [f for f in os.listdir('.') if f.endswith(('.xlsx', '.xls')) and 'roster' in f.lower()]
        if not excel_files:
            print("No roster Excel files found in current directory")
            return

        # Use the first matching Excel file found
        excel_file = excel_files[0]
        print(f"Using Excel file: {excel_file}")

    # Check if file exists
    if not os.path.exists(excel_file):
        print(f"Excel file not found: {excel_file}")
        return

    try:
        # Read the specific sheet
        print(f"Reading 7MS Main Roster sheet from: {excel_file}")
        df = pd.read_excel(excel_file, sheet_name='7MS Main Roster ', header=None)

        if df.empty:
            print("The 7MS Main Roster sheet is empty")
            return

        # Process the sheet
        if args.dry_run:
            print("DRY RUN: Would process the following data:")
            # Just show sample data for dry run
            data_start_row = find_data_start(df)
            if data_start_row < len(df):
                new_header = df.iloc[data_start_row]
                df_sample = df[data_start_row+1:].copy()
                df_sample.columns = new_header
                df_sample = df_sample.reset_index(drop=True)

                # Show sample data
                print(f"Found {len(df_sample)} rows of data")
                print("Sample data:")
                print(df_sample.head())

                # Count different roles
                roles = df_sample['Role'].str.strip().value_counts()
                print("Roles that would be processed:")
                for role, count in roles.items():
                    print(f"  - {role}: {count}")
            return

        # Process all roles
        associate_count, supervisor_count, trainer_count, qa_count, om_count, skipped_count = process_roster_sheet(df, '7MS Main Roster ')

        print(f"\nProcessing complete:")
        print(f"  - Associates: {associate_count} inserted")
        print(f"  - Supervisors: {supervisor_count} inserted")
        print(f"  - Trainers: {trainer_count} inserted")
        print(f"  - QA/Analysts: {qa_count} inserted")
        print(f"  - Operations Managers: {om_count} inserted")
        print(f"  - Total skipped: {skipped_count} records (duplicates or unknown roles)")

    except Exception as e:
        print(f"Error reading Excel file: {e}")

if __name__ == "__main__":
    main()
