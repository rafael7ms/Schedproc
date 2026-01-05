#!/usr/bin/env python3
"""
Script to read the 7MS Main Roster Sheet from the Roster Excel file
and add agents to the roster_db PostgreSQL database.

This script specifically targets the "7MS Main Roster" sheet and processes
only agents with Role = 'Associate' to be added to the agents table.
"""

import pandas as pd
import psycopg2
import os
import sys
import argparse
import logging
from datetime import datetime

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('add_agents_to_roster_db.log')
    ]
)
logger = logging.getLogger(__name__)

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
        logger.error(f"Error connecting to database: {e}")
        return None

def parse_date(date_str):
    """Parse date string to proper date format"""
    if pd.isna(date_str) or date_str == '':
        return None
    try:
        # Handle different date formats
        if isinstance(date_str, str):
            # Try common date formats
            for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d', '%m-%d-%Y', '%d-%m-%Y'):
                try:
                    return datetime.strptime(str(date_str), fmt).date()
                except ValueError:
                    continue
        elif isinstance(date_str, (int, float)):
            # Handle Excel serial date numbers
            return datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(date_str) - 2).date()
        return None
    except Exception as e:
        logger.warning(f"Error parsing date '{date_str}': {e}")
        return None

def check_duplicate(conn, agent_id):
    """Check if a record with the given agent_id already exists in the agents table"""
    if pd.isna(agent_id) or agent_id == '':
        return False

    try:
        cur = conn.cursor()
        cur.execute("SELECT 1 FROM agents WHERE agent_id = %s", (str(agent_id),))
        result = cur.fetchone()
        cur.close()
        return result is not None
    except Exception as e:
        logger.error(f"Error checking duplicate: {e}")
        return False

def insert_agent(conn, record):
    """Insert an agent record into the agents table"""
    try:
        cur = conn.cursor()

        # Prepare the INSERT statement
        columns = [
            'name', 'last_name', 'first_name', 'batch', 'agent_id', 'odoo_id',
            'bo_user', 'axonify', 'supervisor', 'manager', 'tier', 'shift',
            'schedule', 'department', 'role', 'phase_1_date', 'phase_2_date',
            'phase_3_date', 'hire_date'
        ]

        # Extract and clean values from the record
        values = [
            str(record.get('Name', '')).strip() if pd.notna(record.get('Name')) else '',
            str(record.get('Last Name', '')).strip() if pd.notna(record.get('Last Name')) else '',
            str(record.get('First Name', '')).strip() if pd.notna(record.get('First Name')) else '',
            str(record.get('Batch', '')).strip() if pd.notna(record.get('Batch')) else '',
            str(record.get('Agent ID', '')).strip() if pd.notna(record.get('Agent ID')) else '',
            str(record.get('Odoo ID', '')).strip() if pd.notna(record.get('Odoo ID')) else '',
            str(record.get('BO User', '')).strip() if pd.notna(record.get('BO User')) else '',
            str(record.get('Axonify', '')).strip() if pd.notna(record.get('Axonify')) else '',
            str(record.get('Supervisor', '')).strip() if pd.notna(record.get('Supervisor')) else '',
            str(record.get('Manager', '')).strip() if pd.notna(record.get('Manager')) else '',
            str(record.get('Tier', '')).strip() if pd.notna(record.get('Tier')) else '',
            str(record.get('Shift', '')).strip() if pd.notna(record.get('Shift')) else '',
            str(record.get('Schedule', '')).strip() if pd.notna(record.get('Schedule')) else '',
            str(record.get('Department', '')).strip() if pd.notna(record.get('Department')) else '',
            'Associate',  # Role is always Associate for this script
            parse_date(record.get('Phase 1 Date')),
            parse_date(record.get('Phase 2 Date')),
            parse_date(record.get('Phase 3 Date')),
            parse_date(record.get('Hire Date'))
        ]

        placeholders = ', '.join(['%s'] * len(columns))
        columns_str = ', '.join(columns)

        query = f"INSERT INTO agents ({columns_str}) VALUES ({placeholders})"

        # Check for duplicates before inserting
        agent_id = str(record.get('Agent ID', '')).strip() if pd.notna(record.get('Agent ID')) else ''

        if check_duplicate(conn, agent_id):
            logger.info(f"Skipping duplicate agent: {agent_id}")
            cur.close()
            return False

        # Execute the INSERT statement
        cur.execute(query, values)
        conn.commit()
        cur.close()
        logger.info(f"Inserted agent: {agent_id}")
        return True
    except Exception as e:
        logger.error(f"Error inserting agent: {e}")
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
            common_columns = ['name', 'agent', 'id', 'role', 'shift', 'schedule', 'department', 'last', 'first']

            # Count how many common column names are in this row
            matching_columns = sum(1 for val in row_values if any(col in val for col in common_columns))

            # If we have at least 2 matching column names, consider this the header
            if matching_columns >= 2:
                return idx
    return 0  # Default to first row if no header found

def clean_column_names(df):
    """Clean and standardize column names"""
    column_mapping = {
        'agent id': 'Agent ID',
        'employee id': 'Agent ID',
        'id': 'Agent ID',
        'odoo id': 'Odoo ID',
        'odoo': 'Odoo ID',
        'bo user': 'BO User',
        'bo': 'BO User',
        'axonify': 'Axonify',
        'supervisor': 'Supervisor',
        'sup': 'Supervisor',
        'manager': 'Manager',
        'mgr': 'Manager',
        'tier': 'Tier',
        'shift': 'Shift',
        'schedule': 'Schedule',
        'dept': 'Department',
        'department': 'Department',
        'role': 'Role',
        'phase 1': 'Phase 1 Date',
        'phase 2': 'Phase 2 Date',
        'phase 3': 'Phase 3 Date',
        'hire date': 'Hire Date',
        'start date': 'Hire Date',
        'first name': 'First Name',
        'first': 'First Name',
        'given name': 'First Name',
        'last name': 'Last Name',
        'last': 'Last Name',
        'surname': 'Last Name',
        'name': 'Name',
        'full name': 'Name',
        'batch': 'Batch'
    }

    new_columns = []
    for col in df.columns:
        col_lower = str(col).lower().strip()
        # Find the best matching column name
        matched = False
        for key, value in column_mapping.items():
            if key in col_lower:
                new_columns.append(value)
                matched = True
                break
        if not matched:
            new_columns.append(str(col).strip())

    return new_columns

def process_roster_sheet(file_path):
    """Process the 7MS Main Roster sheet and add agents to database"""
    logger.info(f"Processing roster file: {file_path}")

    try:
        # Read the specific sheet
        logger.info("Reading 7MS Main Roster sheet from Excel file")
        df = pd.read_excel(file_path, sheet_name='7MS Main Roster', header=None)

        if df.empty:
            logger.error("The 7MS Main Roster sheet is empty")
            return 0

        # Find the actual start of data (skip empty rows)
        data_start_row = find_data_start(df)
        logger.info(f"Found data starting at row {data_start_row}")

        # Adjust DataFrame to start from the correct row
        if data_start_row < len(df):
            new_header = df.iloc[data_start_row]  # Use the row as column names
            df = df[data_start_row+1:]  # Take data after the header row
            df.columns = new_header  # Set the new column names

            # Reset index to ensure proper handling
            df = df.reset_index(drop=True)

            # Clean column names
            df.columns = clean_column_names(df)

            logger.info(f"Adjusted number of rows: {len(df)}")
            logger.info(f"Columns found: {list(df.columns)}")
        else:
            logger.error("No data found after header row")
            return 0

        # Connect to database
        conn = connect_to_db()
        if not conn:
            logger.error("Failed to connect to database")
            return 0

        try:
            # Initialize counters
            inserted_count = 0
            skipped_count = 0
            total_processed = 0

            # Process each row
            for index, row in df.iterrows():
                # Convert row to dictionary
                record = row.to_dict()

                # Get and check role - only process Associates
                role = str(record.get('Role', '')).strip().lower() if pd.notna(record.get('Role')) else ''
                if role not in ['associate', 'agent']:
                    continue

                total_processed += 1

                if insert_agent(conn, record):
                    inserted_count += 1
                else:
                    skipped_count += 1

            logger.info(f"Processing complete:")
            logger.info(f"  - Total processed: {total_processed}")
            logger.info(f"  - Associates inserted: {inserted_count}")
            logger.info(f"  - Skipped: {skipped_count} (duplicates or invalid data)")

            return inserted_count

        except Exception as e:
            logger.error(f"Error processing sheet: {e}")
            return 0
        finally:
            conn.close()

    except Exception as e:
        logger.error(f"Error reading Excel file: {e}")
        return 0

def main():
    """Main function to parse the 7MS Main Roster sheet and add agents to database"""
    parser = argparse.ArgumentParser(description='Add agents from 7MS Main Roster to roster_db database')
    parser.add_argument('roster_file', help='Path to the Excel roster file')
    parser.add_argument('--dry-run', action='store_true', help='Dry run - show what would be inserted without actually inserting')
    args = parser.parse_args()

    # Check if file exists
    if not os.path.exists(args.roster_file):
        logger.error(f"Excel file not found: {args.roster_file}")
        return

    if args.dry_run:
        logger.info("DRY RUN: Showing what would be processed without inserting")
        try:
            # Read the specific sheet
            df = pd.read_excel(args.roster_file, sheet_name='7MS Main Roster', header=None)

            if df.empty:
                logger.info("The 7MS Main Roster sheet is empty")
                return

            # Find the actual start of data
            data_start_row = find_data_start(df)
            if data_start_row < len(df):
                new_header = df.iloc[data_start_row]
                df = df[data_start_row+1:]
                df.columns = new_header
                df = df.reset_index(drop=True)

                # Clean column names
                df.columns = clean_column_names(df)

                # Filter for Associates only
                associates = df[df['Role'].str.strip().str.lower().isin(['associate', 'agent'])]

                logger.info(f"Found {len(associates)} associate records that would be processed")
                logger.info("Sample data:")
                logger.info(associates.head().to_string())

                # Show column mapping
                logger.info(f"Column mapping: {dict(zip(df.columns, clean_column_names(df)))}")
            return
        except Exception as e:
            logger.error(f"Error during dry run: {e}")
            return

    # Process the roster file
    inserted_count = process_roster_sheet(args.roster_file)

    if inserted_count > 0:
        logger.info(f"Successfully added {inserted_count} agents to the roster_db database")
    else:
        logger.warning("No agents were added to the database")

if __name__ == "__main__":
    main()
