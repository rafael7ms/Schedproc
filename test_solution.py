import os
import sys
import pandas as pd
import sqlite3
from datetime import datetime

# Add the current directory to the Python path
sys.path.append('.')

# Import our modules
import init_db
import roster_parser

def test_database_creation():
    """Test that the database is created with correct schema"""
    print("Testing database creation...")
    
    # Remove existing database if it exists
    db_path = os.path.join('data', 'roster.db')
    if os.path.exists(db_path):
        os.remove(db_path)
    
    # Initialize the database
    init_db.init_db()
    
    # Connect to the database
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Check that all tables exist
    tables = ['agents', 'supervisors', 'trainers', 'quality_analysts', 'operations_managers', 'attrition']
    for table in tables:
        cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table}';")
        result = cursor.fetchone()
        if result:
            print(f"✓ Table '{table}' exists")
        else:
            print(f"✗ Table '{table}' missing")
    
    conn.close()

def test_duplicate_prevention():
    """Test that duplicate records are prevented"""
    print("\nTesting duplicate prevention...")
    
    # Create a test Excel file with duplicate records
    test_data = {
        'Name': ['John Doe', 'Jane Smith', 'John Doe'],  # Duplicate name
        'Last Name': ['Doe', 'Smith', 'Doe'],
        'First Name': ['John', 'Jane', 'John'],
        'Batch': ['B1', 'B2', 'B1'],
        'Agent ID': ['A001', 'A002', 'A001'],  # Duplicate ID
        'ODOO ID': ['O001', 'O002', 'O001'],   # Duplicate ID
        'BO User': ['user1', 'user2', 'user1'],
        'Axonify': ['axon1', 'axon2', 'axon1'],
        'Supervisor': ['Sup1', 'Sup2', 'Sup1'],
        'Manager': ['Mgr1', 'Mgr2', 'Mgr1'],
        'Tier': ['T1', 'T2', 'T1'],
        'Shift': ['S1', 'S2', 'S1'],
        'Schedule': ['Sch1', 'Sch2', 'Sch1'],
        'Department': ['Dept1', 'Dept2', 'Dept1'],
        'Role': ['Associate', 'Associate', 'Associate'],
        'Phase 1 Date': ['2023-01-01', '2023-01-02', '2023-01-01'],
        'Phase 2 Date': ['2023-02-01', '2023-02-02', '2023-02-01'],
        'Phase 3 Date': ['2023-03-01', '2023-03-02', '2023-03-01'],
        'Hire Date': ['2023-04-01', '2023-04-02', '2023-04-01']
    }
    
    # Create DataFrame
    df = pd.DataFrame(test_data)
    
    # Save to Excel file
    test_file = 'test_data.xlsx'
    df.to_excel(test_file, index=False)
    
    # Process the test file
    try:
        # Read the Excel file
        excel_data = pd.read_excel(test_file, sheet_name=None)
        
        # Process each sheet
        for sheet_name, df in excel_data.items():
            if not df.empty:
                roster_parser.process_sheet(df, sheet_name)
    
        # Check database for correct number of records
        db_path = os.path.join('data', 'roster.db')
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        cursor.execute("SELECT COUNT(*) FROM agents")
        count = cursor.fetchone()[0]
        
        if count == 2:  # Should only have 2 unique records, not 3
            print("✓ Duplicate prevention working correctly")
        else:
            print(f"✗ Duplicate prevention failed. Expected 2 records, found {count}")
        
        conn.close()
        
    except Exception as e:
        print(f"Error testing duplicate prevention: {e}")
    finally:
        # Clean up test file
        if os.path.exists(test_file):
            os.remove(test_file)

def main():
    """Run all tests"""
    print("Running solution tests...\n")
    
    test_database_creation()
    test_duplicate_prevention()
    
    print("\nAll tests completed!")

if __name__ == "__main__":
    main()
