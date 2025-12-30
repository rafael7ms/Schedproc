#!/usr/bin/env python3
"""
Verification script for the Roster Database Solution
"""

import sqlite3
import os

def verify_database():
    """Verify that the database was created with all tables"""
    db_path = os.path.join("data", "roster.db")
    
    if not os.path.exists(db_path):
        print("ERROR: Database file not found!")
        return False
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # List of expected tables
        expected_tables = [
            "agents",
            "supervisors", 
            "trainers",
            "quality_analysts",
            "operations_managers",
            "attrition"
        ]
        
        # Check if all tables exist
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        existing_tables = [table[0] for table in cursor.fetchall()]
        
        missing_tables = []
        for table in expected_tables:
            if table not in existing_tables:
                missing_tables.append(table)
        
        if missing_tables:
            print(f"ERROR: Missing tables: {missing_tables}")
            return False
        
        print("✓ All expected tables exist")
        
        # Check row counts in each table
        print("\nTable row counts:")
        for table in expected_tables:
            cursor.execute(f"SELECT COUNT(*) FROM {table}")
            count = cursor.fetchone()[0]
            print(f"  {table}: {count} rows")
        
        conn.close()
        return True
        
    except Exception as e:
        print(f"ERROR: {str(e)}")
        return False

def verify_sample_data():
    """Verify that sample data exists in the tables"""
    db_path = os.path.join("data", "roster.db")
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Check if there's data in the agents table
        cursor.execute("SELECT COUNT(*) FROM agents")
        agent_count = cursor.fetchone()[0]
        
        if agent_count > 0:
            print("✓ Sample data found in agents table")
            
            # Show first few agent records
            cursor.execute("SELECT name, agent_id, role FROM agents LIMIT 3")
            agents = cursor.fetchall()
            print("\nSample agents:")
            for agent in agents:
                print(f"  {agent[0]} ({agent[1]}) - {agent[2]}")
        else:
            print("⚠ No data found in agents table")
        
        # Check if there's data in the attrition table
        cursor.execute("SELECT COUNT(*) FROM attrition")
        attrition_count = cursor.fetchone()[0]
        
        if attrition_count > 0:
            print("✓ Sample data found in attrition table")
            
            # Show first few attrition records
            cursor.execute("SELECT AgentName, AgentID, TerminationDate FROM attrition LIMIT 3")
            attritions = cursor.fetchall()
            print("\nSample attrition records:")
            for attrition in attritions:
                print(f"  {attrition[0]} ({attrition[1]}) - {attrition[2]}")
        else:
            print("⚠ No data found in attrition table")
        
        conn.close()
        return True
        
    except Exception as e:
        print(f"ERROR: {str(e)}")
        return False

def main():
    print("Verifying Roster Database Solution...\n")
    
    # Check if data directory exists
    if not os.path.exists("data"):
        print("ERROR: Data directory not found!")
        return
    
    # Verify database structure
    if not verify_database():
        return
    
    print()
    
    # Verify sample data
    verify_sample_data()
    
    print("\n✓ Verification complete!")

if __name__ == "__main__":
    main()
