import os
import sys
import subprocess

def main():
    """Main script to run the complete solution"""
    print("Roster Database Solution")
    print("=" * 30)
    
    # Check if Excel file is provided as argument
    excel_file = None
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
        if not os.path.exists(excel_file):
            print(f"Error: File '{excel_file}' not found.")
            return 1
    else:
        # Look for Excel files in current directory
        excel_files = [f for f in os.listdir('.') if f.endswith(('.xlsx', '.xls'))]
        if not excel_files:
            print("No Excel files found in current directory.")
            print("Please provide an Excel file as an argument or place one in the current directory.")
            return 1
        
        # Use the first Excel file found
        excel_file = excel_files[0]
        print(f"Using Excel file: {excel_file}")
    
    # Step 1: Initialize the database
    print("\n1. Initializing database...")
    try:
        subprocess.run([sys.executable, 'init_db.py'], check=True)
        print("   Database initialized successfully!")
    except subprocess.CalledProcessError as e:
        print(f"   Error initializing database: {e}")
        return 1
    
    # Step 2: Parse the Excel file and populate the database
    print("\n2. Parsing Excel file and populating database...")
    try:
        cmd = [sys.executable, 'roster_parser.py', excel_file]
        subprocess.run(cmd, check=True)
        print("   Excel file parsed and data inserted successfully!")
    except subprocess.CalledProcessError as e:
        print(f"   Error parsing Excel file: {e}")
        return 1
    
    # Step 3: Show database statistics
    print("\n3. Database statistics:")
    try:
        import psycopg2
        conn = psycopg2.connect(
            host=os.environ.get('DB_HOST', 'localhost'),
            database=os.environ.get('DB_NAME', 'roster_db'),
            user=os.environ.get('DB_USER', 'roster_user'),
            password=os.environ.get('DB_PASSWORD', 'roster_password')
        )
        cursor = conn.cursor()
        
        tables = ['agents', 'supervisors', 'trainers', 'quality_analysts', 'operations_managers', 'attrition']
        for table in tables:
            cursor.execute(f"SELECT COUNT(*) FROM {table}")
            count = cursor.fetchone()[0]
            print(f"   {table.capitalize()}: {count} records")
        
        conn.close()
    except Exception as e:
        print(f"   Error getting database statistics: {e}")
    
    print("\nSolution completed successfully!")
    return 0

if __name__ == "__main__":
    sys.exit(main())
