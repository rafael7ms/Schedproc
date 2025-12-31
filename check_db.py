import psycopg2
import os

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

# Connect to the database
conn = connect_to_db()
if not conn:
    print("Failed to connect to database")
    exit(1)

cursor = conn.cursor()

# Get all tables
cursor.execute("SELECT tablename FROM pg_tables WHERE schemaname = 'public'")
tables = cursor.fetchall()

print("Tables in the database:")
for table in tables:
    print(f"  - {table[0]}")
    
    # Get row count for each table
    cursor.execute(f"SELECT COUNT(*) FROM {table[0]}")
    count = cursor.fetchone()[0]
    print(f"    Rows: {count}")

conn.close()
