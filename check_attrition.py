import psycopg2
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Connect to the database
conn = psycopg2.connect(
    host=os.getenv('DB_HOST', 'localhost'),
    database=os.getenv('DB_NAME', 'roster_db'),
    user=os.getenv('DB_USER', 'roster_user'),
    password=os.getenv('DB_PASSWORD', 'roster_password'),
    port=os.getenv('DB_PORT', '5432')
)
cursor = conn.cursor()

# Get attrition table structure
cursor.execute("""
    SELECT column_name, data_type 
    FROM information_schema.columns 
    WHERE table_name = 'attrition'
""")
columns = cursor.fetchall()

print("Attrition table structure:")
for column in columns:
    print(f"  {column[0]} ({column[1]})")

# Get sample data
cursor.execute("SELECT * FROM attrition LIMIT 3")
rows = cursor.fetchall()

print("\nSample attrition data:")
for row in rows:
    print(f"  {row}")

conn.close()
