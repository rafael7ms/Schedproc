import psycopg2
import os

# Connect to the database
conn = psycopg2.connect(
    host=os.environ.get('DB_HOST', 'localhost'),
    database=os.environ.get('DB_NAME', 'roster_db'),
    user=os.environ.get('DB_USER', 'roster_user'),
    password=os.environ.get('DB_PASSWORD', 'roster_password')
)
cursor = conn.cursor()

# Get attrition table schema
cursor.execute("SELECT column_name, data_type FROM information_schema.columns WHERE table_name = 'attrition' ORDER BY ordinal_position;")
print('Attrition table schema:')
print(cursor.fetchall())

# Get a few rows from the attrition table
cursor.execute("SELECT * FROM attrition LIMIT 5;")
print('\nSample data from attrition table:')
print(cursor.fetchall())

conn.close()
