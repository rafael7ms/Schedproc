import sqlite3

# Connect to the database
conn = sqlite3.connect('data/roster.db')
cursor = conn.cursor()

# Get attrition table structure
cursor.execute("PRAGMA table_info(attrition)")
columns = cursor.fetchall()

print("Attrition table structure:")
for column in columns:
    print(f"  {column[1]} ({column[2]})")

# Get sample data
cursor.execute("SELECT * FROM attrition LIMIT 3")
rows = cursor.fetchall()

print("\nSample attrition data:")
for row in rows:
    print(f"  {row}")

conn.close()
