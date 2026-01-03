import sqlite3

# Connect to the database
conn = sqlite3.connect('roster.db')
c = conn.cursor()

# Check for records with empty/NaN Agent IDs
c.execute('SELECT COUNT(*) FROM attrition WHERE agent_id IS NULL OR agent_id = "" OR agent_id = "NaN"')
result = c.fetchone()
print(f'Records with empty/NaN Agent IDs: {result[0]}')

# Let's also check what those records are
c.execute('SELECT name, agent_id FROM attrition WHERE agent_id IS NULL OR agent_id = "" OR agent_id = "NaN"')
records = c.fetchall()
if records:
    print("\nRecords with empty/NaN Agent IDs:")
    for record in records:
        print(f"  {record[0]}: '{record[1]}'")
else:
    print("\nNo records with empty/NaN Agent IDs found.")

conn.close()
