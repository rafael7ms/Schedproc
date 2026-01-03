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

def main():
    conn = connect_to_db()
    if not conn:
        print("Failed to connect to database")
        return
    
    try:
        cur = conn.cursor()
        
        # Check agents
        cur.execute('SELECT name, agent_id, role FROM agents')
        print('Agents:')
        for row in cur.fetchall():
            print(f"  {row}")
        
        # Check supervisors
        cur.execute('SELECT name, agent_id, role FROM supervisors')
        print('Supervisors:')
        for row in cur.fetchall():
            print(f"  {row}")
        
        # Check trainers
        cur.execute('SELECT name, agent_id, role FROM trainers')
        print('Trainers:')
        for row in cur.fetchall():
            print(f"  {row}")
        
        # Check quality analysts
        cur.execute('SELECT name, agent_id, role FROM quality_analysts')
        print('Quality Analysts:')
        for row in cur.fetchall():
            print(f"  {row}")
        
        # Check operations managers
        cur.execute('SELECT name, agent_id, role FROM operations_managers')
        print('Operations Managers:')
        for row in cur.fetchall():
            print(f"  {row}")
        
        # Check attrition
        cur.execute('SELECT name, agent_id, role FROM attrition')
        print('Attrition:')
        for row in cur.fetchall():
            print(f"  {row}")
        
        conn.close()
    except Exception as e:
        print(f"Error checking database content: {e}")
        if conn:
            conn.close()

if __name__ == "__main__":
    main()
