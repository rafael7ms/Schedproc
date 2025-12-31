import psycopg2
import os
from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT

def init_db():
    """Initialize the PostgreSQL database with required tables"""
    # Database connection parameters
    db_host = os.environ.get('DB_HOST', 'localhost')
    db_name = os.environ.get('DB_NAME', 'roster_db')
    db_user = os.environ.get('DB_USER', 'roster_user')
    db_password = os.environ.get('DB_PASSWORD', 'roster_password')
    
    # Connect to PostgreSQL server (default database)
    conn = psycopg2.connect(
        host=db_host,
        database="postgres",  # Connect to default database first
        user=db_user,
        password=db_password
    )
    conn.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)
    cursor = conn.cursor()
    
    # Create database if it doesn't exist
    try:
        cursor.execute(f"CREATE DATABASE {db_name}")
    except psycopg2.errors.DuplicateDatabase:
        print(f"Database {db_name} already exists")
    finally:
        cursor.close()
        conn.close()
    
    # Connect to the specific database
    conn = psycopg2.connect(
        host=db_host,
        database=db_name,
        user=db_user,
        password=db_password
    )
    cursor = conn.cursor()
    
    # Read the SQL schema from init.sql file
    with open('init.sql', 'r') as f:
        schema = f.read()
    
    # Split schema into individual statements and execute them
    statements = schema.split(';')
    for statement in statements:
        statement = statement.strip()
        if statement:
            try:
                cursor.execute(statement)
            except psycopg2.errors.DuplicateTable:
                print("Table already exists, skipping...")
    
    # Commit changes and close connection
    conn.commit()
    conn.close()
    
    print("Database initialized successfully!")

if __name__ == "__main__":
    init_db()
