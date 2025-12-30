# Roster Database Solution

This solution creates a database from the "Roster - December 2025 updated TC.xlsx" file with tables for different roles and attrition data.

## Solution Components

1. **Database Schema** - SQLite database with tables for:
   - Agents (Associates)
   - Supervisors
   - Trainers
   - Quality Analysts (QA)
   - Operations Managers (OM)
   - Attrition data

2. **Containerization** - Docker configuration for persistent storage

3. **Python Scripts** - For parsing Excel files and populating the database:
   - `init_db.py` - Initializes the database with the schema
   - `roster_parser.py` - Parses Excel files and inserts data into the database
   - `run_solution.py` - Main script to run the complete solution
   - `test_solution.py` - Tests the solution components

## Database Design

The database consists of 6 tables:

1. **agents** - For all employees with Role 'Associate'
2. **supervisors** - For all employees with Role 'Supervisor'
3. **trainers** - For all employees with Role 'Trainer'
4. **quality_analysts** - For all employees with Role 'QA' (Quality Analyst)
5. **operations_managers** - For all employees with Role 'OM' (Operations Manager)
6. **attrition** - For attrition data from the Attrition sheet

All tables have unique constraints on Agent ID and ODOO ID to prevent duplicate entries.

## Requirements

- Python 3.6+
- Docker (for containerization)
- Required Python packages (installed via `pip install -r requirements.txt`):
  - pandas
  - openpyxl
  - sqlite3 (usually included with Python)

## Setup and Usage

1. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Run the solution**:
   ```bash
   python run_solution.py [path/to/roster.xlsx]
   ```

   If no Excel file is provided as an argument, the script will use the first `.xlsx` or `.xls` file found in the current directory.

3. **Run with Docker** (optional):
   ```bash
   docker-compose up
   ```

## How It Works

1. **Database Initialization**:
   - Creates a `data` directory for persistent storage
   - Initializes SQLite database with the schema defined in `init.sql`

2. **Excel Parsing**:
   - Reads the "7MS main roster" sheet to populate role-based tables
   - Reads the "Attrition" sheet to populate the attrition table
   - Handles duplicate detection using Agent ID and ODOO ID
   - Prevents duplicate entries in the database

3. **Data Processing**:
   - Separates employees by role into different tables
   - Cleans and formats data before insertion
   - Provides statistics on the number of records in each table

## Testing

Run the test script to verify the solution works correctly:

```bash
python test_solution.py
```

This will:
- Test database creation
- Verify all tables are created correctly
- Test duplicate prevention mechanisms

## Running the Demo

To process the Roster - December 2025 updated TC.xlsx file:

1. Ensure the Excel file is in the same directory as the script
2. Run the demo script: `python demo.py`

The script will automatically process the Excel file and populate the database with the roster information, separating employees by role into different tables and handling any duplicate records.

## File Structure

```
├── data/                   # Persistent storage for database
├── init.sql                # Database schema
├── init_db.py              # Database initialization script
├── roster_parser.py        # Excel parsing and data insertion
├── run_solution.py         # Main execution script
├── test_solution.py        # Test script
├── requirements.txt        # Python dependencies
├── Dockerfile              # Docker configuration
├── docker-compose.yml      # Docker Compose configuration
└── README.md               # This file
```

## Database Schema Details

### Agents, Supervisors, Trainers, Quality Analysts, Operations Managers Tables
All role-based tables have the same structure:
- Name (TEXT)
- LastName (TEXT)
- FirstName (TEXT)
- Batch (TEXT)
- AgentID (TEXT) - UNIQUE
- ODOO_ID (TEXT) - UNIQUE
- BOUser (TEXT)
- Axonify (TEXT)
- Supervisor (TEXT)
- Manager (TEXT)
- Tier (TEXT)
- Shift (TEXT)
- Schedule (TEXT)
- Department (TEXT)
- Role (TEXT)
- Phase1Date (DATE)
- Phase2Date (DATE)
- Phase3Date (DATE)
- HireDate (DATE)

### Attrition Table
- AgentName (TEXT)
- LastName (TEXT)
- FirstName (TEXT)
- AgentID (TEXT) - UNIQUE
- ODOO_ID (TEXT) - UNIQUE
- HireDate (DATE)
- TerminationDate (DATE)
- TermReason (TEXT)
- Voluntary (TEXT)
- Department (TEXT)
- Role (TEXT)
- Supervisor (TEXT)
- Manager (TEXT)
- Shift (TEXT)
- Schedule (TEXT)
- Tier (TEXT)
- Batch (TEXT)
- TenureDays (INTEGER)
- TenureMonths (REAL)
- TenureYears (REAL)

## Notes

- The solution handles duplicate records by using UNIQUE constraints on Agent ID and ODOO ID
- All data is stored in the `data/roster.db` SQLite database
- The Docker configuration ensures persistent storage of the database file
- The solution is designed to work with the specific structure of the "Roster - December 2025 updated TC.xlsx" file
