# Database Schema Documentation

## Overview

This document describes the database schema for the Roster Database Solution. The database is implemented using SQLite and consists of six tables that store information about employees and attrition data.

## Tables

### 1. agents

Stores information about employees with the role of "Associate".

#### Columns:
| Column Name | Data Type | Constraints | Description |
|-------------|-----------|-------------|-------------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | Unique identifier for each record |
| Name | TEXT | NOT NULL | Full name of the employee |
| LastName | TEXT | NOT NULL | Last name of the employee |
| FirstName | TEXT | NOT NULL | First name of the employee |
| Batch | TEXT | | Batch information |
| AgentID | TEXT | UNIQUE | Unique agent identifier |
| ODOO_ID | TEXT | UNIQUE | Unique ODOO identifier |
| BOUser | TEXT | | BO User information |
| Axonify | TEXT | | Axonify information |
| Supervisor | TEXT | | Name of the supervisor |
| Manager | TEXT | | Name of the manager |
| Tier | TEXT | | Tier level |
| Shift | TEXT | | Shift information |
| Schedule | TEXT | | Schedule information |
| Department | TEXT | | Department name |
| Role | TEXT | | Role (always "Associate" for this table) |
| Phase1Date | DATE | | Phase 1 date |
| Phase2Date | DATE | | Phase 2 date |
| Phase3Date | DATE | | Phase 3 date |
| HireDate | DATE | | Hire date |

### 2. supervisors

Stores information about employees with the role of "Supervisor".

#### Columns:
| Column Name | Data Type | Constraints | Description |
|-------------|-----------|-------------|-------------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | Unique identifier for each record |
| Name | TEXT | NOT NULL | Full name of the employee |
| LastName | TEXT | NOT NULL | Last name of the employee |
| FirstName | TEXT | NOT NULL | First name of the employee |
| Batch | TEXT | | Batch information |
| AgentID | TEXT | UNIQUE | Unique agent identifier |
| ODOO_ID | TEXT | UNIQUE | Unique ODOO identifier |
| BOUser | TEXT | | BO User information |
| Axonify | TEXT | | Axonify information |
| Supervisor | TEXT | | Name of the supervisor |
| Manager | TEXT | | Name of the manager |
| Tier | TEXT | | Tier level |
| Shift | TEXT | | Shift information |
| Schedule | TEXT | | Schedule information |
| Department | TEXT | | Department name |
| Role | TEXT | | Role (always "Supervisor" for this table) |
| Phase1Date | DATE | | Phase 1 date |
| Phase2Date | DATE | | Phase 2 date |
| Phase3Date | DATE | | Phase 3 date |
| HireDate | DATE | | Hire date |

### 3. trainers

Stores information about employees with the role of "Trainer".

#### Columns:
| Column Name | Data Type | Constraints | Description |
|-------------|-----------|-------------|-------------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | Unique identifier for each record |
| Name | TEXT | NOT NULL | Full name of the employee |
| LastName | TEXT | NOT NULL | Last name of the employee |
| FirstName | TEXT | NOT NULL | First name of the employee |
| Batch | TEXT | | Batch information |
| AgentID | TEXT | UNIQUE | Unique agent identifier |
| ODOO_ID | TEXT | UNIQUE | Unique ODOO identifier |
| BOUser | TEXT | | BO User information |
| Axonify | TEXT | | Axonify information |
| Supervisor | TEXT | | Name of the supervisor |
| Manager | TEXT | | Name of the manager |
| Tier | TEXT | | Tier level |
| Shift | TEXT | | Shift information |
| Schedule | TEXT | | Schedule information |
| Department | TEXT | | Department name |
| Role | TEXT | | Role (always "Trainer" for this table) |
| Phase1Date | DATE | | Phase 1 date |
| Phase2Date | DATE | | Phase 2 date |
| Phase3Date | DATE | | Phase 3 date |
| HireDate | DATE | | Hire date |

### 4. quality_analysts

Stores information about employees with the role of "QA" (Quality Analyst).

#### Columns:
| Column Name | Data Type | Constraints | Description |
|-------------|-----------|-------------|-------------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | Unique identifier for each record |
| Name | TEXT | NOT NULL | Full name of the employee |
| LastName | TEXT | NOT NULL | Last name of the employee |
| FirstName | TEXT | NOT NULL | First name of the employee |
| Batch | TEXT | | Batch information |
| AgentID | TEXT | UNIQUE | Unique agent identifier |
| ODOO_ID | TEXT | UNIQUE | Unique ODOO identifier |
| BOUser | TEXT | | BO User information |
| Axonify | TEXT | | Axonify information |
| Supervisor | TEXT | | Name of the supervisor |
| Manager | TEXT | | Name of the manager |
| Tier | TEXT | | Tier level |
| Shift | TEXT | | Shift information |
| Schedule | TEXT | | Schedule information |
| Department | TEXT | | Department name |
| Role | TEXT | | Role (always "QA" for this table) |
| Phase1Date | DATE | | Phase 1 date |
| Phase2Date | DATE | | Phase 2 date |
| Phase3Date | DATE | | Phase 3 date |
| HireDate | DATE | | Hire date |

### 5. operations_managers

Stores information about employees with the role of "OM" (Operations Manager).

#### Columns:
| Column Name | Data Type | Constraints | Description |
|-------------|-----------|-------------|-------------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | Unique identifier for each record |
| Name | TEXT | NOT NULL | Full name of the employee |
| LastName | TEXT | NOT NULL | Last name of the employee |
| FirstName | TEXT | NOT NULL | First name of the employee |
| Batch | TEXT | | Batch information |
| AgentID | TEXT | UNIQUE | Unique agent identifier |
| ODOO_ID | TEXT | UNIQUE | Unique ODOO identifier |
| BOUser | TEXT | | BO User information |
| Axonify | TEXT | | Axonify information |
| Supervisor | TEXT | | Name of the supervisor |
| Manager | TEXT | | Name of the manager |
| Tier | TEXT | | Tier level |
| Shift | TEXT | | Shift information |
| Schedule | TEXT | | Schedule information |
| Department | TEXT | | Department name |
| Role | TEXT | | Role (always "OM" for this table) |
| Phase1Date | DATE | | Phase 1 date |
| Phase2Date | DATE | | Phase 2 date |
| Phase3Date | DATE | | Phase 3 date |
| HireDate | DATE | | Hire date |

### 6. attrition

Stores attrition data from the Attrition sheet.

#### Columns:
| Column Name | Data Type | Constraints | Description |
|-------------|-----------|-------------|-------------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | Unique identifier for each record |
| AgentName | TEXT | NOT NULL | Full name of the agent |
| LastName | TEXT | NOT NULL | Last name of the agent |
| FirstName | TEXT | NOT NULL | First name of the agent |
| AgentID | TEXT | UNIQUE | Unique agent identifier |
| ODOO_ID | TEXT | UNIQUE | Unique ODOO identifier |
| HireDate | DATE | | Hire date |
| TerminationDate | DATE | | Termination date |
| TermReason | TEXT | | Reason for termination |
| Voluntary | TEXT | | Whether the termination was voluntary |
| Department | TEXT | | Department name |
| Role | TEXT | | Role of the employee |
| Supervisor | TEXT | | Name of the supervisor |
| Manager | TEXT | | Name of the manager |
| Shift | TEXT | | Shift information |
| Schedule | TEXT | | Schedule information |
| Tier | TEXT | | Tier level |
| Batch | TEXT | | Batch information |
| TenureDays | INTEGER | | Tenure in days |
| TenureMonths | REAL | | Tenure in months |
| TenureYears | REAL | | Tenure in years |

## Unique Constraints

All tables have unique constraints on the following columns to prevent duplicate entries:
- `AgentID`
- `ODOO_ID`

These constraints ensure that each employee is only represented once in the database, even if the Excel file contains duplicate records.

## Relationships

There are no explicit foreign key relationships between the tables, as each table represents a distinct group of employees. However, the data in all tables follows the same structure for employee information, making it possible to perform cross-table queries when needed.

## Data Types

- **TEXT**: Used for string data
- **DATE**: Used for date values (stored as TEXT in SQLite but formatted as YYYY-MM-DD)
- **INTEGER**: Used for numeric values without decimal points
- **REAL**: Used for numeric values with decimal points

## Indexes

The database automatically creates indexes for the primary key columns (id) in each table. Additional indexes are created for the AgentID and ODOO_ID columns due to their UNIQUE constraints, which improves query performance when searching for specific employees.
