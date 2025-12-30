-- Create tables for the Roster database

-- Agents table for employees with Role = 'Associate'
CREATE TABLE agents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    last_name TEXT NOT NULL,
    first_name TEXT NOT NULL,
    batch TEXT,
    agent_id TEXT UNIQUE,
    odoo_id TEXT,
    bo_user TEXT,
    axonify TEXT,
    supervisor TEXT,
    manager TEXT,
    tier TEXT,
    shift TEXT,
    schedule TEXT,
    department TEXT,
    role TEXT,
    phase_1_date DATE,
    phase_2_date DATE,
    phase_3_date DATE,
    hire_date DATE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Supervisors table for employees with Role = 'Supervisor'
CREATE TABLE supervisors (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    last_name TEXT NOT NULL,
    first_name TEXT NOT NULL,
    batch TEXT,
    agent_id TEXT UNIQUE,
    odoo_id TEXT,
    bo_user TEXT,
    axonify TEXT,
    supervisor TEXT,
    manager TEXT,
    tier TEXT,
    shift TEXT,
    schedule TEXT,
    department TEXT,
    role TEXT,
    phase_1_date DATE,
    phase_2_date DATE,
    phase_3_date DATE,
    hire_date DATE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Trainers table for employees with Role = 'Trainer'
CREATE TABLE trainers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    last_name TEXT NOT NULL,
    first_name TEXT NOT NULL,
    batch TEXT,
    agent_id TEXT UNIQUE,
    odoo_id TEXT,
    bo_user TEXT,
    axonify TEXT,
    supervisor TEXT,
    manager TEXT,
    tier TEXT,
    shift TEXT,
    schedule TEXT,
    department TEXT,
    role TEXT,
    phase_1_date DATE,
    phase_2_date DATE,
    phase_3_date DATE,
    hire_date DATE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Quality Analysts table for employees with Role = 'Analyst'
CREATE TABLE quality_analysts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    last_name TEXT NOT NULL,
    first_name TEXT NOT NULL,
    batch TEXT,
    agent_id TEXT UNIQUE,
    odoo_id TEXT,
    bo_user TEXT,
    axonify TEXT,
    supervisor TEXT,
    manager TEXT,
    tier TEXT,
    shift TEXT,
    schedule TEXT,
    department TEXT,
    role TEXT,
    phase_1_date DATE,
    phase_2_date DATE,
    phase_3_date DATE,
    hire_date DATE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Operations Managers table for employees with Role = 'OM'
CREATE TABLE operations_managers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    last_name TEXT NOT NULL,
    first_name TEXT NOT NULL,
    batch TEXT,
    agent_id TEXT UNIQUE,
    odoo_id TEXT,
    bo_user TEXT,
    axonify TEXT,
    supervisor TEXT,
    manager TEXT,
    tier TEXT,
    shift TEXT,
    schedule TEXT,
    department TEXT,
    role TEXT,
    phase_1_date DATE,
    phase_2_date DATE,
    phase_3_date DATE,
    hire_date DATE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Attrition table
CREATE TABLE attrition (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    last_name TEXT NOT NULL,
    first_name TEXT NOT NULL,
    batch TEXT,
    agent_id TEXT,
    odoo_id TEXT,
    bo_user TEXT,
    axonify TEXT,
    supervisor TEXT,
    manager TEXT,
    tier TEXT,
    shift TEXT,
    schedule TEXT,
    department TEXT,
    role TEXT,
    type_of_attrition TEXT,
    term_date DATE,
    gds_ticket TEXT,
    wfm_ticket TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Create indexes for better performance
CREATE INDEX idx_agents_agent_id ON agents(agent_id);
CREATE INDEX idx_supervisors_agent_id ON supervisors(agent_id);
CREATE INDEX idx_trainers_agent_id ON trainers(agent_id);
CREATE INDEX idx_quality_analysts_agent_id ON quality_analysts(agent_id);
CREATE INDEX idx_operations_managers_agent_id ON operations_managers(agent_id);
CREATE INDEX idx_attrition_agent_id ON attrition(agent_id);
