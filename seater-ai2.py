import pandas as pd
import json
import yaml
import ollama
from collections import defaultdict
from typing import Dict, List, Tuple, Any, Optional
import re
import argparse
import sys
import os
import numpy as np
from datetime import datetime, time

# === Constants ===
TOTAL_SEATS = 69
DEFAULT_RESERVED_SEATS = [61]
OLLAMA_SERVER = "172.16.30.202"
OLLAMA_PORT = 11434
LOG_LEVELS = ["DEBUG", "INFO", "WARNING", "ERROR"]

# === Seat Structure (Area/Subarea Mapping) ===
SEAT_AREA_MAP = {
    i: ("OPS1-A", "A") if 1 <= i <= 12 else
       ("OPS1-B", "B") if 13 <= i <= 18 else
       ("OPS1-C", "C") if 19 <= i <= 24 else
       ("OPS1-D", "D") if 25 <= i <= 39 else
       ("OPS1-E", "E") if 40 <= i <= 47 else
       ("OPS2", "F") if 48 <= i <= 61 else
       ("OPS3", "G") if 62 <= i <= 69 else
       ("TRN", "H") if 71 <= i <= 100 else
       ("UNKNOWN", "UNKNOWN")
    for i in range(1, TOTAL_SEATS + 1)
}

# === Logging System ===
class SeatAssignmentLogger:
    def __init__(self, log_file: str = None, level: str = "INFO"):
        self.log_file = log_file or f"seat_assignment_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        self.level = level
        self.log(f"Seat Assignment System initialized - {datetime.now().isoformat()}", "INFO")
        self.log(f"Using Ollama server at {OLLAMA_SERVER}:{OLLAMA_PORT}", "INFO")

    def log(self, message: str, level: str = "INFO"):
        if LOG_LEVELS.index(level) >= LOG_LEVELS.index(self.level):
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_entry = f"[{timestamp}] [{level}] {message}"
            print(log_entry)
            if self.log_file:
                with open(self.log_file, 'a', encoding='utf-8') as f:
                    f.write(log_entry + '\n')

    def debug(self, message: str):
        self.log(message, "DEBUG")

    def info(self, message: str):
        self.log(message, "INFO")

    def warning(self, message: str):
        self.log(message, "WARNING")

    def error(self, message: str):
        self.log(message, "ERROR")

# === Configure Ollama Client ===
def configure_ollama_client(server_ip: str, port: int):
    try:
        ollama.Client(host=f"http://{server_ip}:{port}")
        response = ollama.list()
        print(f"Successfully connected to Ollama server at {server_ip}:{port}")
        print(f"Available models: {[model['name'] for model in response['models']]}")
        return True
    except Exception as e:
        print(f"Failed to connect to Ollama server at {server_ip}:{port}: {e}")
        return False

# === Shift Overlap Check ===
def parse_shift_time(shift_str: str) -> Tuple[time, time]:
    try:
        start_str, end_str = shift_str.split('-')
        start_time = datetime.strptime(start_str.strip(), '%H:%M').time()
        end_time = datetime.strptime(end_str.strip(), '%H:%M').time()
        return start_time, end_time
    except Exception:
        return time(0, 0), time(23, 59)

def shifts_overlap(shift1: str, shift2: str) -> bool:
    if shift1 == shift2:
        return True

    try:
        s1_start, s1_end = parse_shift_time(shift1)
        s2_start, s2_end = parse_shift_time(shift2)
    except:
        return False

    # Handle night shift which crosses midnight [1]
    if s1_start > s1_end or s2_start > s2_end:
        return True

    # Check if end of shift1 is after start of shift2 [1]
    if s1_end > s2_start and s2_end > s1_start:
        return True

    return False

# === Data Analysis ===
def analyze_agent_data(agent_df: pd.DataFrame, logger: SeatAssignmentLogger) -> Dict:
    logger.info("Analyzing agent data...")

    # Extract date range
    date_range = pd.to_datetime(agent_df['Date']).sort_values()
    date_info = {
        "start_date": str(date_range.min().date()),
        "end_date": str(date_range.max().date()),
        "total_days": len(date_range.unique())
    }
    logger.info(f"Date range: {date_info['start_date']} to {date_info['end_date']} ({date_info['total_days']} days)")

    # Extract unique queues
    queues = agent_df['Queue'].unique().tolist()
    logger.info(f"Found {len(queues)} unique queues: {queues}")

    # Extract unique shifts
    shifts = agent_df['Shift'].unique().tolist()
    logger.info(f"Found {len(shifts)} unique shifts: {shifts}")

    # Extract agent information
    agent_info = []
    for _, row in agent_df.iterrows():
        agent_info.append({
            "id": row['ID'],
            "name": row['Name'],
            "queue": row['Queue'],
            "shift": row['Shift'],
            "supervisor": row['Supervisor'],
            "batch": int(row.get('Batch', 0))
        })

    logger.info(f"Extracted information for {len(agent_info)} agents")

    return {
        "date_range": date_info,
        "queues": queues,
        "shifts": shifts,
        "agents": agent_info
    }

# === Rule Parser with Ollama ===
def parse_rules_with_ollama(rules_input: Any, data_analysis: Dict, model_name: str, logger: SeatAssignmentLogger) -> Dict:
    if isinstance(rules_input, str) and os.path.exists(rules_input) and rules_input.endswith('.txt'):
        try:
            with open(rules_input, 'r', encoding='utf-8') as f:
                rules_text = f.read().strip()
            logger.info(f"Loaded rules from text file: {rules_input}")
        except Exception as e:
            logger.error(f"Error reading text file: {e}")
            rules_text = rules_input
    elif isinstance(rules_input, str):
        rules_text = rules_input
        logger.debug(f"Using direct text input for rules: {rules_text[:100]}...")
    else:
        logger.error("Unsupported rule input format.")
        raise ValueError("Unsupported rule input format.")

    prompt = f"""
    You are an expert in office seating arrangement. Your task is to convert the following natural language rules into a structured JSON format.

    Here is the information about the agents and scheduling period:
    Period: {data_analysis['date_range']['start_date']} to {data_analysis['date_range']['end_date']} ({data_analysis['date_range']['total_days']} days)
    Available Queues: {', '.join(data_analysis['queues'])}
    Available Shifts: {', '.join(data_analysis['shifts'])}
    Total Agents: {len(data_analysis['agents'])}

    Provide your reasoning step by step:
    1. Identify all grouping priorities mentioned [1]
    2. List any reserved seats [1]
    3. Extract area constraints for specific shifts, queues, or agents [1]
    4. Note any proximity requirements [1]
    5. Identify any name-based constraints [1]
    6. Identify any batch/seniority constraints [1]
    7. Identify any shift-specific constraints [1]
    8. Identify any capacity constraints [1]
    9. Note any preferred or avoided seats [1]

    Rules: "{rules_text}"

    Return ONLY valid JSON in this format:
    {{
      "priority": ["supervisor", "queue", "batch"],
      "reserved_seats": [61],
      "area_constraints": {{"Night": ["OPS1-D"]}},
      "proximity_rules": {{"max_distance": 3}},
      "name_constraints": {{"John Doe": {{"preferred_areas": ["OPS1-A"]}}}},
      "batch_constraints": {{"senior_first": true}},
      "shift_constraints": {{"non_overlapping_only": true}},
      "capacity_constraints": {{"OPS1-A": 12}},
      "preferred_seats": {{"A123": [1, 2, 3]}},
      "avoid_seats": {{"B456": [10, 11, 12]}}
    }}
    """

    try:
        logger.info(f"Using Ollama model '{model_name}' to parse rules...")
        response = ollama.generate(
            model=model_name,
            prompt=prompt,
            options={"temperature": 0.3}
        )

        full_response = response['response'].strip()

        if "```json" in full_response:
            reasoning, json_str = full_response.split("```json", 1)
            json_str = json_str.split("```")[0].strip()
        elif "```" in full_response:
            reasoning, json_str = full_response.split("```", 1)
            json_str = json_str.split("```")[0].strip()
        else:
            reasoning = "No explicit reasoning provided by model."
            json_str = full_response

        logger.info("Ollama reasoning process:")
        for line in reasoning.split('\n'):
            if line.strip():
                logger.info(f"  - {line}")

        rules = json.loads(json_str)
        logger.info("Successfully parsed rules from Ollama response")
        return rules

    except Exception as e:
        logger.error(f"Ollama parsing failed: {e}. Using fallback parser.")
        return parse_natural_language_fallback(rules_text, data_analysis)

def parse_natural_language_fallback(prompt: str, data_analysis: Dict) -> Dict:
    rules = {
        "reserved_seats": [61],
        "priority": ["supervisor", "queue", "batch"],
        "area_constraints": {},
        "proximity_rules": {},
        "name_constraints": {},
        "batch_constraints": {"senior_first": True},
        "shift_constraints": {"non_overlapping_only": False},
        "capacity_constraints": {},
        "preferred_seats": {},
        "avoid_seats": {}
    }

    # Reserved seats
    reserved_match = re.search(r"reserve[d]?.*seats?\s*([0-9,\s]+)", prompt, re.IGNORECASE)
    if reserved_match:
        rules["reserved_seats"] = list(map(int, re.findall(r'\d+', reserved_match.group(1))))

    # Priority
    priority_match = re.search(r"group.*by\s+(.+?)(?:\.|,|$)", prompt, re.IGNORECASE)
    if priority_match:
        attrs = priority_match.group(1).split(",")
        rules["priority"] = [attr.strip().lower() for attr in attrs]

    # Area constraints
    area_matches = re.finditer(r"(night|morning|afternoon|mixed|bns|ibc|customer|[\w\s]+)\s+agents?.*in\s+([a-z0-9\- ,]+)", prompt, re.IGNORECASE)
    for match in area_matches:
        shift_or_queue_or_name = match.group(1).strip()
        areas = [a.strip() for a in match.group(2).split(",")]
        rules["area_constraints"][shift_or_queue_or_name.title()] = areas

    # Proximity rules
    proximity_match = re.search(r"no (?:more than|over) (\d+).*seats?.*apart", prompt, re.IGNORECASE)
    if proximity_match:
        rules["proximity_rules"]["max_distance"] = int(proximity_match.group(1))

    return rules

# === Seat Assignment Engine ===
def ai_assign_seats(agent_df: pd.DataFrame, rules: Dict, data_analysis: Dict, logger: SeatAssignmentLogger) -> Dict:
    logger.info("Starting seat assignment process...")

    # Initialize with default values
    reserved_seats = set(rules.get("reserved_seats", DEFAULT_RESERVED_SEATS))
    priority_order = rules.get("priority", ["supervisor", "queue", "batch"])
    area_constraints = rules.get("area_constraints", {})
    proximity_rules = rules.get("proximity_rules", {})
    name_constraints = rules.get("name_constraints", {})
    batch_constraints = rules.get("batch_constraints", {"senior_first": True})
    shift_constraints = rules.get("shift_constraints", {"non_overlapping_only": False})
    capacity_constraints = rules.get("capacity_constraints", {})
    preferred_seats = rules.get("preferred_seats", {})
    avoid_seats = rules.get("avoid_seats", {})

    # Get all unique dates and convert to datetime.date objects
    dates = pd.to_datetime(agent_df['Date']).unique()
    dates = np.sort(dates)
    date_objects = [pd.to_datetime(date).date() for date in dates]

    # Initialize seat assignments dictionary
    seat_assignments = {date: {} for date in date_objects}

    # Precompute seat availability and area info
    available_seats = [i for i in range(1, TOTAL_SEATS + 1) if i not in reserved_seats]
    seat_to_area = {i: SEAT_AREA_MAP[i] for i in available_seats}

    # Apply capacity constraints
    area_capacities = {}
    for area, _ in set(SEAT_AREA_MAP.values()):
        if area in capacity_constraints:
            area_capacities[area] = capacity_constraints[area]
        else:
            area_capacities[area] = len([s for s in available_seats if SEAT_AREA_MAP[s][0] == area])

    logger.info(f"Area capacities: {area_capacities}")

    # Process each date
    for i, date in enumerate(date_objects):
        logger.info(f"Processing date: {date}")

        # Get agents for this date
        date_str = date.strftime('%Y-%m-%d')
        date_agents = agent_df[agent_df['Date'] == date_str]

        # Sort agents based on priority and batch
        sort_columns = [col for col in priority_order if col in date_agents.columns]
        if "batch" in date_agents.columns and "batch" in priority_order:
            if batch_constraints.get("senior_first", True):
                date_agents = date_agents.sort_values(by="batch", ascending=True)
            else:
                date_agents = date_agents.sort_values(by="batch", ascending=False)

        if sort_columns:
            date_agents = date_agents.sort_values(by=sort_columns)

        # Group agents by priority attributes
        def group_key(row):
            return tuple(str(row[attr]) for attr in priority_order if attr in row)

        grouped_agents = defaultdict(list)
        for _, row in date_agents.iterrows():
            key = group_key(row)
            grouped_agents[key].append(row)

        logger.info(f"  Grouped {len(date_agents)} agents into {len(grouped_agents)} groups for {date}")

        # Initialize available seats for this date
        date_available_seats = available_seats.copy()
        date_seat_map = {}

        # Process each group
        for key, agents in grouped_agents.items():
            group_name = " | ".join(str(k) for k in key)
            logger.info(f"  Processing group: {group_name} ({len(agents)} agents)")

            # Determine area constraints for this group
            allowed_areas = set()
            for constraint_key, areas in area_constraints.items():
                for agent in agents:
                    if (constraint_key.lower() in str(agent.get("Shift", "")).lower() or
                        constraint_key.lower() in str(agent.get("Queue", "")).lower() or
                        constraint_key.lower() in str(agent.get("Name", "")).lower()):
                        allowed_areas.update(areas)
                        logger.debug(f"    Applied constraint: '{constraint_key}' → {areas}")

            if not allowed_areas:
                logger.debug("    No specific area constraints for this group")
                allowed_areas = set([area for area, _ in set(SEAT_AREA_MAP.values())])
            else:
                logger.info(f"    Allowed areas for group: {allowed_areas}")

            # Filter available seats by area and capacity
            eligible_seats = []
            for s in date_available_seats:
                area, _ = seat_to_area[s]
                if area in allowed_areas:
                    if area in area_capacities:
                        current_count = len([seat for seat, a in date_seat_map.items() if a[0] == area])
                        if current_count < area_capacities[area]:
                            eligible_seats.append(s)
                    else:
                        eligible_seats.append(s)

            logger.debug(f"    Eligible seats: {len(eligible_seats)}")

            # Apply name constraints
            for agent in agents:
                agent_id = agent['ID']
                if agent_id in preferred_seats:
                    for seat in preferred_seats[agent_id]:
                        if seat in eligible_seats and seat not in date_seat_map:
                            date_seat_map[seat] = (seat_to_area[seat][0], agent['Name'], agent['Shift'])
                            eligible_seats.remove(seat)
                            logger.debug(f"    Assigned preferred seat {seat} to {agent['Name']}")
                            break

                if agent_id in avoid_seats:
                    for seat in avoid_seats[agent_id]:
                        if seat in eligible_seats:
                            eligible_seats.remove(seat)
                            logger.debug(f"    Removed avoided seat {seat} for {agent['Name']}")

            # Apply name constraints from name_constraints
            for agent in agents:
                agent_name = agent['Name']
                if agent_name in name_constraints:
                    constraints = name_constraints[agent_name]
                    if "preferred_areas" in constraints:
                        preferred_areas = set(constraints["preferred_areas"])
                        eligible_seats = [s for s in eligible_seats if seat_to_area[s][0] in preferred_areas]
                        logger.debug(f"    Filtered to preferred areas {preferred_areas} for {agent_name}")

                    if "avoid_seats" in constraints:
                        avoid = set(constraints["avoid_seats"])
                        eligible_seats = [s for s in eligible_seats if s not in avoid]
                        logger.debug(f"    Removed avoided seats {avoid} for {agent_name}")

            # Try to assign seats contiguously [1]
            assigned_seats = []
            for i in range(len(eligible_seats) - len(agents) + 1):
                block = eligible_seats[i:i + len(agents)]
                if len(block) == len(agents):
                    # Check proximity rules [1]
                    if "max_distance" in proximity_rules:
                        max_dist = proximity_rules["max_distance"]
                        valid = True
                        for j in range(len(block) - 1):
                            if abs(block[j] - block[j + 1]) > max_dist:
                                valid = False
                                break
                        if valid:
                            assigned_seats = block
                            break
                    else:
                        assigned_seats = block
                        break

            if not assigned_seats:
                assigned_seats = eligible_seats[:len(agents)]
                logger.warning(f"    Could not find contiguous seats. Assigning non-contiguous seats: {assigned_seats}")

            if len(assigned_seats) < len(agents):
                logger.error(f"    Insufficient seats: needed {len(agents)}, assigned {len(assigned_seats)}")
                assigned_seats = assigned_seats[:len(agents)]

            # Assign seats to agents
            for agent, seat in zip(agents, assigned_seats):
                if seat in date_seat_map:
                    logger.error(f"    Seat conflict: Seat {seat} already assigned")
                    continue

                area, subarea = seat_to_area[seat]
                date_seat_map[seat] = (area, agent['Name'], agent['Shift'])
                if seat in date_available_seats:
                    date_available_seats.remove(seat)
                logger.debug(f"    Assigned seat {seat} ({area}) to {agent['Name']} ({agent['Shift']})")

        # Store assignments for this date
        seat_assignments[date] = date_seat_map

        # Apply shift constraints for non-overlapping shifts
        if shift_constraints.get("non_overlapping_only", False):
            assigned_shifts = [info[2] for info in date_seat_map.values()]

            # For each future date, ensure no overlapping shifts
            for future_date in date_objects[i+1:]:
                future_date_str = future_date.strftime('%Y-%m-%d')
                future_agents = agent_df[agent_df['Date'] == future_date_str]

                non_overlapping_indices = []
                for idx, agent in future_agents.iterrows():
                    shift = agent['Shift']
                    overlap = any(shifts_overlap(shift, assigned_shift) for assigned_shift in assigned_shifts)
                    if not overlap:
                        non_overlapping_indices.append(idx)

                agent_df.loc[agent_df['Date'] == future_date_str, 'NonOverlapping'] = False
                agent_df.loc[non_overlapping_indices, 'NonOverlapping'] = True

    return seat_assignments

# === Generate Output ===
def generate_output(seat_assignments: Dict, output_file: str, logger: SeatAssignmentLogger):
    logger.info("Generating output file...")

    # Get all unique dates
    dates = sorted(seat_assignments.keys())
    date_strs = [date.strftime('%Y-%m-%d') for date in dates]

    # Get all unique areas and seats
    areas = set()
    seats = set()
    for date_assignments in seat_assignments.values():
        for seat, (area, _, _) in date_assignments.items():
            areas.add(area)
            seats.add(seat)

    areas = sorted(areas)
    seats = sorted(seats)

    # Create a DataFrame with areas and seats as rows
    rows = []
    for area in areas:
        rows.append({"Area/Seat": area, "Type": "Area"})
        for seat in seats:
            if SEAT_AREA_MAP.get(seat, ("", ""))[0] == area:
                rows.append({"Area/Seat": f"Seat {seat}", "Type": "Seat"})

    df = pd.DataFrame(rows)

    # Add columns for each date
    for date, date_str in zip(dates, date_strs):
        assignments = seat_assignments[date]

        # Create a column for this date
        df[date_str] = ""

        # Fill in the assignments
        for i, row in df.iterrows():
            if row['Type'] == "Seat":
                seat_num = int(row['Area/Seat'].split()[1])
                if seat_num in assignments:
                    area, name, shift = assignments[seat_num]
                    df.at[i, date_str] = f"{name} ({shift})"

    # Reorganize the DataFrame for better readability
    area_seat_tuples = []
    for _, row in df.iterrows():
        if row['Type'] == "Area":
            current_area = row['Area/Seat']
            area_seat_tuples.append((current_area, ""))
        else:
            seat_num = row['Area/Seat']
            area_seat_tuples.append((current_area, seat_num))

    index = pd.MultiIndex.from_tuples(area_seat_tuples, names=['Area', 'Seat'])
    final_df = pd.DataFrame(index=index, columns=date_strs)

    for date_str in date_strs:
        for i, row in df.iterrows():
            if row['Type'] == "Seat":
                seat_num = int(row['Area/Seat'].split()[1])
                final_df.at[(area_seat_tuples[i][0], area_seat_tuples[i][1]), date_str] = row[date_str]

    # Save to Excel
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            final_df.to_excel(writer, sheet_name='Seat Assignments')

            summary_data = {
                "Metric": ["Total Agents", "Total Seats", "Reserved Seats", "Total Dates", "Start Date", "End Date"],
                "Value": [
                    len(agent_df['ID'].unique()),
                    TOTAL_SEATS,
                    len(DEFAULT_RESERVED_SEATS),
                    len(dates),
                    dates[0].strftime('%Y-%m-%d'),
                    dates[-1].strftime('%Y-%m-%d')
                ]
            }
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)

            legend_data = {
                "Item": ["Area", "Seat", "Agent Assignment"],
                "Description": [
                    "The operational area (e.g., OPS1-A, OPS2)",
                    "The specific seat number within an area",
                    "Agent name followed by shift in parentheses (e.g., 'John Doe (Morning)')"
                ]
            }
            pd.DataFrame(legend_data).to_excel(writer, sheet_name='Legend', index=False)

        logger.info(f"Successfully generated output file: {output_file}")
        return True
    except Exception as e:
        logger.error(f"Error generating output file: {e}")
        return False

# === Main Function ===
def main(agent_file_path: str, rules_input: Any, output_file: str = "seat_assignments.xlsx",
         rules_output: str = None, model_name: str = None, log_file: str = None,
         log_level: str = "INFO", server_ip: str = OLLAMA_SERVER, server_port: int = OLLAMA_PORT):
    if not model_name:
        raise ValueError("Model name is required. Please specify an Ollama model.")

    if not configure_ollama_client(server_ip, server_port):
        print(f"Warning: Could not connect to Ollama server at {server_ip}:{server_port}. Using fallback parser.")

    logger = SeatAssignmentLogger(log_file=log_file, level=log_level)
    logger.info(f"Starting seat assignment process with model '{model_name}'")

    try:
        agent_df = pd.read_excel(agent_file_path)
        logger.info(f"Loaded {len(agent_df)} agent records from {agent_file_path}")

        required_columns = ['ID', 'Name', 'Date', 'Queue', 'Shift', 'Supervisor']
        missing_columns = [col for col in required_columns if col not in agent_df.columns]
        if missing_columns:
            logger.error(f"Missing required columns in input file: {missing_columns}")
            return False

        if 'Batch' not in agent_df.columns:
            logger.warning("Batch column not found. Using default batch value of 0 for all agents.")
            agent_df['Batch'] = 0

    except Exception as e:
        logger.error(f"Error loading agent data: {e}")
        return False

    data_analysis = analyze_agent_data(agent_df, logger)
    rules = parse_rules_with_ollama(rules_input, data_analysis, model_name, logger)

    if rules_output:
        try:
            with open(rules_output, 'w', encoding='utf-8') as f:
                json.dump(rules, f, indent=2)
            logger.info(f"Parsed rules saved to {rules_output}")
        except Exception as e:
            logger.error(f"Error saving parsed rules: {e}")

    seat_assignments = ai_assign_seats(agent_df, rules, data_analysis, logger)

    if not generate_output(seat_assignments, output_file, logger):
        return False

    logger.info("Seat assignment process completed successfully.")
    return True

# === Command Line Interface ===
def cli():
    parser = argparse.ArgumentParser(
        description="AI-Powered Seat Assignment System with Comprehensive Rules",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
Examples:
  python seater.py agents.xlsx "Group by supervisor, then queue. Night shifts in OPS1-D." -m llama3
  python seater.py agents.xlsx rules.txt -o assignments.xlsx -m mistral --rules-output parsed_rules.json
  python seater.py agents.xlsx "Reserve seat 61. Avoid seat 15 for John Doe." -m codellama --log assignment.log --server 172.16.30.202
        """
    )

    parser.add_argument(
        "agent_file",
        help="Path to Excel file containing agent schedules"
    )

    parser.add_argument(
        "rules",
        help="Natural language rules (text), or path to rules file (.txt, .json, .yaml)"
    )

    parser.add_argument(
        "-o", "--output",
        default="seat_assignments.xlsx",
        help="Output Excel file path (default: seat_assignments.xlsx)"
    )

    parser.add_argument(
        "--rules-output",
        help="Path to save parsed rules as JSON file"
    )

    parser.add_argument(
        "-m", "--model",
        required=True,
        help="Ollama model name to use for rule parsing (REQUIRED)"
    )

    parser.add_argument(
        "--log",
        help="Path to save detailed log file (default: auto-generated filename)"
    )

    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=LOG_LEVELS,
        help="Logging verbosity level (default: INFO)"
    )

    parser.add_argument(
        "--server",
        default=OLLAMA_SERVER,
        help=f"IP address of the Ollama server (default: {OLLAMA_SERVER})"
    )

    parser.add_argument(
        "--port",
        type=int,
        default=OLLAMA_PORT,
        help=f"Port of the Ollama server (default: {OLLAMA_PORT})"
    )

    args = parser.parse_args()

    if not os.path.exists(args.agent_file):
        print(f"Error: Agent file '{args.agent_file}' not found.")
        sys.exit(1)

    if os.path.exists(args.rules) and args.rules.endswith(('.txt', '.json', '.yaml', '.yml')):
        if not os.path.exists(args.rules):
            print(f"Error: Rules file '{args.rules}' not found.")
            sys.exit(1)

    success = main(
        agent_file_path=args.agent_file,
        rules_input=args.rules,
        output_file=args.output,
        rules_output=args.rules_output,
        model_name=args.model,
        log_file=args.log,
        log_level=args.log_level,
        server_ip=args.server,
        server_port=args.port
    )

    if not success:
        sys.exit(1)

if __name__ == "__main__":
    cli()