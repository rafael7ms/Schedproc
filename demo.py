#!/usr/bin/env python3
"""
Demo script for the Roster Database Solution
"""

import os
import sys
import pandas as pd
from roster_parser import main as process_excel_file

def main():
    # Check if the roster file exists
    roster_file = "Roster - December  2025 updated TC.xlsx"
    
    if not os.path.exists(roster_file):
        print(f"Error: {roster_file} not found in the current directory.")
        print("Please ensure the roster file is in the same directory as this script.")
        sys.exit(1)
    
    print(f"Processing {roster_file}...")
    print("This may take a moment depending on the file size.")
    
    try:
        # Process the Excel file by calling the main function with the file as argument
        sys.argv = [sys.argv[0], roster_file]  # Set up argv to pass the file to the main function
        process_excel_file()
        
        # Display results
        print("\nProcessing complete!")
        print("\nDatabase has been updated with the latest roster information.")
        print("Duplicate records were automatically detected and skipped.")
        
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
