#!/usr/bin/env python3
"""
Test script to verify the add_agents_from_roster.py functionality
"""

import subprocess
import sys
import os

def test_dry_run():
    """Test the script in dry-run mode to see what would be processed"""
    print("=== Testing dry-run mode ===")
    try:
        result = subprocess.run([
            sys.executable, 'add_agents_from_roster.py', '--dry-run'
        ], capture_output=True, text=True, cwd=os.getcwd())

        print("STDOUT:")
        print(result.stdout)

        if result.stderr:
            print("STDERR:")
            print(result.stderr)

        if result.returncode == 0:
            print("✓ Dry-run test completed successfully")
            return True
        else:
            print(f"✗ Dry-run test failed with return code {result.returncode}")
            return False

    except Exception as e:
        print(f"✗ Error running dry-run test: {e}")
        return False

def test_help():
    """Test the script help functionality"""
    print("\n=== Testing help functionality ===")
    try:
        result = subprocess.run([
            sys.executable, 'add_agents_from_roster.py', '--help'
        ], capture_output=True, text=True, cwd=os.getcwd())

        print("STDOUT:")
        print(result.stdout)

        if result.returncode == 0:
            print("✓ Help test completed successfully")
            return True
        else:
            print(f"✗ Help test failed with return code {result.returncode}")
            return False

    except Exception as e:
        print(f"✗ Error running help test: {e}")
        return False

if __name__ == "__main__":
    print("Testing add_agents_from_roster.py functionality...")

    # Test help first
    help_success = test_help()

    # Test dry-run
    dry_run_success = test_dry_run()

    if help_success and dry_run_success:
        print("\n✓ All tests completed successfully!")
        print("\nThe script is ready to use. To run it:")
        print("  1. For a dry run (show what would be processed):")
        print("     python add_agents_from_roster.py --dry-run")
        print("  2. To actually process and insert data:")
        print("     python add_agents_from_roster.py")
        print("  3. To specify a specific roster file:")
        print("     python add_agents_from_roster.py --file path/to/your_roster.xlsx")
    else:
        print("\n✗ Some tests failed. Please check the output above.")
        sys.exit(1)
