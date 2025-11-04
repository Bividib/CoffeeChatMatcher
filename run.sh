#!/bin/bash
#
# This script automates the setup and running of the
# KBC coffee pairing script for macOS and Linux.
#
# It will:
# 1. Stop if any command fails (set -e).
# 2. Find the directory this script is in.
# 3. Check that Python 3 is installed.
# 4. Create a 'requirements.txt' file.
# 5. Check for venv, create if missing, then install libraries.
# 6. Run the Python script, passing along an optional file path.
#

# --- 1. Stop the script on any error ---
set -e

echo "Starting the KBC Coffee Pair setup..."
echo "----------------------------------------"

# --- 2. Find the script's own directory ---
# This ensures all files (venv, requirements.txt)
# are created next to this script, not where the user
# is currently in their terminal.
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" &> /dev/null && pwd)"
cd "$SCRIPT_DIR"

echo "Running in project directory: $SCRIPT_DIR"

# --- 3. Check for Python 3 ---
if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python 3 is not installed."
    echo "Please install Python 3 to continue."
    exit 1
else
    echo "[OK] Python 3 found."
fi


# --- 5. Create Virtual Environment and Install ---
echo "Setting up environment and installing libraries... (This may take a moment)"

# Detect OS and set the correct venv activation path
# MINGW or MSYS indicates Git Bash on Windows
if [[ "$(uname -s)" == *"MINGW"* ]] || [[ "$(uname -s)" == *"MSYS"* ]]; then
    ACTIVATE_SCRIPT="venv/Scripts/activate"
else
    ACTIVATE_SCRIPT="venv/bin/activate"
fi

# Check if the venv activation script already exists
# [ ! -f "path" ] means "if this file does NOT exist"
if [ ! -f "$ACTIVATE_SCRIPT" ]; then
    echo "  Creating new virtual environment (one-time setup)..."
    # --- MODIFIED ---
    # We WANT to see the error if this fails.
    python3 -m venv venv
else
    echo "  Found existing virtual environment."
fi

# Activate the virtual environment using the correct path
source "$ACTIVATE_SCRIPT"

# --- MODIFIED ---
# We WANT to see what's being installed.
echo "  Installing/checking requirements..."
pip install -r requirements.txt

echo "[OK] Environment is ready."

# --- 6. Run the Python Script ---
echo "----------------------------------------"
echo "Running the coffee pair script..."
echo ""

# The Python script to run
PYTHON_SCRIPT="create_coffee_pairs.py"

# Check if the user provided an argument (a file path)
# [ -n "$1" ] means "if $1 is not empty string"
if [ -n "$1" ]; then
    # If $1 exists, pass it to the Python script.
    # The quotes are vital to handle paths with spaces.
    echo "Using provided Excel file: $1"
    python "$PYTHON_SCRIPT" "$1"
else
    # If $1 does not exist, run the script with no arguments.
    # The Python script will then search for its own .xlsx file.
    echo "No Excel file provided. Python script will search for one..."
    python "$PYTHON_SCRIPT"
fi

echo ""
echo "----------------------------------------"
echo "Pairing complete!"
echo "Check the output above"

