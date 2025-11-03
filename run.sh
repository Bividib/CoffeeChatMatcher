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
# 5. Create a Python virtual environment and install libraries (silently).
# 6. Run the Python script, passing along an optional file path.
#

# --- 1. Stop the script on any error ---
set -e

echo "☕ Starting the KBC Coffee Pair setup..."
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
    echo "❌ Error: Python 3 is not installed."
    echo "Please install Python 3 to continue."
    exit 1
else
    echo "✅ Python 3 found."
fi

# --- 4. Create requirements.txt (silently) ---
# This uses a 'here document' to write the lines
# between <<EOF and EOF into the file.
cat > requirements.txt << EOF
pandas
openpyxl
EOF

# --- 5. Create Virtual Environment and Install (silently) ---
echo "Setting up environment and installing libraries... (This may take a moment)"

# Create venv and redirect all output to null
python3 -m venv venv > /dev/null 2>&1

# Activate the virtual environment
source venv/bin/activate

# Install dependencies and redirect all output to null
pip install -r requirements.txt > /dev/null 2>&1

echo "✅ Environment is ready."

# --- 6. Run the Python Script ---
echo "----------------------------------------"
echo "🚀 Running the coffee pair script..."
echo ""

# The Python script to run
PYTHON_SCRIPT="create_coffee_pairs.py"

# Check if the user provided an argument (a file path)
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
echo "🎉 Pairing complete!"
echo "Check the output above"
