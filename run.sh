#!/bin/bash
# PDF Manager Launcher Script

# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Check if virtual environment exists
if [ ! -d "$SCRIPT_DIR/.venv" ]; then
    echo "Virtual environment not found. Creating one..."
    python3 -m venv "$SCRIPT_DIR/.venv"
    echo "Installing dependencies..."
    "$SCRIPT_DIR/.venv/bin/pip" install -r "$SCRIPT_DIR/requirements.txt"
fi

# Run the application
"$SCRIPT_DIR/.venv/bin/python" "$SCRIPT_DIR/pdf_manager.py"
