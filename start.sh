#!/bin/bash

echo "========================================"
echo "    Attendance Report Tool"
echo "========================================"
echo ""
echo "Starting the application..."
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ Error: Python 3 is not installed or not in PATH"
    echo "Please install Python 3.7 or higher"
    exit 1
fi

# Make the script executable
chmod +x run.py

# Run the application
python3 run.py 