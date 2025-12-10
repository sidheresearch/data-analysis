#!/bin/bash
# Run script for Excel Refractor Flask Application

echo "Starting Excel Refractor Application..."
echo ""

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "ERROR: Virtual environment not found!"
    echo "Please run setup.sh first: bash setup.sh"
    exit 1
fi

# Activate virtual environment
source venv/bin/activate

# Create necessary directories
mkdir -p uploads processed cache

# Run the Flask application
echo "Application starting on http://0.0.0.0:5000"
echo "Press Ctrl+C to stop the server"
echo ""

python app.py
