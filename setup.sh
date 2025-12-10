#!/bin/bash
# Setup script for Excel Refractor Flask Application

echo "====================================="
echo "Excel Refractor - Setup Script"
echo "====================================="
echo ""

# Check if Python 3 is installed
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed"
    echo "Installing Python 3..."
    sudo dnf install -y python3 python3-pip
fi

echo "Python version: $(python3 --version)"
echo ""

# Create virtual environment
echo "Creating virtual environment..."
if [ -d "venv" ]; then
    echo "Virtual environment already exists. Skipping..."
else
    python3 -m venv venv
    echo "Virtual environment created successfully!"
fi
echo ""

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate

# Upgrade pip
echo "Upgrading pip..."
pip install --upgrade pip

# Install requirements
echo "Installing Python dependencies..."
pip install -r requirements.txt

if [ $? -eq 0 ]; then
    echo ""
    echo "====================================="
    echo "Setup completed successfully!"
    echo "====================================="
    echo ""
    echo "To run the application:"
    echo "1. Activate virtual environment: source venv/bin/activate"
    echo "2. Run the application: ./run.sh"
    echo ""
    echo "Or use: bash run.sh (it will auto-activate venv)"
else
    echo ""
    echo "ERROR: Failed to install dependencies"
    exit 1
fi
