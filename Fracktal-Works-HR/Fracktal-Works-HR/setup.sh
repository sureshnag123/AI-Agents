#!/bin/bash
# Fracktal Works HR - Automated Setup Script
# Run this script to set up the agent environment from scratch
# Usage: ./setup.sh

echo "============================================"
echo "Setting up: Fracktal Works HR"
echo "============================================"
echo ""

# Check Python is installed
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python3 is not installed"
    echo "Please install Python 3.10+ from https://www.python.org/downloads/"
    exit 1
fi
echo "✓ Found $(python3 --version)"

# Create virtual environment if it doesn't exist
if [ ! -d ".venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv .venv
    echo "✓ Virtual environment created"
else
    echo "✓ Virtual environment already exists"
fi

# Activate virtual environment
echo "Activating virtual environment..."
source .venv/bin/activate

# Upgrade pip
echo "Upgrading pip..."
pip install --upgrade pip --quiet

# Install requirements
echo "Installing dependencies..."
if [ -f "requirements.txt" ]; then
    pip install -r requirements.txt --quiet
    echo "✓ Dependencies installed"
else
    echo "WARNING: requirements.txt not found"
fi

# Copy .env.example to .env if .env doesn't exist
if [ ! -f ".env" ]; then
    if [ -f ".env.example" ]; then
        cp .env.example .env
        echo "✓ Created .env from .env.example"
        echo "  (Only needed if you use the optional Google Sheets export)"
    fi
else
    echo "✓ .env file already exists"
fi

# Create .tmp directory if it doesn't exist
mkdir -p .tmp
echo "✓ Created .tmp directory"

echo ""
echo "============================================"
echo "✓ SETUP COMPLETE!"
echo "============================================"
echo ""
echo "Next steps:"
echo "  1. Run: python execution/generate_hr_master_workbook.py --fy \"2026-27\""
echo "  2. Fill in Statutory_Compliance and Master_Employees with real data"
echo "  3. Open this folder in VS Code"
echo "  4. Select 'Fracktal Works HR' from Copilot Chat agent dropdown"
echo ""
echo "To activate the environment manually:"
echo "  source .venv/bin/activate"
echo ""
