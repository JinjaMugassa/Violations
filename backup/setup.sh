#!/bin/bash
# Quick setup script for Wialon Violation Reports system

echo "=========================================="
echo "Wialon Violation Reports - Setup"
echo "=========================================="
echo ""

# Check Python version
echo "Checking Python version..."
python_version=$(python3 --version 2>&1 | awk '{print $2}')
echo "Found Python $python_version"
echo ""

# Create virtual environment
echo "Creating virtual environment..."
python3 -m venv venv
echo "✓ Virtual environment created"
echo ""

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate
echo "✓ Virtual environment activated"
echo ""

# Install dependencies
echo "Installing dependencies..."
pip install --upgrade pip
pip install -r requirements.txt
echo "✓ Dependencies installed"
echo ""

# Create .env file if it doesn't exist
if [ ! -f .env ]; then
    echo "Creating .env file from template..."
    cp .env.example .env
    echo "✓ .env file created"
    echo ""
    echo "⚠️  IMPORTANT: Edit .env file and add your Wialon token!"
    echo ""
else
    echo "✓ .env file already exists"
    echo ""
fi

# Create output directory
echo "Creating output directory..."
mkdir -p violation_reports_raw
echo "✓ Output directory created"
echo ""

echo "=========================================="
echo "Setup Complete!"
echo "=========================================="
echo ""
echo "Next steps:"
echo "1. Edit .env file and add your WIALON_TOKEN"
echo "2. Activate virtual environment: source venv/bin/activate"
echo "3. Run reports: python run_pull_violation.py"
echo ""