# Fracktal Works HR - Automated Setup Script
# Run this script to set up the agent environment from scratch
# Usage: .\setup.ps1

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "Setting up: Fracktal Works HR" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Check Python is installed
$pythonVersion = python --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Python is not installed or not in PATH" -ForegroundColor Red
    Write-Host "Please install Python 3.10+ from https://www.python.org/downloads/" -ForegroundColor Yellow
    exit 1
}
Write-Host "✓ Found $pythonVersion" -ForegroundColor Green

# Create virtual environment if it doesn't exist
if (-not (Test-Path ".venv")) {
    Write-Host "Creating virtual environment..." -ForegroundColor Yellow
    python -m venv .venv
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: Failed to create virtual environment" -ForegroundColor Red
        exit 1
    }
    Write-Host "✓ Virtual environment created" -ForegroundColor Green
} else {
    Write-Host "✓ Virtual environment already exists" -ForegroundColor Green
}

# Activate virtual environment
Write-Host "Activating virtual environment..." -ForegroundColor Yellow
& .\.venv\Scripts\Activate.ps1

# Upgrade pip
Write-Host "Upgrading pip..." -ForegroundColor Yellow
python -m pip install --upgrade pip --quiet

# Install requirements
Write-Host "Installing dependencies..." -ForegroundColor Yellow
if (Test-Path "requirements.txt") {
    pip install -r requirements.txt --quiet
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: Failed to install requirements" -ForegroundColor Red
        exit 1
    }
    Write-Host "✓ Dependencies installed" -ForegroundColor Green
} else {
    Write-Host "WARNING: requirements.txt not found" -ForegroundColor Yellow
}

# Copy .env.example to .env if .env doesn't exist
if (-not (Test-Path ".env")) {
    if (Test-Path ".env.example") {
        Copy-Item ".env.example" ".env"
        Write-Host "✓ Created .env from .env.example" -ForegroundColor Green
        Write-Host "  (Only needed if you use the optional Google Sheets export)" -ForegroundColor Yellow
    }
} else {
    Write-Host "✓ .env file already exists" -ForegroundColor Green
}

# Create .tmp directory if it doesn't exist
if (-not (Test-Path ".tmp")) {
    New-Item -ItemType Directory -Path ".tmp" | Out-Null
    Write-Host "✓ Created .tmp directory" -ForegroundColor Green
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "✓ SETUP COMPLETE!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor White
Write-Host "  1. Run: python execution/generate_hr_master_workbook.py --fy `"2026-27`"" -ForegroundColor White
Write-Host "  2. Fill in Statutory_Compliance and Master_Employees with real data" -ForegroundColor White
Write-Host "  3. Open this folder in VS Code" -ForegroundColor White
Write-Host "  4. Select 'Fracktal Works HR' from Copilot Chat agent dropdown" -ForegroundColor White
Write-Host ""
Write-Host "To activate the environment manually:" -ForegroundColor Gray
Write-Host "  .\.venv\Scripts\Activate.ps1" -ForegroundColor Gray
Write-Host ""
