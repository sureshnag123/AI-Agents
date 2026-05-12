# Odoo to Tally Sync Setup (Windows)
# Run this script once to set up the workspace

Write-Host 'Setting up Odoo-Tally Sync Agent...' -ForegroundColor Cyan

# Create virtual environment
if (-Not (Test-Path '.venv')) {
    python -m venv .venv
    Write-Host 'Virtual environment created.' -ForegroundColor Green
}

# Activate
.\.venv\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements.txt
Write-Host 'Dependencies installed.' -ForegroundColor Green

# Create .env from example if it does not exist
if (-Not (Test-Path '.env')) {
    Copy-Item '.env.example' '.env'
    Write-Host '.env file created from .env.example - EDIT IT with your credentials!' -ForegroundColor Yellow
}

# Create .tmp directory
New-Item -ItemType Directory -Force -Path '.tmp' | Out-Null
New-Item -ItemType Directory -Force -Path '.tmp\logs' | Out-Null
Write-Host '.tmp directory ready.' -ForegroundColor Green

Write-Host ''
Write-Host 'Setup complete! Next steps:' -ForegroundColor Cyan
Write-Host '  1. Edit .env with your Odoo and Tally connection details'
Write-Host '  2. Enable Tally XML Server: F12 > Connectivity > Tally Prime Server = Yes'
Write-Host '  3. Test Odoo:  python execution/odoo_connector.py --test'
Write-Host '  4. Test Tally: python execution/tally_connector.py --test'
Write-Host '  5. Generate mapping: python execution/odoo_tally_sync.py --generate-mapping'
Write-Host '  6. Dry run: python execution/odoo_tally_sync.py --dry-run'
