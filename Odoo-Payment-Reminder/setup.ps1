# Odoo Payment Reminder Agent — Setup Script
# Run: .\setup.ps1

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Odoo Payment Reminder Agent Setup"
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check Python
$pythonCmd = $null
foreach ($cmd in @("python", "python3", "py")) {
    if (Get-Command $cmd -ErrorAction SilentlyContinue) {
        $ver = & $cmd --version 2>&1
        if ($ver -match "Python 3") {
            $pythonCmd = $cmd
            Write-Host "✓ Found $ver" -ForegroundColor Green
            break
        }
    }
}

if (-not $pythonCmd) {
    Write-Host "✗ Python 3 not found. Please install from https://python.org" -ForegroundColor Red
    exit 1
}

# Create virtual environment
if (-not (Test-Path "venv")) {
    Write-Host "Creating virtual environment..." -ForegroundColor Yellow
    & $pythonCmd -m venv venv
    Write-Host "✓ Virtual environment created" -ForegroundColor Green
} else {
    Write-Host "✓ Virtual environment already exists" -ForegroundColor Green
}

# Activate venv
Write-Host "Activating virtual environment..." -ForegroundColor Yellow
& .\venv\Scripts\Activate.ps1

# Install dependencies
Write-Host "Installing dependencies..." -ForegroundColor Yellow
pip install -r requirements.txt
Write-Host "✓ Dependencies installed" -ForegroundColor Green

# Create .env from example if not exists
if (-not (Test-Path ".env")) {
    Copy-Item ".env.example" ".env"
    Write-Host "✓ Created .env from .env.example" -ForegroundColor Green
    Write-Host ""
    Write-Host "⚠  IMPORTANT: Edit .env with your Odoo connection details!" -ForegroundColor Yellow
    Write-Host "   Required fields:" -ForegroundColor Yellow
    Write-Host "     ODOO_URL      = https://your-company.odoo.com" -ForegroundColor White
    Write-Host "     ODOO_DB       = your-database-name" -ForegroundColor White
    Write-Host "     ODOO_USERNAME  = admin@yourcompany.com" -ForegroundColor White
    Write-Host "     ODOO_PASSWORD  = your-api-key" -ForegroundColor White
} else {
    Write-Host "✓ .env already exists" -ForegroundColor Green
}

# Create .tmp directory
New-Item -ItemType Directory -Force -Path ".tmp" | Out-Null

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Setup Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor White
Write-Host "  1. Edit .env with your Odoo credentials" -ForegroundColor White
Write-Host "  2. Test connection:  python execution/odoo_connector.py" -ForegroundColor White
Write-Host "  3. Dry run:          python execution/send_payment_reminders.py --mode overdue --dry-run" -ForegroundColor White
Write-Host "  4. Send reminders:   python execution/send_payment_reminders.py --mode overdue" -ForegroundColor White
Write-Host "  5. Schedule daily:   python execution/schedule_reminders.py --hour 9" -ForegroundColor White
Write-Host ""
