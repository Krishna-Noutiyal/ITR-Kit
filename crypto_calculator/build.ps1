# Enable ANSI escape codes in PowerShell
$esc = [char]27

function Write-Info($msg) {
    Write-Host "$esc[1;34m[INFO]$esc[0m $msg"
}
function Write-Success($msg) {
    Write-Host "$esc[1;32m[SUCCESS]$esc[0m $msg"
}
function Write-ErrorMsg($msg) {
    Write-Host "$esc[1;31m[ERROR]$esc[0m $msg"
}
function Write-Section($msg) {
    Write-Host ""
    Write-Host "$esc[1;36m======== $msg ========`n$esc[0m"
}

# Step 1: Install Requirements
Write-Section "Installing Python Packages"
try {
    .\install_requirments.ps1
    if ($LASTEXITCODE -ne 0) {
        Write-ErrorMsg "Dependency installation failed. Exiting."
        exit 1
    } else {
        Write-Success "All Python packages installed successfully."
    }
} catch {
    Write-ErrorMsg "Exception during requirements installation: $_"
    exit 1
}

clear

# Step 2: Build the Flet App
Write-Section "Building Flet Windows Application"

try {
    flet build windows `
    --project "CryptoAIS" `
    --company "Pooja ITR Center" `
    --description "Cryptocurrency Trade calculatore that updated Form-16 for Crypto Calculations" `
    --product "CryptoAIS" `
    --build-version "1.1.0" `
    --company "Pooja ITR Center" `
    --copyright "Copyright (C) 2025 Pooja ITR Center" `
    --exclude "release, icons, requirements.txt, README.md, certs, test, .venv" `
    --clear-cache --compile-app --cleanup-app --cleanup-packages --skip-flutter-doctor --module-name .\main.py

    if ($LASTEXITCODE -eq 0) {
        Write-Success "Flet Windows build completed successfully!"
    } else {
        Write-ErrorMsg "Flet build failed. Exit code: $LASTEXITCODE"
        exit 1
    }
} catch {
    Write-ErrorMsg "Exception during Flet build: $_"
    exit 1
}
