# ANSI escape characters for colored output
$esc = [char]27
function Info($msg) { Write-Host "$esc[1;34m[INFO]$esc[0m $msg" }
function Success($msg) { Write-Host "$esc[1;32m[SUCCESS]$esc[0m $msg" }
function ErrorMsg($msg) { Write-Host "$esc[1;31m[ERROR]$esc[0m $msg" }

# === Setup ===
$branch = "3.2"
$openpyxlDir = "test\openpyxl"

# Step 1: Ensure pip & Mercurial are available
Info "Ensuring pip and Mercurial are installed..."
pip install --upgrade pip
pip install mercurial

# Step 2: Clone openpyxl 3.2 branch
if (Test-Path $openpyxlDir) {
    Info "Removing previous openpyxl clone..."
    Remove-Item -Recurse -Force $openpyxlDir
}

Info "Cloning openpyxl 3.2 branch..."
hg clone -b $branch https://foss.heptapod.net/openpyxl/openpyxl $openpyxlDir

# Step 3: Install openpyxl from local source
Set-Location $openpyxlDir
Info "Installing openpyxl from 3.2 branch source..."
pip install .

# Return to project root
Set-Location ../..

# Step 4: Clean up
Info "Cleaning up cloned repository..."
Remove-Item -Recurse -Force $openpyxlDir

Success "OpenPyXL from 3.2 branch installed successfully."
# Step 6: Install other dependencies
Info "Installing additional project dependencies..."

$requirements = Get-Content "req.txt" -Raw

pip install -r req.txt

Success "All packages (including openpyxl from topic branch) installed successfully."