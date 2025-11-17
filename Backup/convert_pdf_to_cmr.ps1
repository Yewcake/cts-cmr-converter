# PDF to CMR Converter - PowerShell Script
# Converts CTS Packing List PDFs to CMR Excel documents

param(
    [Parameter(Mandatory=$true)]
    [string]$Input,
    
    [Parameter(Mandatory=$false)]
    [string]$TemplatePath = "CTS_NL_CMR_Template.xlsx",
    
    [Parameter(Mandatory=$false)]
    [string]$OutputDir = ".\output"
)

# Ensure output directory exists
if (!(Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

# Check if Python is installed
$pythonCmd = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCmd) {
    Write-Host "Error: Python is not installed or not in PATH" -ForegroundColor Red
    Write-Host "Please install Python 3.8 or higher from https://www.python.org" -ForegroundColor Yellow
    exit 1
}

# Check if required Python packages are installed
Write-Host "Checking Python dependencies..." -ForegroundColor Cyan

$packagesToInstall = @("pdfplumber", "openpyxl")
$missingPackages = @()

foreach ($package in $packagesToInstall) {
    $installed = python -c "import $package" 2>$null
    if ($LASTEXITCODE -ne 0) {
        $missingPackages += $package
    }
}

if ($missingPackages.Count -gt 0) {
    Write-Host "Installing missing packages: $($missingPackages -join ', ')" -ForegroundColor Yellow
    python -m pip install $missingPackages --quiet
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error: Failed to install required packages" -ForegroundColor Red
        exit 1
    }
}

# Determine if input is a file or packing list number
if (Test-Path $Input -PathType Leaf) {
    $pdfPath = $Input
    $packingListNo = [System.IO.Path]::GetFileNameWithoutExtension($Input) -replace 'Packing_List_', ''
} else {
    # Assume it's a packing list number
    $packingListNo = $Input
    $pdfPath = "Packing_List_$packingListNo.pdf"
    
    if (!(Test-Path $pdfPath)) {
        Write-Host "Error: PDF file not found at $pdfPath" -ForegroundColor Red
        exit 1
    }
}

# Generate output filename
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = Join-Path $OutputDir "CMR_${packingListNo}_${timestamp}.xlsx"

# Call the Python script
Write-Host "`nProcessing packing list $packingListNo..." -ForegroundColor Cyan
Write-Host "  PDF: $pdfPath" -ForegroundColor Gray
Write-Host "  Output: $outputFile" -ForegroundColor Gray

# Get the directory where this script is located
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonScript = Join-Path $scriptDir "pdf_to_cmr.py"

if (!(Test-Path $pythonScript)) {
    Write-Host "Error: pdf_to_cmr.py not found in script directory" -ForegroundColor Red
    exit 1
}

# Execute Python script
python $pythonScript $pdfPath

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n✓ Success! CMR document created" -ForegroundColor Green
    
    # Move output file to output directory if it's in current directory
    $currentDirOutput = "CMR_*_*.xlsx"
    $createdFile = Get-ChildItem $currentDirOutput | Select-Object -First 1
    if ($createdFile) {
        Move-Item $createdFile.FullName $outputFile -Force
        Write-Host "  Saved to: $outputFile" -ForegroundColor Green
        
        # Ask if user wants to open the file
        $response = Read-Host "`nWould you like to open the file? (Y/N)"
        if ($response -eq 'Y' -or $response -eq 'y') {
            Start-Process $outputFile
        }
    }
} else {
    Write-Host "`n✗ Error: Failed to process packing list" -ForegroundColor Red
    exit 1
}
