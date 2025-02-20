# Define parameters for script execution
param (
    [string]$Action,
    [string]$JsonInput
)

# Define paths
$BaseDir = "$env:USERPROFILE\PDF_Extractor"
$GridSaveFile = "$BaseDir\saved_grid.csv"

# Ensure base directory exists
if (!(Test-Path $BaseDir)) { New-Item -ItemType Directory -Path $BaseDir -Force | Out-Null }

# Function to save grid data
function Save-Grid {
    param ([string]$JsonData)

    try {
        # Convert JSON input to PowerShell object
        $GridData = $JsonData | ConvertFrom-Json

        # Write to CSV
        $GridData | Export-Csv -Path $GridSaveFile -NoTypeInformation -Encoding UTF8
        Write-Host "✅ Grid successfully saved to $GridSaveFile"
    } catch {
        Write-Host "❌ Error saving grid: $_"
    }
}

# Function to load grid data
function Load-Grid {
    if (!(Test-Path $GridSaveFile)) {
        Write-Host "⚠ No saved grid file found."
        return
    }

    try {
        # Read CSV and convert to JSON
        $GridData = Import-Csv -Path $GridSaveFile
        $JsonOutput = $GridData | ConvertTo-Json -Depth 1
        Write-Output $JsonOutput
    } catch {
        Write-Host "❌ Error loading grid: $_"
    }
}

# Determine action based on argument
if ($Action -eq "load") {
    Load-Grid
} elseif ($Action -eq "save") {
    Save-Grid -JsonData $JsonInput
} else {
    Write-Host "❌ No valid action specified. Use 'save' or 'load'."
}
