# Define paths
$ProgramDir = "$env:USERPROFILE\PDF_Extractor"
$DependenciesDir = "$ProgramDir\dependencies"
$DownloadURL = "https://github.com/genderlesspit/pdf_extractor/archive/refs/heads/main.zip"
$ZipFile = "$env:TEMP\pdf_extractor.zip"
$ExtractPath = "$env:TEMP\pdf_extractor-main"
$CSVFile = "$ProgramDir\temp_grid_update.csv"

# Ensure base directory exists
if (!(Test-Path $ProgramDir)) {
    Write-Host "📂 Creating program directory: $ProgramDir"
    New-Item -ItemType Directory -Path $ProgramDir -Force | Out-Null
}

# Function to download dependencies
function Download-Dependencies {
    Write-Host "📥 Downloading latest dependencies from GitHub..."
    try {
        Invoke-WebRequest -Uri $DownloadURL -OutFile $ZipFile -ErrorAction Stop
        Write-Host "✅ Download complete: $ZipFile"
    } catch {
        Write-Host "❌ Error downloading dependencies: $_"
        exit 1
    }
}

# Function to clean old dependencies
function Clean-OldDependencies {
    Write-Host "🧹 Removing old dependencies..."
    if (Test-Path $DependenciesDir) {
        Remove-Item -Path $DependenciesDir -Recurse -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }
    New-Item -ItemType Directory -Path $DependenciesDir -Force | Out-Null
    Write-Host "✅ Old dependencies removed."
}

# Function to extract new dependencies
function Extract-Dependencies {
    Write-Host "📂 Extracting new dependencies..."
    try {
        Expand-Archive -Path $ZipFile -DestinationPath $env:TEMP -Force
        if (Test-Path $ExtractPath) {
            Move-Item -Path "$ExtractPath\*" -Destination $DependenciesDir -Force
            Write-Host "✅ Dependencies installed to $DependenciesDir"
        } else {
            Write-Host "❌ Extraction failed: Directory not found."
            exit 1
        }
    } catch {
        Write-Host "❌ Error extracting dependencies: $_"
        exit 1
    }
}

# Function to create temp_grid_update.csv if missing
function Ensure-CSVFile {
    if (!(Test-Path $CSVFile)) {
        Write-Host "📄 Creating missing CSV file: $CSVFile"
        "Master PDF,Output Name,Page Range" | Out-File -Encoding UTF8 $CSVFile
        Write-Host "✅ CSV file created successfully."
    } else {
        Write-Host "✅ CSV file already exists."
    }
}

# Execute update process
Download-Dependencies
Clean-OldDependencies
Extract-Dependencies
Ensure-CSVFile  # <-- NEW STEP ADDED

# Cleanup temporary files
Remove-Item -Path $ZipFile -Force
Remove-Item -Path $ExtractPath -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "🚀 Dependencies updated successfully! You're ready to run the program."

# Ensure $ProgramDir is set
if (-not $ProgramDir -or $ProgramDir -eq "") {
    Write-Host "❌ Error: Program directory is not set correctly."
    exit 1
}

# Define Paths (Matching setup.ps1)
$ProgramDir = "$env:USERPROFILE\PDF_Extractor"
$DependenciesDir = "$ProgramDir\dependencies"
$GUIPath = Join-Path $DependenciesDir "GUI.ps1"  # Corrected path

# Check if GUI.ps1 Exists
if (!(Test-Path $GUIPath)) {
    Write-Host "❌ GUI script not found at expected location: $GUIPath"
    [System.Windows.Forms.MessageBox]::Show("GUI script not found: $GUIPath", "Error", "OK", "Error")
    exit 1
}

Write-Host "🚀 Running GUI script from: $GUIPath"

# Read `GUI.ps1` as a script block and execute
try {
    $GUIContent = Get-Content -Path $GUIPath -Raw
    $ScriptBlock = [scriptblock]::Create($GUIContent)
    & $ScriptBlock  # Run the script block in the current session
    Write-Host "✅ GUI executed successfully!"
} catch {
    Write-Host "❌ Error executing GUI script: $($_.Exception.Message)"
    [System.Windows.Forms.MessageBox]::Show("Error executing GUI script: $($_.Exception.Message)", "Error", "OK", "Error")
    exit 1
}

