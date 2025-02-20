# === Define Paths and Variables ===
$PythonVersion       = "3.11.6"
$PythonZipUrl        = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-embed-amd64.zip"
$UserPythonFolder    = "$env:USERPROFILE\Python"
$PythonZipPath       = "$env:TEMP\python_embed.zip"
$PythonExe           = "$UserPythonFolder\python.exe"
$PipInstallerPath    = "$env:TEMP\get-pip.py"
$LogFolder           = "$UserPythonFolder\logs"
$LogFile             = "$LogFolder\python_install_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

# GitHub base URL for your repository (adjust as needed)
$GitHubBaseURL       = "https://github.com/genderlesspit/pdf_extractor"

# Local folder to store downloaded scripts
$ScriptsFolder       = "$env:TEMP\pdf_extractor_scripts"
if (!(Test-Path $ScriptsFolder)) { New-Item -ItemType Directory -Path $ScriptsFolder -Force | Out-Null }
$FixPipLocalPath     = Join-Path $ScriptsFolder "fix_pip.py"
$SetupLocalPath      = Join-Path $ScriptsFolder "setup.py"

# === Logging Function ===
if (!(Test-Path $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null }
function Log-Message {
    param ([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "[$Timestamp] $Message"
    Write-Host $Message
}

# === Step 1: Install Python (if not present) ===
Log-Message "Starting Python installation (No Admin Required)..."
if (Test-Path $PythonExe) {
    Log-Message "Python is already installed at $UserPythonFolder"
} else {
    Log-Message "Downloading Python ZIP package..."
    try {
        Invoke-WebRequest -Uri $PythonZipUrl -OutFile $PythonZipPath -ErrorAction Stop
        Expand-Archive -Path $PythonZipPath -DestinationPath $UserPythonFolder -Force
        Log-Message "Python installed at: $UserPythonFolder"
    } catch {
        Log-Message "Error installing Python: $_"
        exit
    }
}

# === Step 2: Add Python to User PATH if not present ===
$CurrentPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
if ($CurrentPath -notlike "*$UserPythonFolder*") {
    $NewPath = "$UserPythonFolder;$CurrentPath"
    [System.Environment]::SetEnvironmentVariable("Path", $NewPath, "User")
    Log-Message "Added Python to user PATH"
}

# === Step 3: Install pip (if not present) ===
$PipCheck = & "$PythonExe" -m pip --version 2>&1
if ($PipCheck -match "pip") {
    Log-Message "pip is already installed."
} else {
    Log-Message "Downloading and installing pip..."
    try {
        Invoke-WebRequest -Uri "https://bootstrap.pypa.io/get-pip.py" -OutFile $PipInstallerPath -ErrorAction Stop
        & "$PythonExe" $PipInstallerPath --no-cache-dir
        Log-Message "pip installed successfully."
    } catch {
        Log-Message "Error installing pip: $_"
        exit
    }
}

# === Step 4: Set user base for pip installations ===
[System.Environment]::SetEnvironmentVariable("PYTHONUSERBASE", "$UserPythonFolder", "User")

# === Step 5: Ensure 'requests' module is installed ===
Log-Message "Ensuring 'requests' module is available..."
try {
    & "$PythonExe" -m pip install --user --no-cache-dir requests | Out-Null
    Log-Message "'requests' installed successfully."
} catch {
    Log-Message "Error installing 'requests': $_"
    exit
}

# === Step 6: Download necessary scripts from GitHub using Python ===
# Create a temporary Python script to download files with requests (verify disabled)
$DownloadScriptPath = Join-Path $env:TEMP "download_files.py"
$DownloadScriptContent = @"
import os
import requests
import urllib3
# Disable insecure request warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

base_url = '$GitHubBaseURL'
# The following paths are interpolated by PowerShell
files_to_download = {
    'fix_pip.py': os.path.join(r'$ScriptsFolder', 'fix_pip.py'),
    'setup.py': os.path.join(r'$ScriptsFolder', 'setup.py')
}

for fname, local_path in files_to_download.items():
    url = f'{base_url}/{fname}'
    print(f'Downloading {fname} from {url}...')
    response = requests.get(url, verify=False)
    response.raise_for_status()
    with open(local_path, 'wb') as f:
        f.write(response.content)
    print(f'Downloaded {fname} to {local_path}')
"@
$DownloadScriptContent | Out-File -Encoding utf8 $DownloadScriptPath
Log-Message "Executing Python download script..."
try {
    & "$PythonExe" $DownloadScriptPath
    Log-Message "Downloaded fix_pip.py and setup.py successfully."
} catch {
    Log-Message "Error executing download script: $_"
    exit
}

# Check that files exist before running them
if (!(Test-Path $FixPipLocalPath)) {
    Log-Message "Error: fix_pip.py not found at $FixPipLocalPath"
    exit
}
if (!(Test-Path $SetupLocalPath)) {
    Log-Message "Error: setup.py not found at $SetupLocalPath"
    exit
}

# === Step 7: Run fix_pip.py to fix pip vendor issues ===
Log-Message "Running fix_pip.py..."
try {
    & "$PythonExe" $FixPipLocalPath
    Log-Message "fix_pip.py executed successfully."
} catch {
    Log-Message "Error executing fix_pip.py: $_"
    exit
}

# === Step 8: Run setup.py to complete application setup ===
Log-Message "Running setup.py..."
try {
    & "$PythonExe" $SetupLocalPath
    Log-Message "setup.py executed successfully."
} catch {
    Log-Message "Error executing setup.py: $_"
    exit
}
