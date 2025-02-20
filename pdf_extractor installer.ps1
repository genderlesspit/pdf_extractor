# === Define Paths ===
$PythonVersion = "3.11.6"
$PythonZipUrl = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-embed-amd64.zip"
$UserPythonFolder = "$env:USERPROFILE\\Python"
$PythonZipPath = "$env:TEMP\\python_embed.zip"
$LogFolder = "$UserPythonFolder\\logs"
$LogFile = "$LogFolder\\python_install_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$PythonExe = "$UserPythonFolder\\python.exe"
$PipInstallerPath = "$env:TEMP\\get-pip.py"
$OutputEncoding = [System.Text.Encoding]::UTF8

# === Define Modular Dependencies ===
$DefaultDependencies = @("pypdf", "pdf2image", "pillow")  # Default set of dependencies
$CustomDependencies = @()  # Placeholder for user-defined dependencies

# Ensure log directory exists
if (!(Test-Path $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null }

# Function to log messages
function Log-Message {
    param ([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "[$Timestamp] $Message"
    Write-Host $Message
}

# === Initialize Dependencies ===
function Initialize-Dependencies {
    param ([string[]]$ExtraDependencies)

    $Dependencies = $DefaultDependencies + $ExtraDependencies
    return $Dependencies
}

# === Start Script Execution ===
Log-Message "Starting Python installation (No Admin Required)..."

# === Step 1: Check if Python is already installed ===
if (Test-Path $PythonExe) {
    Log-Message "Python is already installed at $UserPythonFolder"
} else {
    Log-Message "Downloading Python ZIP package..."
    try {
        Invoke-WebRequest -Uri $PythonZipUrl -OutFile $PythonZipPath
        Expand-Archive -Path $PythonZipPath -DestinationPath $UserPythonFolder -Force
        Log-Message "Python installed at: $UserPythonFolder"
    } catch { Log-Message "Error installing Python: $_"; exit }
}

# === Step 2: Add Python to User PATH if not present ===
$CurrentPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
if ($CurrentPath -notlike "*$UserPythonFolder*") {
    $NewPath = "$UserPythonFolder;$CurrentPath"
    [System.Environment]::SetEnvironmentVariable("Path", $NewPath, "User")
    Log-Message "Added Python to user PATH"
}

# === Step 3: Check if pip is installed ===
$PipCheck = & "$PythonExe" -m pip --version 2>&1
if ($PipCheck -match "pip") {
    Log-Message "pip is already installed."
} else {
    Log-Message "Downloading and installing pip..."
    try {
        Invoke-WebRequest -Uri "https://bootstrap.pypa.io/get-pip.py" -OutFile $PipInstallerPath
        & "$PythonExe" $PipInstallerPath --no-cache-dir
        Log-Message "pip installed successfully."
    } catch { Log-Message "Error installing pip: $_"; exit }
}

# === Ensure pip recognizes the user directory for installation ===
[System.Environment]::SetEnvironmentVariable("PYTHONUSERBASE", "$UserPythonFolder", "User")

# Verify pip is working in user mode
$PipUserCheck = & "$PythonExe" -m site --user-site 2>&1
if ($PipUserCheck -match "$UserPythonFolder") {
    Log-Message "✅ pip is now set to install packages in the user directory: $PipUserCheck"
} else {
    Log-Message "⚠ pip user directory not recognized. Forcing target install mode."
}

# === Step 4: Install Required Dependencies ===
$DependenciesToInstall = Initialize-Dependencies -ExtraDependencies $CustomDependencies

Log-Message "Checking and installing required dependencies..."
$MissingPackages = & "$PythonExe" -c "
import sys
try:
    import pypdf
    import pdf2image
    from PIL import Image
    print('')
except ModuleNotFoundError as e:
    print(e.name)" 2>&1

if ($MissingPackages -eq "") {
    Log-Message "All dependencies are already installed."
} else {
    foreach ($Package in $MissingPackages -split "`n") {
        if ($DependenciesToInstall -contains $Package) {
            Log-Message "Installing $Package..."
            $InstallOutput = & "$PythonExe" -m pip install --user --no-cache-dir $Package 2>&1
            if ($InstallOutput -match "Successfully installed") {
                Log-Message "Successfully installed $Package."
            } else {
                Log-Message "Error installing ${Package}: $InstallOutput"
            }
        }
    }
}

Log-Message "Python setup completed."

### === Define Directories Properly ===
$FallbackDir = "C:\PDF_Extractor"

try {
    $UserDocuments = [System.Environment]::GetFolderPath("MyDocuments")
    if ([string]::IsNullOrWhiteSpace($UserDocuments)) {
        throw "Unable to determine user 'Documents' folder."
    }
} catch {
    Write-Host "⚠ Warning: Failed to fetch 'Documents' directory. Using fallback location: $FallbackDir"
    $UserDocuments = $FallbackDir
}

# Allow custom base directory from config.txt
$ProgramDir = Join-Path $UserDocuments "PDF_Extractor"
$MasterPDFDir = Join-Path $ProgramDir "Master_PDFs"
$ExportsDir = Join-Path $ProgramDir "Exports"
$LogDir = Join-Path $ProgramDir "Logs"

### === Load Custom Directory from config.txt (if available) ===
$ConfigPath = ".\config.txt"

if (Test-Path $ConfigPath) {
    try {
        # Read and parse config file
        $ConfigContent = Get-Content $ConfigPath | ConvertFrom-StringData
        $CustomPath = $ConfigContent["ProgramDir"].Trim()

        if (![string]::IsNullOrWhiteSpace($CustomPath) -and (Test-Path $CustomPath)) {
            $ProgramDir = $CustomPath
            $MasterPDFDir = Join-Path $ProgramDir "Master_PDFs"
            $ExportsDir = Join-Path $ProgramDir "Exports"
            $LogDir = Join-Path $ProgramDir "Logs"
            Write-Host "📂 Using custom Program Directory: $ProgramDir"
        } else {
            Write-Host "⚠ Invalid or missing directory in config.txt. Resetting to default."
        }
    }
    catch {
        Write-Host "❌ Error reading config.txt: $_. Using default directories."
    }
}

### === Ensure all necessary directories exist ===
@($ProgramDir, $MasterPDFDir, $ExportsDir, $LogDir) | ForEach-Object {
    try {
        if (!(Test-Path $_)) { 
            if (-not (Test-Path (Split-Path $_))) {
                throw "Parent directory does not exist: $(Split-Path $_)"
            }
            New-Item -ItemType Directory -Path $_ -Force -ErrorAction Stop | Out-Null
            Write-Host "✅ Created directory: $_"
        }
    }
    catch {
        Write-Host "❌ Error creating directory '$_': $($_.Exception.Message)"
    }
}

# === Define User-Writable Directory ===
$DependenciesDir = Join-Path $ProgramDir "dependencies"
$GitHubTarUrl = "https://github.com/genderlesspit/pdf_extractor/archive/refs/tags/testtwo.tar.gz"
$TarFilePath = Join-Path $DependenciesDir "test.tar.gz"

# Ensure dependencies directory exists
if (!(Test-Path $DependenciesDir)) { 
    New-Item -ItemType Directory -Path $DependenciesDir -Force | Out-Null 
    Write-Host "✅ Created dependencies directory: $DependenciesDir"
}

# === Download `test.tar.gz` from GitHub ===
Write-Host "📥 Downloading test.tar.gz from GitHub..."
try {
    Invoke-WebRequest -Uri $GitHubTarUrl -OutFile $TarFilePath -UseBasicParsing -ErrorAction Stop
    Write-Host "✅ Successfully downloaded test.tar.gz to: $TarFilePath"
} catch {
    Write-Host "❌ Failed to download test.tar.gz: $_"
    exit
}

# Verify the file exists after download
if (Test-Path $TarFilePath) {
    Write-Host "✅ File verification successful: $TarFilePath"
} else {
    Write-Host "❌ Error: File was not downloaded properly."
    exit
}

$ExtractPath = Join-Path $DependenciesDir "github-scripts"

# Ensure dependencies directory exists
if (!(Test-Path $DependenciesDir)) { 
    New-Item -ItemType Directory -Path $DependenciesDir -Force | Out-Null 
    Write-Host "✅ Created dependencies directory: $DependenciesDir"
}

# Verify the tar.gz file exists before extracting
if (!(Test-Path $TarFilePath)) {
    Write-Host "❌ Error: test.tar.gz not found at $TarFilePath. Please download it first."
    exit
}

# Ensure the extraction directory exists
if (!(Test-Path $ExtractPath)) {
    New-Item -ItemType Directory -Path $ExtractPath -Force | Out-Null
    Write-Host "✅ Created extraction directory: $ExtractPath"
}

# === Extract `test.tar.gz` ===
Write-Host "📂 Extracting test.tar.gz to $ExtractPath..."
try {
    tar -xf $TarFilePath -C $ExtractPath
    Write-Host "✅ Successfully extracted test.tar.gz"
} catch {
    Write-Host "❌ Extraction failed: $_"
    exit
}

# Verify extraction
if (Test-Path $ExtractPath) {
    Write-Host "✅ Extraction verification successful. Files are in: $ExtractPath"
} else {
    Write-Host "❌ Error: Extraction failed or directory does not exist."
    exit
}

# === Fix Broken pip Installation ===
function Fix-Pip {
    Log-Message "Checking if pip is working..."
    $PipCheck = & "$PythonExe" -m pip --version 2>&1

    if ($PipCheck -match "pip") {
        Log-Message "pip is installed: $PipCheck"
    } else {
        Log-Message "pip is broken or missing. Attempting full reinstall..."

        try {
            # Remove broken pip installation
            Remove-Item -Recurse -Force "$UserPythonFolder\Lib\site-packages\pip" -ErrorAction SilentlyContinue
            Remove-Item -Recurse -Force "$UserPythonFolder\Scripts\pip*" -ErrorAction SilentlyContinue

            # Download and reinstall pip
            Log-Message "Downloading get-pip.py..."
            Invoke-WebRequest -Uri "https://bootstrap.pypa.io/get-pip.py" -OutFile $PipInstallerPath -ErrorAction Stop

            # Reinstall pip properly
            Log-Message "Reinstalling pip..."
            & "$PythonExe" $PipInstallerPath --no-cache-dir --target="$DependenciesDir"

            # Verify pip installation
            $PipCheckNew = & "$PythonExe" -m pip --version 2>&1
            if ($PipCheckNew -match "pip") {
                Log-Message "pip reinstalled successfully: $PipCheckNew"
            } else {
                Log-Message "pip reinstall failed."
                exit
            }
        } catch { 
            Log-Message "Error reinstalling pip: $_"
            exit
        }
    }

    # Locate pip's vendor directory
    $PipVendorDir = & "$PythonExe" -c "import os, pip; print(os.path.join(os.path.dirname(pip.__file__), '_vendor'))" 2>&1
    if (!(Test-Path $PipVendorDir)) {
        Log-Message "Could not find pip vendor directory. Exiting."
        exit
    }
    Log-Message "pip vendor directory: $PipVendorDir"

    # Ensure `rich` is installed inside the correct pip vendor directory
    Log-Message "Installing `rich` inside pip's vendor directory..."
    $RichInstall = & "$PythonExe" -m pip install --no-cache-dir --target="$PipVendorDir" rich 2>&1

    if ($RichInstall -match "Successfully installed") {
        Log-Message "rich installed successfully inside pip vendor directory."
    } else {
        Log-Message "Failed to install `rich`: $RichInstall"
        exit
    }

# Ensure `_vendor` is in Python's `sys.path` before running setup.py
Log-Message "Adding `_vendor` to Python's `sys.path`..."
$VendorPathFix = & "$PythonExe" -c @'
import sys
import os
import pip
import subprocess

vendor_path = os.path.join(os.path.dirname(pip.__file__), "_vendor")
if vendor_path not in sys.path:
    sys.path.insert(0, vendor_path)
    print(f"Added vendor path to sys.path: {vendor_path}")
else:
    print(f"Vendor path already recognized: {vendor_path}")

# Check if `rich` exists
rich_path = os.path.join(vendor_path, "rich")
if os.path.exists(rich_path):
    print(f"rich module found at: {rich_path}")
else:
    print(f"rich module NOT FOUND at expected location: {rich_path}")
    print("Reinstalling `rich` inside vendor directory...")
    subprocess.run([sys.executable, "-m", "pip", "install", "--no-cache-dir", "--target", vendor_path, "rich"], check=True)

# Check if `rich.markdown` exists
markdown_path = os.path.join(rich_path, "markdown.py")
if os.path.exists(markdown_path):
    print(f"rich.markdown module found at: {markdown_path}")
else:
    print(f"rich.markdown module NOT FOUND. Manual intervention required.")
'@ 2>&1

Log-Message "$VendorPathFix"

    # Ensure pip is updated
    Log-Message "Updating pip, setuptools, and wheel..."
    $PipUpgrade = & "$PythonExe" -m pip install --upgrade --no-cache-dir --target="$DependenciesDir" pip setuptools wheel 2>&1

    if ($PipUpgrade -match "Successfully installed") {
        Log-Message "pip and dependencies updated successfully."
    } else {
        Log-Message "pip was already up to date or update failed: $PipUpgrade"
    }
}

# Call the Fix-Pip function before proceeding
Fix-Pip

# === Execute `setup.py` ===
Write-Host "🚀 Running setup.py..."
try {
    Start-Process -FilePath $PythonExe -ArgumentList "`"$SetupScriptPath`"" -NoNewWindow -Wait
    Write-Host "✅ setup.py executed successfully."
} catch {
    Write-Host "❌ Error executing setup.py: $_"
    exit
}
