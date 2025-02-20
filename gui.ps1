# === Define Required Assemblies ===
if (-not ("System.Windows.Forms.Form" -as [Type])) { Add-Type -AssemblyName System.Windows.Forms }
if (-not ("System.Drawing.Graphics" -as [Type])) { Add-Type -AssemblyName System.Drawing }

# === PowerShell Version Check ===
if ($PSVersionTable.PSVersion.Major -ge 7) {
    Write-Host "⚠ WARNING: This script is designed for Windows PowerShell 5.1. Some features may not work in PowerShell Core 7.x or later." -ForegroundColor Yellow
    [System.Windows.Forms.MessageBox]::Show("This script is designed for Windows PowerShell 5.1. Some features may not work in PowerShell Core 7.x or later.", "Compatibility Warning", "OK", "Warning")
    exit
}

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
$DependenciesDir = Join-Path $ProgramDir "dependencies"
$ScriptsDir = Join-Path $DependenciesDir "scripts"
$PythonExe = "$UserPythonFolder\\python.exe"

# === Ensure Python Exists ===
if (!(Test-Path $PythonExe)) {
    Write-Host "❌ Python not found in dependencies. Please install Python first."
    [System.Windows.Forms.MessageBox]::Show("Python not found in dependencies. Please install Python first.", "Error", "OK", "Error")
    exit
}

# === Function to Execute Python Scripts ===
function Run-PythonScript {
    param ([string]$ScriptName)
    $ScriptPath = Join-Path $ScriptsDir $ScriptName

    if (!(Test-Path $ScriptPath)) {
        Write-Host "❌ Python script not found: $ScriptPath"
        [System.Windows.Forms.MessageBox]::Show("Python script not found: $ScriptPath", "Error", "OK", "Error")
        return
    }

    Write-Host "🚀 Running $ScriptName..."
    Start-Process -FilePath $PythonExe -ArgumentList "`"$ScriptPath`"" -NoNewWindow -Wait
}

# === GUI Form Setup ===
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "PDF Extractor"
$Form.Size = New-Object System.Drawing.Size(1300, 650)
$Form.StartPosition = "CenterScreen"
$Form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$Form.FormBorderStyle = "FixedDialog"

# === Define Font ===
$Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)

# === DataGridView Setup ===
$DataGridView = New-Object System.Windows.Forms.DataGridView
$DataGridView.Location = New-Object System.Drawing.Point(20, 20)
$DataGridView.Size = New-Object System.Drawing.Size(1160, 450)
$DataGridView.AutoSizeColumnsMode = "Fill"
$DataGridView.AllowUserToAddRows = $true
$DataGridView.BorderStyle = "None"
$DataGridView.BackgroundColor = [System.Drawing.Color]::WhiteSmoke
$DataGridView.Font = $Font
$Form.Controls.Add($DataGridView)

# === Adjust Button Layout ===
$ButtonWidth = 170
$ButtonHeight = 40
$Spacing = 10
$TotalButtons = 7
$TotalWidth = ($ButtonWidth * $TotalButtons) + ($Spacing * ($TotalButtons - 1))
$StartX = [Math]::Max(20, ($Form.ClientSize.Width - $TotalWidth) / 2)
$StartY = 520

# === Function to Create Styled Buttons ===
function Create-RoundedButton {
    param ([string]$Text, [int]$X, [int]$Y, [string]$PythonScript)
    $Button = New-Object System.Windows.Forms.Button
    $Button.Text = $Text
    $Button.Size = New-Object System.Drawing.Size($ButtonWidth, $ButtonHeight)
    $Button.Location = New-Object System.Drawing.Point($X, $Y)
    $Button.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $Button.FlatStyle = "Flat"
    $Button.FlatAppearance.BorderSize = 1
    $Button.Font = $Font
    $Button.Add_Click({ Run-PythonScript -ScriptName $PythonScript })
    return $Button
}

# === Create Buttons ===
$BtnUploadPDF = Create-RoundedButton "Upload PDF" $StartX $StartY "upload_pdf.py"
$Form.Controls.Add($BtnUploadPDF)

$BtnMassMasterPDFs = Create-RoundedButton "Mass Input Master PDFs" ($StartX + ($ButtonWidth + $Spacing) * 1) $StartY "mass_input_master_pdfs.py"
$Form.Controls.Add($BtnMassMasterPDFs)

$BtnMassInput = Create-RoundedButton "Mass Input Output Name" ($StartX + ($ButtonWidth + $Spacing) * 2) $StartY "mass_input_output.py"
$Form.Controls.Add($BtnMassInput)

$BtnPageRange = Create-RoundedButton "Mass Input Page Range" ($StartX + ($ButtonWidth + $Spacing) * 3) $StartY "mass_input_page_range.py"
$Form.Controls.Add($BtnPageRange)

$BtnStart = Create-RoundedButton "Start" ($StartX + ($ButtonWidth + $Spacing) * 4) $StartY "start_process.py"
$Form.Controls.Add($BtnStart)

$BtnRetry = Create-RoundedButton "Retry Failed" ($StartX + ($ButtonWidth + $Spacing) * 5) $StartY "retry_failed.py"
$Form.Controls.Add($BtnRetry)

$BtnClearGrid = Create-RoundedButton "Clear Grid" ($StartX + ($ButtonWidth + $Spacing) * 6) $StartY "clear_grid.py"
$Form.Controls.Add($BtnClearGrid)

# === Create DataTable ===
$DataTable = New-Object System.Data.DataTable
$DataTable.Columns.Add("Master PDF", [string]) | Out-Null
$DataTable.Columns.Add("Output Name") | Out-Null
$DataTable.Columns.Add("Page Range") | Out-Null
$DataGridView.DataSource = $DataTable

# === Enable Auto-Complete for PDFs ===
function Enable-AutoComplete {
    $MasterPDFDir = Join-Path $ProgramDir "Master_PDFs"

    if (!(Test-Path $MasterPDFDir)) {
        Write-Host "⚠ Warning: Master PDF directory not found: $MasterPDFDir"
        return
    }

    $MasterPDFs = Get-ChildItem -Path $MasterPDFDir -Filter "*.pdf" | Select-Object -ExpandProperty Name
    if ($MasterPDFs.Count -eq 0) {
        Write-Host "⚠ No PDFs found in $MasterPDFDir. Auto-complete will be empty."
        return
    }

    foreach ($Row in $DataGridView.Rows) {
        if ($Row.Cells["Master PDF"] -is [System.Windows.Forms.DataGridViewTextBoxCell]) {
            $AutoComplete = New-Object System.Windows.Forms.AutoCompleteStringCollection
            $AutoComplete.AddRange($MasterPDFs)

            $TextBox = New-Object System.Windows.Forms.TextBox
            $TextBox.AutoCompleteMode = "SuggestAppend"
            $TextBox.AutoCompleteSource = "CustomSource"
            $TextBox.AutoCompleteCustomSource = $AutoComplete

            $Row.Cells["Master PDF"].Tag = $TextBox
        }
    }

    Write-Host "✅ Auto-complete enabled with $($MasterPDFs.Count) PDFs."
}

function Update-GridFromCSV {
    $CSVPath = "$ProgramDir\\temp_grid_update.csv"

    if (!(Test-Path $CSVPath)) {
        Write-Host "⚠ No new grid data found."
        return
    }

    try {
        $CSVData = Import-Csv -Path $CSVPath -Encoding UTF8

        # Clear existing grid
        $DataTable.Clear()

        foreach ($Row in $CSVData) {
            $NewRow = $DataTable.NewRow()
            $NewRow["Master PDF"] = $Row."Column1"
            $NewRow["Output Name"] = $Row."Column2"
            $NewRow["Page Range"] = $Row."Column3"
            $DataTable.Rows.Add($NewRow)
        }

        # Clear the temp CSV file
        Remove-Item $CSVPath -Force

        Write-Host "✅ Grid updated with new entries."
    } catch {
        Write-Host "❌ Error updating grid: $_"
    }
}

# Modify the buttons to call the unified script with arguments
$BtnMassMasterPDFs.Add_Click({
    Run-PythonScript -ScriptName "mass_input.py master_pdf"
    Start-Sleep -Seconds 1
    Update-GridFromCSV
})

$BtnMassInput.Add_Click({
    Run-PythonScript -ScriptName "mass_input.py output_name"
    Start-Sleep -Seconds 1
    Update-GridFromCSV
})

$BtnPageRange.Add_Click({
    Run-PythonScript -ScriptName "mass_input.py page_range"
    Start-Sleep -Seconds 1
    Update-GridFromCSV
})

# === Show GUI ===
$Form.ShowDialog()
