# === Define Required Assemblies ===
if (-not ("System.Windows.Forms.Form" -as [Type])) { Add-Type -AssemblyName System.Windows.Forms }
if (-not ("System.Drawing.Graphics" -as [Type])) { Add-Type -AssemblyName System.Drawing }

# === Define Directories ===
$ProgramDir = Join-Path ([System.Environment]::GetFolderPath("MyDocuments")) "PDF_Extractor"
$ScriptsDir = Join-Path $ProgramDir "dependencies\scripts"
$PythonExe = "$ProgramDir\dependencies\python.exe"

# === Function to Run Python Scripts ===
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

# === Font Style ===
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

# === Define Button Layout ===
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
    if ($PythonScript -ne "") {
        $Button.Add_Click({ Run-PythonScript -ScriptName $PythonScript })
    }
    return $Button
}

# === Add Buttons ===
$BtnUploadPDF = Create-RoundedButton "Upload PDF" $StartX $StartY "upload_pdf.py"
$Form.Controls.Add($BtnUploadPDF)

$BtnMassMasterPDFs = Create-RoundedButton "Mass Input Master PDFs" ($StartX + ($ButtonWidth + $Spacing) * 1) $StartY "mass_input.py master_pdf"
$Form.Controls.Add($BtnMassMasterPDFs)

$BtnMassInput = Create-RoundedButton "Mass Input Output Name" ($StartX + ($ButtonWidth + $Spacing) * 2) $StartY "mass_input.py output_name"
$Form.Controls.Add($BtnMassInput)

$BtnPageRange = Create-RoundedButton "Mass Input Page Range" ($StartX + ($ButtonWidth + $Spacing) * 3) $StartY "mass_input.py page_range"
$Form.Controls.Add($BtnPageRange)

$BtnStart = Create-RoundedButton "Start Process" ($StartX + ($ButtonWidth + $Spacing) * 4) $StartY "validate_inputs.py"
$BtnStart.Add_Click({
    Run-PythonScript -ScriptName "validate_inputs.py"
    Start-Sleep -Seconds 1

    if (Read-ErrorsFromJSON) {
        Run-PythonScript -ScriptName "process_extraction.py"
        Start-Sleep -Seconds 1
        Run-PythonScript -ScriptName "export.py"
    }
})
$Form.Controls.Add($BtnStart)

$BtnSaveLoadGrid = Create-RoundedButton "Save/Load DataGrid" ($StartX + ($ButtonWidth + $Spacing) * 5) $StartY ""
$BtnSaveLoadGrid.Add_Click({
    $Choice = [System.Windows.Forms.MessageBox]::Show("Do you want to save or load the grid?", "Save or Load", "YesNoCancel", "Question")
    if ($Choice -eq "Yes") {
        Run-PythonScript -ScriptName "save_load_grid.py"
    } elseif ($Choice -eq "No") {
        Run-PythonScript -ScriptName "save_load_grid.py load"
    }
})
$Form.Controls.Add($BtnSaveLoadGrid)

$BtnClearGrid = Create-RoundedButton "Clear Grid" ($StartX + ($ButtonWidth + $Spacing) * 6) $StartY "clear_grid.py"
$Form.Controls.Add($BtnClearGrid)

# === Define DataTable for GridView ===
$DataTable = New-Object System.Data.DataTable
$DataTable.Columns.Add("Master PDF", [string]) | Out-Null
$DataTable.Columns.Add("Output Name") | Out-Null
$DataTable.Columns.Add("Page Range") | Out-Null
$DataGridView.DataSource = $DataTable

# === Function to Read Error Logs ===
function Read-ErrorsFromJSON {
    $ErrorFile = "$ProgramDir\\errors.json"

    if (!(Test-Path $ErrorFile)) {
        return $null
    }

    try {
        $Errors = Get-Content -Raw -Path $ErrorFile | ConvertFrom-Json
        if ($Errors.missing_pdfs.Count -gt 0 -or $Errors.invalid_page_ranges.Count -gt 0 -or $Errors.empty_fields.Count -gt 0 -or $Errors.unreadable_pdfs.Count -gt 0) {
            $ErrorMessage = "🚨 Validation Errors:`n`n"
            if ($Errors.missing_pdfs.Count -gt 0) { $ErrorMessage += "❌ Missing PDFs:`n" + ($Errors.missing_pdfs -join "`n") + "`n`n" }
            if ($Errors.invalid_page_ranges.Count -gt 0) { $ErrorMessage += "⚠ Invalid Page Ranges:`n" + ($Errors.invalid_page_ranges -join "`n") + "`n`n" }
            if ($Errors.empty_fields.Count -gt 0) { $ErrorMessage += "❌ Missing Fields:`n" + ($Errors.empty_fields -join "`n") + "`n`n" }

            [System.Windows.Forms.MessageBox]::Show($ErrorMessage, "Validation Errors", "OK", "Error")
            return $false
        }
    } catch {
        Write-Host "❌ Error reading validation logs."
    }

    return $true
}

# === Display GUI ===
$Form.ShowDialog()
