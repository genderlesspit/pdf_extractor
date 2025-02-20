# === Define Required Assemblies ===
if (-not ("System.Windows.Forms.Form" -as [Type])) { Add-Type -AssemblyName System.Windows.Forms }
if (-not ("System.Drawing.Graphics" -as [Type])) { Add-Type -AssemblyName System.Drawing }

# === Define Paths (Matching setup.ps1) ===
$ProgramDir = "$env:USERPROFILE\PDF_Extractor"
$ScriptsDir = "$ProgramDir"
$DependenciesDir = "$ProgramDir\dependencies"
$CSVFile = "$ProgramDir\temp_grid_update.csv"

# === Function to Run PowerShell Scripts ===
function Run-PowerShellScript {
    param ([string]$ScriptName)
    $ScriptPath = Join-Path $ScriptsDir $ScriptName

    if (!(Test-Path $ScriptPath)) {
        [System.Windows.Forms.MessageBox]::Show("PowerShell script not found: $ScriptPath", "Error", "OK", "Error")
        return
    }

    Write-Host "🚀 Running $ScriptName..."
    Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$ScriptPath`"" -NoNewWindow -Wait
}

# === GUI Form Setup ===
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "PDF Extractor"
$Form.Size = New-Object System.Drawing.Size(1300, 700)
$Form.StartPosition = "CenterScreen"
$Form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$Form.FormBorderStyle = "FixedDialog"

# === Font Style ===
$Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)

# === DataGridView Setup (Restored Grid Box) ===
$DataGridView = New-Object System.Windows.Forms.DataGridView
$DataGridView.Location = New-Object System.Drawing.Point(20, 20)
$DataGridView.Size = New-Object System.Drawing.Size(1160, 450)
$DataGridView.AutoSizeColumnsMode = "Fill"
$DataGridView.AllowUserToAddRows = $true
$DataGridView.BorderStyle = "None"
$DataGridView.BackgroundColor = [System.Drawing.Color]::WhiteSmoke
$DataGridView.Font = $Font
$Form.Controls.Add($DataGridView)

# === Define DataTable for GridView ===
$DataTable = New-Object System.Data.DataTable
$DataTable.Columns.Add("Master PDF", [string]) | Out-Null
$DataTable.Columns.Add("Output Name", [string]) | Out-Null
$DataTable.Columns.Add("Page Range", [string]) | Out-Null
$DataGridView.DataSource = $DataTable

# === Load CSV Data into Grid ===
function Load-GridData {
    if (!(Test-Path $CSVFile)) {
        Write-Host "⚠ No saved grid data found."
        return
    }

    try {
        $GridData = Import-Csv -Path $CSVFile
        $DataTable.Rows.Clear()
        foreach ($Row in $GridData) {
            $DataTable.Rows.Add($Row."Master PDF", $Row."Output Name", $Row."Page Range") | Out-Null
        }
    } catch {
        Write-Host "❌ Error loading grid data: $_"
    }
}

# === Save Grid Data to CSV ===
function Save-GridData {
    try {
        $DataTable | Export-Csv -Path $CSVFile -NoTypeInformation -Encoding UTF8
        Write-Host "✅ Grid data saved successfully."
    } catch {
        Write-Host "❌ Error saving grid data: $_"
    }
}

# === Define Button Layout ===
$ButtonWidth = 170
$ButtonHeight = 40
$Spacing = 10
$TotalButtons = 7
$TotalWidth = ($ButtonWidth * $TotalButtons) + ($Spacing * ($TotalButtons - 1))
$StartX = [Math]::Max(20, ($Form.ClientSize.Width - $TotalWidth) / 2)
$StartY = 500

# === Function to Create Styled Buttons ===
function Create-RoundedButton {
    param ([string]$Text, [int]$X, [int]$Y, [string]$ScriptName)
    $Button = New-Object System.Windows.Forms.Button
    $Button.Text = $Text
    $Button.Size = New-Object System.Drawing.Size($ButtonWidth, $ButtonHeight)
    $Button.Location = New-Object System.Drawing.Point($X, $Y)
    $Button.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $Button.FlatStyle = "Flat"
    $Button.FlatAppearance.BorderSize = 1
    $Button.Font = $Font
    if ($ScriptName -ne "") {
        $Button.Add_Click({ Run-PowerShellScript -ScriptName $ScriptName })
    }
    return $Button
}

# === Add Buttons ===
$BtnUploadPDF = Create-RoundedButton "Upload PDF" $StartX $StartY "upload-uDF.ps1"
$Form.Controls.Add($BtnUploadPDF)

$BtnMassMasterPDFs = Create-RoundedButton "Mass Input Master PDFs" ($StartX + ($ButtonWidth + $Spacing) * 1) $StartY "mass-Input.ps1 master_pdf"
$Form.Controls.Add($BtnMassMasterPDFs)

$BtnMassInput = Create-RoundedButton "Mass Input Output Name" ($StartX + ($ButtonWidth + $Spacing) * 2) $StartY "mass-Input.ps1 output_name"
$Form.Controls.Add($BtnMassInput)

$BtnPageRange = Create-RoundedButton "Mass Input Page Range" ($StartX + ($ButtonWidth + $Spacing) * 3) $StartY "mass-Input.ps1 page_range"
$Form.Controls.Add($BtnPageRange)

$BtnStart = Create-RoundedButton "Start Process" ($StartX + ($ButtonWidth + $Spacing) * 4) $StartY "validate-Inputs.ps1"
$BtnStart.Add_Click({
    Run-PowerShellScript -ScriptName "validate-inputs.ps1"
    Start-Sleep -Seconds 1

    if (Read-ErrorsFromJSON) {
        Run-PowerShellScript -ScriptName "process-extraction.ps1"
        Start-Sleep -Seconds 1
        Run-PowerShellScript -ScriptName "export-files.ps1"
    }
})
$Form.Controls.Add($BtnStart)

$BtnSaveLoadGrid = Create-RoundedButton "Save/Load DataGrid" ($StartX + ($ButtonWidth + $Spacing) * 5) $StartY ""
$BtnSaveLoadGrid.Add_Click({
    $Choice = [System.Windows.Forms.MessageBox]::Show("Do you want to save or load the grid?", "Save or Load", "YesNoCancel", "Question")
    if ($Choice -eq "Yes") {
        Save-GridData
    } elseif ($Choice -eq "No") {
        Load-GridData
    }
})
$Form.Controls.Add($BtnSaveLoadGrid)

$BtnClearGrid = Create-RoundedButton "Clear Grid" ($StartX + ($ButtonWidth + $Spacing) * 6) $StartY ""
$BtnClearGrid.Add_Click({
    $DataTable.Rows.Clear()
    Write-Host "✅ Grid cleared."
})
$Form.Controls.Add($BtnClearGrid)

# === Load Grid Data on Startup ===
Load-GridData

# === Display GUI ===
$Form.ShowDialog()
