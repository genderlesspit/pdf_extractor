# Define script parameters
param ([string]$Type)

# Ensure System.Windows.Forms is loaded
if (-not ("System.Windows.Forms.Form" -as [Type])) { Add-Type -AssemblyName System.Windows.Forms }

# Define paths
$BaseDir = "$env:USERPROFILE\PDF_Extractor"
$GridUpdateFile = "$BaseDir\temp_grid_update.csv"

# Ensure base directory exists
if (!(Test-Path $BaseDir)) { New-Item -ItemType Directory -Path $BaseDir -Force | Out-Null }

# Function to prompt user for mass input
function Get-MassInput {
    param ([string]$InputType)

    # Define prompt messages
    $PromptMessages = @{
        "master_pdf" = "Enter Master PDFs (one per line):"
        "output_name" = "Enter Output Names (one per line):"
        "page_range" = "Enter Page Ranges (one per line):"
    }

    # Create input form
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Mass Input - $InputType"
    $Form.Size = New-Object System.Drawing.Size(400, 300)
    $Form.StartPosition = "CenterScreen"

    $TextBox = New-Object System.Windows.Forms.TextBox
    $TextBox.Multiline = $true
    $TextBox.ScrollBars = "Vertical"
    $TextBox.Size = New-Object System.Drawing.Size(360, 200)
    $TextBox.Location = New-Object System.Drawing.Point(10, 10)
    $Form.Controls.Add($TextBox)

    $BtnOK = New-Object System.Windows.Forms.Button
    $BtnOK.Text = "OK"
    $BtnOK.Location = New-Object System.Drawing.Point(100, 250)
    $BtnOK.Size = New-Object System.Drawing.Size(80, 30)
    $Form.Controls.Add($BtnOK)

    $BtnCancel = New-Object System.Windows.Forms.Button
    $BtnCancel.Text = "Cancel"
    $BtnCancel.Location = New-Object System.Drawing.Point(200, 250)
    $BtnCancel.Size = New-Object System.Drawing.Size(80, 30)
    $Form.Controls.Add($BtnCancel)

    # OK Button Click Event
    $BtnOK.Add_Click({
        $Lines = $TextBox.Text -split "`r`n"  # Split input into lines

        if ($Lines.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No input provided!", "Error", "OK", "Error")
            return
        }

        Update-Grid -InputType $InputType -Values $Lines
        $Form.Close()
    })

    # Cancel Button Click Event
    $BtnCancel.Add_Click({ $Form.Close() })

    $Form.ShowDialog()
}

# Function to update the grid CSV
function Update-Grid {
    param (
        [string]$InputType,
        [string[]]$Values
    )

    # Load existing data if available
    $ExistingData = @()
    if (Test-Path $GridUpdateFile) {
        $ExistingData = Import-Csv -Path $GridUpdateFile
    }

    # Determine column index
    $ColumnMap = @{
        "master_pdf"  = "Master PDF"
        "output_name" = "Output Name"
        "page_range"  = "Page Range"
    }
    $ColumnName = $ColumnMap[$InputType]

    if (-not $ColumnName) {
        Write-Host "❌ Invalid input type specified."
        return
    }

    # Update data
    for ($i = 0; $i -lt $Values.Count; $i++) {
        if ($i -lt $ExistingData.Count) {
            $ExistingData[$i].$ColumnName = $Values[$i]
        } else {
            $NewRow = [PSCustomObject]@{
                "Master PDF"  = ""
                "Output Name" = ""
                "Page Range"  = ""
            }
            $NewRow.$ColumnName = $Values[$i]
            $ExistingData += $NewRow
        }
    }

    # Save updated data
    $ExistingData | Export-Csv -Path $GridUpdateFile -NoTypeInformation -Encoding UTF8
    Write-Host "✅ Successfully updated $($Values.Count) entries for $InputType."
}

# Main execution: Call the input function based on user selection
if ($Type) {
    Get-MassInput -InputType $Type
} else {
    Write-Host "❌ No input type specified. Use 'master_pdf', 'output_name', or 'page_range'."
}
