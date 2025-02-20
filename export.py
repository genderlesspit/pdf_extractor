# Ensure System.Windows.Forms is loaded
if (-not ("System.Windows.Forms.FolderBrowserDialog" -as [Type])) { Add-Type -AssemblyName System.Windows.Forms }

# Define paths
$BaseDir = "$env:USERPROFILE\PDF_Extractor"
$ExportsDir = "$BaseDir\Exports"

# Function to prompt user for export folder
function Get-ExportFolder {
    $FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderDialog.Description = "Select a folder to export extracted PDFs"

    if ($FolderDialog.ShowDialog() -eq "OK") {
        return $FolderDialog.SelectedPath
    } else {
        Write-Host "❌ No folder selected. Export canceled."
        exit 1
    }
}

# Function to move extracted PDFs
function Export-Files {
    if (!(Test-Path $ExportsDir)) {
        Write-Host "❌ Exports directory does not exist: $ExportsDir"
        exit 1
    }

    $ExportFolder = Get-ExportFolder
    if (-not $ExportFolder) {
        exit 1
    }

    Write-Host "📂 Moving extracted PDFs to: $ExportFolder"

    $PDFs = Get-ChildItem -Path $ExportsDir -Filter "*.pdf"
    if ($PDFs.Count -eq 0) {
        Write-Host "⚠ No PDFs found in Exports directory."
        exit 1
    }

    foreach ($PDF in $PDFs) {
        $Destination = Join-Path $ExportFolder $PDF.Name
        try {
            Move-Item -Path $PDF.FullName -Destination $Destination -Force
            Write-Host "✅ Moved: $($PDF.Name) → $ExportFolder"
        } catch {
            Write-Host "❌ Error moving $($PDF.Name): $($_.Exception.Message)"
        }
    }

    Write-Host "🚀 Export process completed successfully!"
}

# Run export function
Export-Files
