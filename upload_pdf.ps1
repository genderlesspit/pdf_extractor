# Ensure System.Windows.Forms is loaded
if (-not ("System.Windows.Forms.OpenFileDialog" -as [Type])) { Add-Type -AssemblyName System.Windows.Forms }

# Define destination folder
$MasterPDFDir = "$env:USERPROFILE\PDF_Extractor\Master_PDFs"

# Ensure the folder exists
if (!(Test-Path $MasterPDFDir)) { New-Item -ItemType Directory -Path $MasterPDFDir -Force | Out-Null }

# Open file dialog to select a PDF
$FileDialog = New-Object System.Windows.Forms.OpenFileDialog
$FileDialog.Filter = "PDF Files (*.pdf)|*.pdf"
$FileDialog.Title = "Select a PDF to Upload"

if ($FileDialog.ShowDialog() -eq "OK") {
    $SelectedFile = $FileDialog.FileName
    $FileName = [System.IO.Path]::GetFileName($SelectedFile)
    $Destination = Join-Path $MasterPDFDir $FileName

    try {
        Copy-Item -Path $SelectedFile -Destination $Destination -Force
        Write-Host "✅ Successfully uploaded: $FileName → $Destination"
    } catch {
        Write-Host "❌ Error copying file: $_"
    }
} else {
    Write-Host "❌ No file selected."
}
