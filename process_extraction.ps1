# Define paths
$BaseDir = "$env:USERPROFILE\PDF_Extractor"
$MasterPDFDir = "$BaseDir\Master_PDFs"
$ExportsDir = "$BaseDir\Exports"
$GridUpdateFile = "$BaseDir\temp_grid_update.csv"
$FailureLogFile = "$BaseDir\failed_extractions.json"

# Ensure required directories exist
if (!(Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

# Maximum retries for failed extractions
$MAX_RETRIES = 3

# Function to extract specified pages from a PDF (Placeholder - Replace with actual PDF library)
function Extract-PDFPages {
    param (
        [string]$PDFName,
        [string]$PageRanges,
        [string]$OutputName
    )

    $PDFPath = Join-Path $MasterPDFDir $PDFName
    $OutputPath = Join-Path $ExportsDir "$OutputName.pdf"

    if (!(Test-Path $PDFPath)) {
        Write-Host "❌ File not found: $PDFPath"
        return $false
    }

    try {
        Write-Host "📄 Extracting pages $PageRanges from $PDFName → $OutputName.pdf"

        # Replace this with actual PDF extraction logic
        # Example: Using a PowerShell module like PDFtk or Pdfium
        # PdfiumExtract -Input $PDFPath -Output $OutputPath -Pages $PageRanges

        # Simulated success
        Start-Sleep -Seconds 2  # Simulate processing time
        Write-Host "✅ Extraction complete for $OutputName.pdf"

        return $true
    } catch {
        Write-Host "❌ Error extracting pages from $PDFName: $_"
        return $false
    }
}

# Function to retry failed extractions
function Retry-FailedExtractions {
    if (!(Test-Path $FailureLogFile)) { return }

    try {
        $FailedExtractions = Get-Content -Raw -Path $FailureLogFile | ConvertFrom-Json
    } catch {
        Write-Host "⚠ Error reading `${FailureLogFile}`. Skipping retries."
        return
    }

    if ($FailedExtractions.Count -eq 0) { return }

    Write-Host "Retrying failed extractions..."

    $NewFailedExtractions = @{}

    foreach ($Key in $FailedExtractions.Keys) {
        $Attempts = $FailedExtractions[$Key]
        $Parts = $Key -split "\|"
        $PDFName = $Parts[0]
        $OutputName = $Parts[1]
        $PageRanges = $Parts[2]

        Write-Host "Retrying extraction ($Attempts/$MAX_RETRIES) for: `${OutputName}`"

        $Success = Extract-PDFPages -PDFName $PDFName -PageRanges $PageRanges -OutputName $OutputName

        if (!$Success -and $Attempts -lt $MAX_RETRIES) {
            $NewFailedExtractions[$Key] = $Attempts + 1
        } elseif ($Success) {
            Write-Host "✅ Successfully retried: `${OutputName}`"
        } else {
            Write-Host "❌ Max retries reached for: `${OutputName}`"
        }
    }

    # Update failure log
    $NewFailedExtractions | ConvertTo-Json -Depth 2 | Out-File -Encoding UTF8 $FailureLogFile

    if ($NewFailedExtractions.Count -eq 0) {
        Remove-Item $FailureLogFile -Force
    }
}

# Function to process extractions
function Process-Extractions {
    if (!(Test-Path $GridUpdateFile)) {
        Write-Host "❌ No input data found."
        return
    }

    $GridData = Import-Csv -Path $GridUpdateFile
    if ($GridData.Count -eq 0) {
        Write-Host "❌ No valid rows found."
        return
    }

    $FailedExtractions = @{}

    foreach ($Row in $GridData) {
        try {
            $PDFName = $Row."Master PDF"
            $OutputName = $Row."Output Name"
            $PageRanges = $Row."Page Range"

            $Success = Extract-PDFPages -PDFName $PDFName -PageRanges $PageRanges -OutputName $OutputName

            if (!$Success) {
                $FailedExtractions["$PDFName|$OutputName|$PageRanges"] = 0
            }
        } catch {
            Write-Host "❌ Skipping invalid row."
        }
    }

    # Save failures for retry
    $FailedExtractions | ConvertTo-Json -Depth 2 | Out-File -Encoding UTF8 $FailureLogFile

    # Retry failed extractions
    Retry-FailedExtractions
}

# Run the extraction process
Process-Extractions
