# Define paths
$BaseDir = "$env:USERPROFILE\PDF_Extractor"
$MasterPDFDir = "$BaseDir\Master_PDFs"
$GridUpdateFile = "$BaseDir\temp_grid_update.csv"
$ErrorLogFile = "$BaseDir\errors.json"

# Ensure required directories exist
if (!(Test-Path $BaseDir)) { New-Item -ItemType Directory -Path $BaseDir -Force | Out-Null }

# Function to get total pages in a PDF
function Get-PDFPageCount {
    param ([string]$PDFPath)

    try {
        $PDFInfo = & "pdftotext.exe" -l 1 -q $PDFPath - 2>&1  # Dummy command for simulation
        if ($PDFInfo -match "Syntax Error") {
            return $null  # Simulating an error for unreadable PDFs
        }
        return 10  # Replace with actual page count logic when integrated with a PDF tool
    } catch {
        return $null
    }
}

# Function to validate page range format
function Validate-PageRange {
    param ([string]$PageRange, [int]$TotalPages)
    $InvalidPages = @()

    # Split by commas
    $Ranges = $PageRange -split ","

    foreach ($Range in $Ranges) {
        if ($Range -match "^\d+-\d+$") {
            $Start, $End = $Range -split "-"
            if ([int]$Start -lt 1 -or [int]$End -gt $TotalPages -or [int]$Start -gt [int]$End) {
                $InvalidPages += $Range
            }
        } elseif ($Range -match "^\d+$") {
            if ([int]$Range -lt 1 -or [int]$Range -gt $TotalPages) {
                $InvalidPages += $Range
            }
        } else {
            $InvalidPages += $Range
        }
    }

    return $InvalidPages
}

# Function to validate grid inputs
function Validate-Inputs {
    $Errors = @{
        "missing_pdfs"      = @()
        "invalid_page_ranges" = @()
        "empty_fields"       = @()
        "unreadable_pdfs"    = @()
    }

    if (!(Test-Path $GridUpdateFile)) {
        Write-Host "❌ No input data found."
        return $false
    }

    $GridData = Import-Csv -Path $GridUpdateFile

    foreach ($Row in $GridData) {
        $PDFName = $Row."Master PDF"
        $OutputName = $Row."Output Name"
        $PageRange = $Row."Page Range"
        $PDFPath = Join-Path $MasterPDFDir $PDFName

        # Check for empty required fields
        if (-not $PDFName -or -not $OutputName -or -not $PageRange) {
            $Errors["empty_fields"] += "Row: Missing required fields."
        }

        # Check if the PDF file exists
        if ($PDFName -and !(Test-Path $PDFPath)) {
            $Errors["missing_pdfs"] += "Row: $PDFName not found."
            continue
        }

        # Get total pages in the PDF
        $TotalPages = Get-PDFPageCount -PDFPath $PDFPath
        if ($null -eq $TotalPages) {
            $Errors["unreadable_pdfs"] += "Row: Could not read $PDFName."
            continue
        }

        # Validate page range
        $InvalidPages = Validate-PageRange -PageRange $PageRange -TotalPages $TotalPages
        if ($InvalidPages.Count -gt 0) {
            $Errors["invalid_page_ranges"] += "Row: Invalid pages [$($InvalidPages -join ', ')] for $PDFName (Total: $TotalPages pages)."
        }
    }

    # Write errors to JSON
    $Errors | ConvertTo-Json -Depth 2 | Out-File -Encoding UTF8 $ErrorLogFile

    if ($Errors["missing_pdfs"].Count -gt 0 -or
        $Errors["invalid_page_ranges"].Count -gt 0 -or
        $Errors["empty_fields"].Count -gt 0 -or
        $Errors["unreadable_pdfs"].Count -gt 0) {
        Write-Host "❌ Validation failed. Errors logged to errors.json."
        return $false
    }

    Write-Host "✅ All inputs are valid."
    return $true
}

# Run the validation
$Valid = Validate-Inputs
exit (if ($Valid) { 0 } else { 1 })
