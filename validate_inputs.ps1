# Define paths
$BaseDir = "$env:USERPROFILE\PDF_Extractor"
$MasterPDFDir = "$BaseDir\Master_PDFs"
$GridUpdateFile = "$BaseDir\temp_grid_update.csv"
$ErrorLogFile = "$BaseDir\errors.json"

# Ensure directories exist
if (!(Test-Path $BaseDir)) { New-Item -ItemType Directory -Path $BaseDir -Force | Out-Null }

Write-Host "🔍 Starting validation process..."

# Function to get total pages in a PDF (Dummy Implementation)
function Get-PDFPageCount {
    param ([string]$PDFPath)

    try {
        Write-Host "📄 Checking page count for: ${PDFPath}"
        return 10  # Simulated fixed page count for now
    } catch {
        Write-Host "❌ Error reading PDF: $_"
        return $null
    }
}

# Function to validate page range format
function Validate-PageRange {
    param ([string]$PageRange, [int]$TotalPages)
    $InvalidPages = @()

    if (-not $PageRange -or -not $TotalPages -or $TotalPages -le 0) {
        Write-Host "❌ Invalid page range input: '$PageRange' (Total: $TotalPages)"
        return @("Invalid Input")
    }

    Write-Host "🔍 Validating Page Range: '$PageRange' (Total Pages: $TotalPages)"

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
    Write-Host "📂 Checking grid file: ${GridUpdateFile}"

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

    try {
        Write-Host "🔄 Loading CSV..."
        $GridData = Import-Csv -Path $GridUpdateFile -ErrorAction Stop
        Write-Host "✅ CSV loaded successfully. Rows found: $($GridData.Count)"
    } catch {
        Write-Host "❌ Error reading CSV file: $_"
        return $false
    }

    if ($GridData.Count -eq 0) {
        Write-Host "⚠ Grid file is empty!"
        return $false
    }

    foreach ($Row in $GridData) {
        try {
            $PDFName = $Row."Master PDF"
            $OutputName = $Row."Output Name"
            $PageRange = $Row."Page Range"
            $PDFPath = Join-Path $MasterPDFDir $PDFName

            Write-Host "🔎 Validating Row -> PDF: '${PDFName}', Output: '${OutputName}', Page Range: '${PageRange}'"

            # Check for empty required fields
            if (-not $PDFName -or -not $OutputName -or -not $PageRange) {
                Write-Host "⚠ Missing required fields in row!"
                $Errors["empty_fields"] += "Row: Missing required fields."
                continue
            }

            # Check if the PDF file exists
            if ($PDFName -and !(Test-Path $PDFPath)) {
                Write-Host "⚠ PDF not found: ${PDFName}"
                $Errors["missing_pdfs"] += "Row: ${PDFName} not found."
                continue
            }

            # Get total pages in the PDF
            $TotalPages = Get-PDFPageCount -PDFPath $PDFPath
            if ($null -eq $TotalPages) {
                Write-Host "⚠ Could not read PDF: ${PDFName}"
                $Errors["unreadable_pdfs"] += "Row: Could not read ${PDFName}."
                continue
            }

            # Validate page range
            $InvalidPages = Validate-PageRange -PageRange $PageRange -TotalPages $TotalPages
            if ($InvalidPages.Count -gt 0) {
                Write-Host "⚠ Invalid pages detected in `${PDFName}`: $($InvalidPages -join ', ')"
                $Errors["invalid_page_ranges"] += "Row: Invalid pages [$($InvalidPages -join ', ')] for ${PDFName} (Total: ${TotalPages} pages)."
            }

        } catch {
            Write-Host "❌ Unexpected error processing row: $_"
        }
    }

    # Write errors to JSON
    Write-Host "💾 Saving validation errors..."
    try {
        $Errors | ConvertTo-Json -Depth 2 | Out-File -Encoding UTF8 $ErrorLogFile
        Write-Host "✅ Errors saved successfully to errors.json"
    } catch {
        Write-Host "❌ Failed to write errors to file: $_"
        return $false
    }

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

# Run the validation and properly exit with a status code
Write-Host "🔄 Running validation..."
$Valid = $false

try {
    $Valid = Validate-Inputs
} catch {
    Write-Host "❌ Script crashed unexpectedly: $_"
    exit 1
}

if ($Valid) {
    Write-Host "🚀 Validation completed successfully!"
    exit 0
} else {
    Write-Host "❌ Validation encountered errors. Check errors.json."
    exit 1
}
