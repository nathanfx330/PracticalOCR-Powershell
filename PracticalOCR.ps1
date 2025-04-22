<#
.SYNOPSIS
Processes a selected PDF file by extracting pages as images, performing OCR,
and merging the results into a final searchable PDF. Uses ImageMagick 7 syntax.

.DESCRIPTION
Checks for required tools (magick, pdfinfo, tesseract, pdftk) using user-configured paths.
Prompts the user to select a PDF from the './pdfs' directory.
Optionally applies image level adjustments during conversion.
Converts each PDF page to JPG using 'magick convert'.
Runs Tesseract OCR on each JPG to create a single-page searchable PDF.
Merges all single-page PDFs into a final document using PDFtk by manually constructing
the command string and executing with cmd.exe.
Supports resuming partially completed jobs by checking for existing output files.
Includes detailed error reporting for PDFtk merge failures.

.NOTES
Requires ImageMagick 7+, Poppler utils (pdfinfo), Tesseract OCR (with language data), and PDFtk.
The FULL PATH to each tool's executable MUST be configured in the
'Tool Path Configuration' section below after installation.
Place the PDF files you want to process in a subdirectory named 'pdfs'
relative to where this script is saved.
Output files will be placed in 'output', 'ocr_output', and 'final_merge' subdirectories.
#>

#Requires -Version 5.1 # Specify minimum PowerShell version if needed

# --- Configuration ---
$ScriptRoot = $PSScriptRoot # Directory containing the script
$PdfSourceDir = Join-Path -Path $ScriptRoot -ChildPath "pdfs"
$OutputDir = Join-Path -Path $ScriptRoot -ChildPath "output"
$OcrOutputDir = Join-Path -Path $ScriptRoot -ChildPath "ocr_output"
$FinalMergeDir = Join-Path -Path $ScriptRoot -ChildPath "final_merge"

# --- Tool Settings ---
$ImageDensity = 150 # Recommended: 300 for better OCR quality, but slower/larger files
$ImageQuality = 85  # JPG quality (0-100)
$OcrLanguage = "eng" # Tesseract language code(s) (e.g., "eng+fra" for English and French)

# --- !! IMPORTANT: Tool Path Configuration !! ---
#
# YOU MUST INSTALL THE FOLLOWING TOOLS YOURSELF:
# 1. ImageMagick (Version 7 or later recommended): Provides 'magick.exe'
#    Download: https://imagemagick.org/script/download.php
# 2. Poppler (Windows binaries): Provides 'pdfinfo.exe'
#    Download: Often found via searches like "poppler windows binaries" (e.g., https://github.com/oschwartz10612/poppler-windows/releases)
#    Extract the archive and find the 'bin' folder containing pdfinfo.exe.
# 3. Tesseract OCR: Provides 'tesseract.exe' AND language data (e.g., 'eng.traineddata')
#    Download: https://github.com/UB-Mannheim/tesseract/wiki (Recommended installer)
#    IMPORTANT: Ensure you install the required language data packs (e.g., English). The data
#    must reside in a 'tessdata' subfolder within the main Tesseract installation directory.
# 4. PDFtk Server (Free command-line version) or PDFtk Java Port: Provides 'pdftk.exe'
#    Download: https://www.pdflabs.com/tools/pdftk-server/ (Original) or search for 'pdftk-java' ports.
#
# AFTER INSTALLING, FIND THE FULL PATH TO EACH EXECUTABLE (.exe file)
# AND REPLACE THE PLACEHOLDER PATHS BELOW WITH YOUR ACTUAL PATHS.
# Use double quotes around the paths, especially if they contain spaces.
# Example: $ConvertPath = "C:\Program Files\ImageMagick-7.1.1-Q16\magick.exe"

$ConvertPath   = "C:\Path\To\ImageMagick\magick.exe"       # <-- EDIT THIS LINE with your path to magick.exe
$PdfInfoPath   = "C:\Path\To\poppler\bin\pdfinfo.exe"      # <-- EDIT THIS LINE with your path to pdfinfo.exe
$TesseractPath = "C:\Path\To\Tesseract-OCR\tesseract.exe"  # <-- EDIT THIS LINE with your path to tesseract.exe
$PdftkPath     = "C:\Path\To\PDFtk\bin\pdftk.exe"          # <-- EDIT THIS LINE with your path to pdftk.exe

# --- End Configuration ---

# --- Sanity Check Executable Paths ---
Write-Host "--- Checking Tool Paths (User Configured) ---" -ForegroundColor Yellow
$Global:AllToolsFound = $true # Use global scope for helper function visibility

# Helper function for testing a tool
function Test-ToolPath {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ToolName,

        [Parameter(Mandatory=$true)]
        [string]$ToolPath,

        [Parameter(Mandatory=$true)]
        [string[]]$VersionArgs
    )

    Write-Host "Testing '$ToolName' path: $ToolPath" -ForegroundColor Cyan
    # Basic check if placeholder path was likely not changed
     if ($ToolPath -like "*\Path\To\*") {
         Write-Error "Placeholder path detected for '$ToolName'. Please edit the script and set the correct path in the configuration section."
         $Global:AllToolsFound = $false
         return
     }

    try {
        # Test if the file actually exists first for a clearer error
        if (-not (Test-Path -Path $ToolPath -PathType Leaf)) {
             throw "File not found at specified path: '$ToolPath'. Please verify the path in the script's configuration."
        }
        # Now try executing
        & $ToolPath $VersionArgs *> $null # Execute with version args, discard output
        if ($LASTEXITCODE -ne 0) {
             if ($ToolName -ne "magick" -and $ToolName -ne "convert") {
                 Write-Error "'$ToolName' command executed using '$ToolPath' but failed (Exit Code: $LASTEXITCODE). Check installation and path."
                 $Global:AllToolsFound = $false
             } else {
                 Write-Host "'$ToolName' seems OK (Exit code ignored for magick version check)." -ForegroundColor Green
             }
        } else {
            Write-Host "'$ToolName' seems OK." -ForegroundColor Green
        }
    } catch {
        Write-Error "Failed to execute '$ToolName' using path '$ToolPath'. Is it installed correctly and is the path configured properly in the script? Error: $($_.Exception.Message)"
        $Global:AllToolsFound = $false
    }
}

# Test each tool
Test-ToolPath -ToolName "magick"    -ToolPath $ConvertPath   -VersionArgs "-version"
Test-ToolPath -ToolName "pdfinfo"   -ToolPath $PdfInfoPath   -VersionArgs "-v"
Test-ToolPath -ToolName "tesseract" -ToolPath $TesseractPath -VersionArgs "-v"
Test-ToolPath -ToolName "pdftk"     -ToolPath $PdftkPath     -VersionArgs "--version"

# Final check and exit if any tool failed
if (-not $Global:AllToolsFound) {
    Write-Error "One or more required tools could not be verified. Please check the paths in the script's configuration section and ensure the tools are installed correctly. Exiting."
    exit 1
} else {
    Write-Host "--- All tool paths verified successfully. ---" -ForegroundColor Green
}
Write-Host "" # Add a blank line for readability
# --- End Sanity Check ---


# --- Main Script Logic ---

# Ensure necessary output directories exist
Write-Host "Ensuring output directories exist..."
@( $OutputDir, $OcrOutputDir, $FinalMergeDir ) | ForEach-Object {
    if (-not (Test-Path -Path $_ -PathType Container)) {
        try {
            New-Item -ItemType Directory -Path $_ -Force -ErrorAction Stop | Out-Null
            Write-Host "Created directory: $_"
        } catch {
            Write-Error "Failed to create directory '$_'. Error: $($_.Exception.Message)"
            exit 1
        }
    }
}

# Ensure the ./pdfs input directory exists
Write-Host "Checking for input directory: $PdfSourceDir"
if (-not (Test-Path -Path $PdfSourceDir -PathType Container)) {
    Write-Error "Error: The required input directory '$PdfSourceDir' does not exist. Please create it and place PDF files inside."
    exit 1
}

# List all PDFs in the ./pdfs directory
$pdfFiles = Get-ChildItem -Path $PdfSourceDir -Filter *.pdf | Sort-Object Name
if ($pdfFiles.Count -eq 0) {
    Write-Host "No PDF files found in '$PdfSourceDir'."
    exit 0 # Not an error, just nothing to do
}

Write-Host "Found the following PDFs in '$PdfSourceDir':"
for ($i = 0; $i -lt $pdfFiles.Count; $i++) {
    Write-Host ("{0}. {1}" -f ($i + 1), $pdfFiles[$i].Name)
}

# Prompt user for selection
$choice = $null
$choiceInt = 0
while ($true) {
    try {
        $choice = Read-Host "Select a PDF to process (enter number)"
        $choiceInt = [int]::Parse($choice)
        if ($choiceInt -ge 1 -and $choiceInt -le $pdfFiles.Count) {
            break # Valid choice
        } else {
            Write-Warning "Invalid selection. Please enter a number between 1 and $($pdfFiles.Count)."
        }
    } catch {
        Write-Warning "Invalid input. Please enter a number."
    }
}

$selectedPdf = $pdfFiles[$choiceInt - 1]
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($selectedPdf.Name)

Write-Host "Processing: $($selectedPdf.FullName)" -ForegroundColor Green

# Get total page count of the PDF
$totalPages = 0
try {
    Write-Host "Getting page count..."
    $pdfInfoOutput = & $PdfInfoPath $selectedPdf.FullName 2>&1 # Capture stdout and stderr
    if ($LASTEXITCODE -ne 0) {
        throw "pdfinfo failed to execute. Exit code: $LASTEXITCODE. Output: $pdfInfoOutput"
    }
    # Try to parse the output
    $pagesLine = $pdfInfoOutput | Select-String -Pattern 'Pages:\s+(\d+)'
    if ($pagesLine -and $pagesLine.Matches[0].Groups[1].Value) {
        $totalPages = [int]$pagesLine.Matches[0].Groups[1].Value
    } else {
         throw "Could not parse page count from pdfinfo output.`nOutput was:`n$pdfInfoOutput"
    }
} catch {
     Write-Error "Failed to get page count for '$($selectedPdf.FullName)'. Error: $($_.Exception.Message)"
     exit 1
}

if ($totalPages -le 0) {
     Write-Error "Could not determine page count (or page count is zero)."
     exit 1
}
Write-Host "Total pages: $totalPages"

# Define final path early (should be defined before any potential error in Stage 3)
$finalPdfPath = Join-Path -Path $FinalMergeDir -ChildPath "${baseName}_final.pdf"

# Check if the final merged PDF exists and is complete
if (Test-Path -Path $finalPdfPath -PathType Leaf) {
    $existingPages = 0
    try {
        $finalPdfInfoOutput = & $PdfInfoPath $finalPdfPath 2>&1
         if ($LASTEXITCODE -eq 0) {
            $finalPagesLine = $finalPdfInfoOutput | Select-String -Pattern 'Pages:\s+(\d+)'
            if ($finalPagesLine -and $finalPagesLine.Matches[0].Groups[1].Value) {
                $existingPages = [int]$finalPagesLine.Matches[0].Groups[1].Value
            }
        } # Ignore errors, just assume incomplete if pdfinfo fails
    } catch {
        Write-Warning "Could not check page count of existing final PDF '$finalPdfPath'. Assuming incomplete."
    }

    if ($existingPages -eq $totalPages) {
        Write-Host "Final PDF already exists and is complete: $finalPdfPath" -ForegroundColor Green
        exit 0
    } else {
        Write-Host "Final PDF exists but seems incomplete ($existingPages / $totalPages pages). Resuming processing..." -ForegroundColor Yellow
    }
}

# Prompt for levels adjustment option
$exportChoice = $null
$blackPoint = $null
$whitePoint = $null
while ($exportChoice -ne '1' -and $exportChoice -ne '2') {
    Write-Host "Choose image export mode:"
    Write-Host "1. Export without levels adjustment (Faster)"
    Write-Host "2. Export with levels adjustment (May improve OCR for faded scans)"
    $exportChoice = Read-Host "? (1/2)"
}

if ($exportChoice -eq '2') {
    while ($blackPoint -eq $null) {
         try {
            $bpInput = Read-Host "Enter black point percentage (e.g., 10 for 10%, higher = darker blacks)"
            $blackPoint = [int]::Parse($bpInput)
            if ($blackPoint -lt 0 -or $blackPoint -gt 99) { throw "Must be between 0 and 99." }
         } catch { Write-Warning "Invalid input. Please enter a number between 0 and 99. Error: $($_.Exception.Message)" ; $blackPoint = $null }
    }
     while ($whitePoint -eq $null) {
         try {
            $wpInput = Read-Host "Enter white point percentage (e.g., 90 for 90%, lower = brighter whites)"
            $whitePoint = [int]::Parse($wpInput)
            if ($whitePoint -lt 1 -or $whitePoint -gt 100) { throw "Must be between 1 and 100." }
            if ($whitePoint -le $blackPoint) {
                 throw "White point must be greater than black point."
            }
         } catch { Write-Warning "Invalid input. Error: $($_.Exception.Message)"; $whitePoint = $null }
    }
    Write-Host "Using levels adjustment: Black=$blackPoint%, White=$whitePoint%"
}

# === Stage 1: Convert PDF pages to JPG ===
Write-Host "`n--- Starting Stage 1: Image Extraction ---" -ForegroundColor Yellow
$ProcessingErrorOccurred = $false
for ($i = 0; $i -lt $totalPages; $i++) {
    $currentPage = $i + 1
    $pageNumberStr = "{0:D3}" -f $i # Zero-padded page number (000, 001, ...)
    $outputFile = Join-Path -Path $OutputDir -ChildPath "${baseName}-${pageNumberStr}.jpg"
    # Use temp file in script's dir to avoid cluttering output dir during processing
    $tempFile = Join-Path -Path $ScriptRoot -ChildPath "temp_${baseName}-${pageNumberStr}.jpg"

    # Resume check
    if (Test-Path -Path $outputFile -PathType Leaf) {
        Write-Host "Page ${currentPage}/${totalPages}: Image already exists, skipping extraction." # Reduced verbosity
        continue
    }

    # Clean up potential leftover temp file from previous run
    if (Test-Path $tempFile) { Remove-Item $tempFile -Force }

    Write-Host "Page ${currentPage}/${totalPages}: Converting PDF page to image..."

    # Construct convert command arguments (these remain the same)
    $convertArgs = @(
        "-density", $ImageDensity,
        "$($selectedPdf.FullName)[$i]", # Page selector syntax (index starts at 0)
        "-quality", $ImageQuality,
        "-background", "white", # Ensure transparent areas become white
        "-alpha", "remove",
        "-alpha", "off"
    )

    # Convert to temporary file first
    try {
        & $ConvertPath convert $convertArgs $tempFile 2>&1 | Out-Null # Added 'convert' keyword for IMv7
        if ($LASTEXITCODE -ne 0) { throw "ImageMagick 'magick convert' failed. Exit code: $LASTEXITCODE." }
        if (-not (Test-Path $tempFile -PathType Leaf)) { throw "Temporary image file '$tempFile' was not created."}

        # Apply levels adjustment if selected, otherwise just move
        if ($exportChoice -eq '2') {
            $levelArg = "{0}%,{1}%" -f $blackPoint, $whitePoint
            Write-Host "Page ${currentPage}/${totalPages}: Applying levels $levelArg..."
            & $ConvertPath convert $tempFile -level $levelArg $outputFile 2>&1 | Out-Null # Added 'convert' keyword for IMv7
            if ($LASTEXITCODE -ne 0) { throw "ImageMagick 'magick convert' (levels) failed. Exit code: $LASTEXITCODE." }
            Remove-Item $tempFile -Force # Clean up temp file
        } else {
            Move-Item -Path $tempFile -Destination $outputFile -Force
        }

        # Final check for output file
         if (-not (Test-Path -Path $outputFile -PathType Leaf)) {
             throw "Final image file '$outputFile' was not created after conversion/move."
         }
         # Write-Host "Page ${currentPage}/${totalPages}: Successfully created $outputFile" # Optional success message

    } catch {
        Write-Error "Error processing page ${currentPage}: $($_.Exception.Message)"
        # Clean up temp file if it exists on error
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force }
        $ProcessingErrorOccurred = $true
        break # Stop processing further pages on error
    }
} # End page conversion loop

# Exit if errors occurred during conversion
if ($ProcessingErrorOccurred) {
    Write-Error "Errors occurred during image extraction. Script halted."
    exit 1
}
Write-Host "--- Stage 1: Image Extraction Complete ---" -ForegroundColor Green


# === Stage 2: Run OCR on each JPG ===
Write-Host "`n--- Starting Stage 2: OCR Processing ---" -ForegroundColor Yellow
$ProcessingErrorOccurred = $false # Reset error flag
for ($i = 0; $i -lt $totalPages; $i++) {
    $currentPage = $i + 1
    $pageNumberStr = "{0:D3}" -f $i
    $imgFile = Join-Path -Path $OutputDir -ChildPath "${baseName}-${pageNumberStr}.jpg"
    # Tesseract creates output file based on input path + specified extension. Base path required.
    $pdfFileBase = Join-Path -Path $OcrOutputDir -ChildPath "${baseName}-${pageNumberStr}"
    $expectedPdfFile = "${pdfFileBase}.pdf" # Tesseract adds .pdf extension

    # Resume check
    if (Test-Path -Path $expectedPdfFile -PathType Leaf) {
        Write-Host "Page ${currentPage}/${totalPages}: OCR output PDF already exists, skipping OCR."
        continue
    }

    # Check if input image exists (should always exist if Stage 1 completed)
    if (-not (Test-Path -Path $imgFile -PathType Leaf)) {
        Write-Error "Error: Expected image file missing for OCR: $imgFile. Stage 1 might have failed partially."
        $ProcessingErrorOccurred = $true
        break
    }

    Write-Host "Page ${currentPage}/${totalPages}: Performing OCR..."
    $tesseractArgs = @(
        $imgFile,         # Input file
        $pdfFileBase,     # Output base name (Tesseract adds .pdf)
        "-l", $OcrLanguage,
        "pdf"             # Output format (creates searchable PDF)
    )

    try {
        $tessOutput = & $TesseractPath $tesseractArgs 2>&1 # Capture output for potential debugging
        if ($LASTEXITCODE -ne 0) {
             # Check for TESSDATA_PREFIX specific error
             if ($tessOutput -match "Please make sure the TESSDATA_PREFIX") {
                 throw "Tesseract language data not found. Ensure '$OcrLanguage.traineddata' exists in a 'tessdata' subfolder of your Tesseract installation (parent directory of '$TesseractPath'). Tesseract Output: $tessOutput"
             } else {
                 throw "Tesseract OCR failed. Exit code: $LASTEXITCODE. Output: $tessOutput"
             }
        }

        # Verify Tesseract actually created the PDF
        if (-not (Test-Path -Path $expectedPdfFile -PathType Leaf)) {
            throw "Tesseract completed but output PDF '$expectedPdfFile' was not found. Output: $tessOutput"
        }
         # Write-Host "Page ${currentPage}/${totalPages}: Successfully created $expectedPdfFile" # Optional success message

    } catch {
        Write-Error "Error during Tesseract OCR on page ${currentPage} ($imgFile): $($_.Exception.Message)"
        $ProcessingErrorOccurred = $true
        break # Stop processing further pages on error
    }
} # End OCR loop

# Exit if errors occurred during OCR
if ($ProcessingErrorOccurred) {
    Write-Error "Errors occurred during OCR processing. Script halted."
    exit 1
}
Write-Host "--- Stage 2: OCR Processing Complete ---" -ForegroundColor Green


# === Stage 3: Verify and Merge OCR PDFs (Manual Command String via cmd) ===
Write-Host "`n--- Starting Stage 3: Verification and Merging ---" -ForegroundColor Yellow

# Verify that all expected OCR PDFs exist before merging
Write-Host "Verifying all $totalPages OCR page PDFs exist..."
$missingPages = $false
$ocrPdfFilesForMerge = [System.Collections.Generic.List[string]]::new()

for ($i = 0; $i -lt $totalPages; $i++) {
    $pageNumberStr = "{0:D3}" -f $i
    $pdfFile = Join-Path -Path $OcrOutputDir -ChildPath "${baseName}-${pageNumberStr}.pdf"

    if (-not (Test-Path -Path $pdfFile -PathType Leaf)) {
        Write-Error "FATAL: Missing expected OCR PDF after processing: $pdfFile"
        $missingPages = $true
    } else {
        # Add path enclosed in double quotes for safety when building command string
        $ocrPdfFilesForMerge.Add("""$($pdfFile)""")
    }
}

if ($missingPages) {
    Write-Error "Cannot merge because some intermediate OCR PDF files are missing. Please check errors from Stage 2."
    exit 1
}
Write-Host "All intermediate OCR PDFs verified."

# Merge all individual OCR PDFs into one final PDF
Write-Host "Merging $totalPages individual PDFs into final document: $finalPdfPath ..."

# --- Construct the command string manually ---
# Quote the pdftk executable path
$pdftkCmd = """$($PdftkPath)"""

# Join the already quoted input file paths with spaces
$inputFilesString = $ocrPdfFilesForMerge -join " "

# Quote the output path
$quotedOutputPath = """$($finalPdfPath)"""

# Build the full command string for cmd.exe
$fullCommand = "$pdftkCmd $inputFilesString cat output $quotedOutputPath"

Write-Host "Executing pdftk command via cmd /c..." -ForegroundColor DarkGray # Debug output, removed full command to avoid excessive length in console

$cmdOutput = "" # Initialize variable

try {
    # --- Execute using cmd /c ---
    $cmdOutput = cmd /c $fullCommand 2>&1 # Capture output from cmd

    # Check $LASTEXITCODE. cmd /c should preserve the exit code of the executed command.
    if ($LASTEXITCODE -ne 0) {
        # Throw error including the output from cmd/pdftk
        throw "PDFtk merge failed when executed via cmd /c. Exit code: $LASTEXITCODE. Output: $cmdOutput"
    }

    # Optional: Double-check output for error strings just in case exit code was 0 despite errors
    if ($cmdOutput -match "Error:" -or $cmdOutput -match "Exception") {
         Write-Warning "PDFtk seemed to succeed (Exit Code: 0 via cmd /c) but output contains 'Error:' or 'Exception'. Output: $cmdOutput"
    }

    # Final check on the merged file existence
    if (-not (Test-Path -Path $finalPdfPath -PathType Leaf)) {
         throw "PDFtk command executed via cmd /c, but the final merged file '$finalPdfPath' was not found. Exit Code: $LASTEXITCODE. Output: $cmdOutput"
    }

     # Optional: Verify page count of final PDF again
     $finalCheckPages = 0
     $finalCheckOutput = & $PdfInfoPath $finalPdfPath 2>&1
     if ($LASTEXITCODE -eq 0) {
         $finalCheckLine = $finalCheckOutput | Select-String -Pattern 'Pages:\s+(\d+)'
         if ($finalCheckLine -and $finalCheckLine.Matches[0].Groups[1].Value) {
             $finalCheckPages = [int]$finalCheckLine.Matches[0].Groups[1].Value
         }
     }
     # Only warn if page count check was possible and the count mismatch is definite
     if ($finalCheckPages -gt 0 -and $finalCheckPages -ne $totalPages) {
          Write-Warning "Final PDF created, but page count ($finalCheckPages) does not match expected ($totalPages). The merge might be incomplete."
     } elseif ($finalCheckPages -eq $totalPages) {
          Write-Host "Final PDF page count verified."
     } else {
         Write-Warning "Could not verify final PDF page count."
     }

} catch {
    # Report the error, including any output captured from cmd/pdftk
    Write-Error "Error merging PDFs with PDFtk via cmd /c: $($_.Exception.Message) `nCommand Output was (if any):`n$cmdOutput"
    # Clean up potentially incomplete final PDF
    if (-not [string]::IsNullOrEmpty($finalPdfPath) -and (Test-Path $finalPdfPath)) {
        Write-Warning "Attempting to remove potentially incomplete final PDF: $finalPdfPath"
        Remove-Item $finalPdfPath -Force -ErrorAction SilentlyContinue
    }
    exit 1
} finally {
     # No temporary files in this approach
     Write-Host "Stage 3 processing finished."
}


Write-Host "`n--- Processing Complete ---" -ForegroundColor Green
Write-Host "Final searchable PDF created successfully: $finalPdfPath"
exit 0 # Explicit success exit code
