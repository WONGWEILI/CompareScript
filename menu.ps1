# ============================================================
# PDF BULK COMPARE
# ============================================================

# Base folder = folder containing this PowerShell script
$BasePath = $PSScriptRoot

# Configuration
$Config = @{
    FileFilter = "*.pdf"
    TestPath   = Join-Path $BasePath "test"
    LivePath   = Join-Path $BasePath "live"
    ResultPath = Join-Path $BasePath "results"
    DiffPdfExe = Join-Path $BasePath "diff-pdf.exe"
    TextResultPath = Join-Path $BasePath "textresult"
    PdfToTextExe   = "pdftotext"
    GitExe         = "git"
}
# ============================================================
# Change all file to pdf
# ============================================================
function Rename-FilesToPdf  {

    $folders = @(
        $Config.TestPath
        $Config.LivePath
    )

    foreach ($folder in $folders) {

        if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
            Write-Warning "Folder not found: $folder"
            continue
        }

        try {
            $files = @(
                Get-ChildItem `
                    -LiteralPath $folder `
                    -File `
                    -ErrorAction Stop
            )
            if ($files.Count -eq 0) {
                Write-Host " No file in '$folder'"
                continue
            }

            foreach ($file in $files) {

                $newName = [System.IO.Path]::ChangeExtension(
                    $file.Name,
                    ".pdf"
                )

                # Skip files already ending in .pdf
                if ($file.Name -ceq $newName) {
                    continue
                }

                $newPath = Join-Path $folder $newName

                # Prevent overwriting an existing file
                if (Test-Path -LiteralPath $newPath) {
                    Write-Warning "Cannot rename '$($file.Name)' because '$newName' already exists."
                    continue
                }

                Rename-Item `
                    -LiteralPath $file.FullName `
                    -NewName $newName `
                    -ErrorAction Stop

                Write-Host "Renamed: $($file.Name) -> $newName" -ForegroundColor Green
            }
        }
        catch {
            Write-Warning "Failed processing '$folder': $($_.Exception.Message)"
        }
    }
}



# ============================================================
# Check for unmatched filenames
# ============================================================
function Test-UnmatchedFiles {
    $testFolder = $Config.TestPath
    $liveFolder = $Config.LivePath
    $filter     = $Config.FileFilter

    Write-Host ""
    Write-Host "Checking files..." -ForegroundColor Cyan

    # Validate directories
    if (-not (Test-Path -LiteralPath $testFolder -PathType Container)) {
        Write-Warning "Test folder not found: $testFolder"
        return
    }

    if (-not (Test-Path -LiteralPath $liveFolder -PathType Container)) {
        Write-Warning "Live folder not found: $liveFolder"
        return
    }

    # Get files
    try {
        $testFiles = @(
            Get-ChildItem `
                -LiteralPath $testFolder `
                -Filter $filter `
                -File `
                -ErrorAction Stop
        )

        $liveFiles = @(
            Get-ChildItem `
                -LiteralPath $liveFolder `
                -Filter $filter `
                -File `
                -ErrorAction Stop
        )
    }
    catch {
        Write-Warning "Failed to read files: $($_.Exception.Message)"
        return
    }

    # Both folders empty
    if ($testFiles.Count -eq 0 -and $liveFiles.Count -eq 0) {
        Write-Host "Both folders contain no pdf files." -ForegroundColor Cyan
        return
    }

    # Test empty
    if ($testFiles.Count -eq 0) {
        Write-Host ""
        Write-Host "No matching files found in:" -ForegroundColor Yellow
        Write-Host $testFolder -ForegroundColor Yellow
        Write-Host ""

        $liveFiles |
            Select-Object Name, @{
                Name       = "Location"
                Expression = { "Only in Live" }
            } |
            Format-Table -AutoSize

        return
    }

    # Live empty
    if ($liveFiles.Count -eq 0) {
        Write-Host ""
        Write-Host "No matching files found in:" -ForegroundColor Yellow
        Write-Host $liveFolder -ForegroundColor Yellow
        Write-Host ""

        $testFiles |
            Select-Object Name, @{
                Name       = "Location"
                Expression = { "Only in Test" }
            } |
            Format-Table -AutoSize

        return
    }

    # Compare filenames only
    $comparison = Compare-Object `
        -ReferenceObject $testFiles `
        -DifferenceObject $liveFiles `
        -Property Name

    if (-not $comparison) {
        Write-Host ""
        Write-Host "Both folders contain the same PDF filenames." `
            -ForegroundColor Green

        return
    }

    Write-Host ""
    Write-Host "Unmatched files:" -ForegroundColor Yellow
    Write-Host ""

    $comparison |
        Select-Object Name, @{
            Name = "Location"

            Expression = {
                if ($_.SideIndicator -eq "<=") {
                    "Only in Test"
                }
                else {
                    "Only in Live"
                }
            }
        } |
        Sort-Object Name |
        Format-Table -AutoSize
}


# ============================================================
# Compare PDFs using diff-pdf.exe
# ============================================================
function Invoke-PdfCompare {
    $testFolder = $Config.TestPath
    $liveFolder = $Config.LivePath
    $resultsDir = $Config.ResultPath
    $diffPdfExe = $Config.DiffPdfExe
    $filter     = $Config.FileFilter

    Write-Host ""

    # ========================================================
    # Validate directories
    # ========================================================
    if (-not (Test-Path -LiteralPath $testFolder -PathType Container)) {
        Write-Warning "Test folder not found: $testFolder"
        return
    }

    if (-not (Test-Path -LiteralPath $liveFolder -PathType Container)) {
        Write-Warning "Live folder not found: $liveFolder"
        return
    }

    # ========================================================
    # Validate diff-pdf.exe
    # ========================================================
    if (-not (Test-Path -LiteralPath $diffPdfExe -PathType Leaf)) {
        Write-Warning "diff-pdf.exe was not found:"
        Write-Warning $diffPdfExe
        return
    }

    # ========================================================
    # Create results folder
    # ========================================================
    try {
        if (-not (Test-Path -LiteralPath $resultsDir -PathType Container)) {
            New-Item `
                -ItemType Directory `
                -Path $resultsDir `
                -Force `
                -ErrorAction Stop |
                Out-Null
        }
    }
    catch {
        Write-Warning "Could not create results folder: $($_.Exception.Message)"
        return
    }

    # ========================================================
    # Load test files
    # ========================================================
    try {
        $testFiles = @(
            Get-ChildItem `
                -LiteralPath $testFolder `
                -Filter $filter `
                -File `
                -ErrorAction Stop
        )
    }
    catch {
        Write-Warning "Could not read test files: $($_.Exception.Message)"
        return
    }

    if ($testFiles.Count -eq 0) {
        Write-Warning "No PDF files found in:"
        Write-Warning $testFolder
        return
    }

    # ========================================================
    # Counters
    # ========================================================
    $identicalCount = 0
    $differentCount = 0
    $missingCount   = 0
    $failedCount    = 0

    # Track filenames in comparison processing order
    $failedFiles    = @()
    $differentFiles = @()
    $missingFiles   = @()

    # ========================================================
    # Compare PDFs
    # ========================================================
    foreach ($file in $testFiles) {

        $original = $file.FullName
        $revised  = Join-Path $liveFolder $file.Name
        $output   = Join-Path $resultsDir "diff_$($file.Name)"

        Write-Host ""
        Write-Host "Comparing: $($file.Name)" -ForegroundColor Cyan

        # ----------------------------------------------------
        # Matching Live file does not exist
        # ----------------------------------------------------
        if (-not (Test-Path -LiteralPath $revised -PathType Leaf)) {
            Write-Host "  MISSING   : No matching file in Live." `
                -ForegroundColor Yellow

            $missingCount++
            $missingFiles += $file.Name
            continue
        }

        # ----------------------------------------------------
        # Remove old diff result
        # ----------------------------------------------------
        if (Test-Path -LiteralPath $output -PathType Leaf) {
            try {
                Remove-Item `
                    -LiteralPath $output `
                    -Force `
                    -ErrorAction Stop
            }
            catch {
                Write-Warning "Failed to remove old result: $output"
                $failedCount++
                $failedFiles += $file.Name
                continue
            }
        }

        # ----------------------------------------------------
        # Run diff-pdf
        # ----------------------------------------------------
        try {
            & $diffPdfExe `
                -s `
                -m `
                $original `
                $revised

            $exitCode = $LASTEXITCODE

            switch ($exitCode) {

                # --------------------------------------------
                # Exit code 0 = PDFs are identical
                # --------------------------------------------
                0 {
                    Write-Host "  IDENTICAL" -ForegroundColor Green
                    $identicalCount++
                }

                # --------------------------------------------
                # Exit code 1 = PDFs are different
                # --------------------------------------------
                1 {
                    Write-Host "  DIFFERENT" -ForegroundColor Yellow
                    $differentCount++
                    $differentFiles += $file.Name

                    # Generate output PDF only when files are different
                    try {
                        & $diffPdfExe `
                            "--output-diff=$output" `
                            -s `
                            -m `
                            $original `
                            $revised

                        if ($LASTEXITCODE -ne 1) {
                            Write-Warning "Failed to generate difference PDF for '$($file.Name)'. Exit code: $LASTEXITCODE"
                        }
                    }
                    catch {
                        Write-Warning "Failed to generate difference PDF for '$($file.Name)': $($_.Exception.Message)"
                    }
                }

                # --------------------------------------------
                # Anything else = unexpected failure
                # --------------------------------------------
                default {
                    Write-Warning (
                        "FAILED: diff-pdf returned unexpected exit code " +
                        "$exitCode for '$($file.Name)'."
                    )

                    $failedCount++
                    $failedFiles += $file.Name
                }
            }
        }
        catch {
            Write-Warning (
                "FAILED: '$($file.Name)': " +
                $_.Exception.Message
            )

            $failedCount++
            $failedFiles += $file.Name
        }
    }

    Show-CompareSummary `
        -IdenticalCount $identicalCount `
        -DifferentCount $differentCount `
        -MissingCount $missingCount `
        -FailedCount $failedCount `
        -TotalCount $testFiles.Count `
        -FailedFiles $failedFiles `
        -DifferentFiles $differentFiles `
        -MissingFiles $missingFiles `
        -ResultsDir $resultsDir
}


# ============================================================
# Print comparison summary and filename lists
# ============================================================
function Show-CompareSummary {
    param(
        [int]$IdenticalCount,
        [int]$DifferentCount,
        [int]$MissingCount,
        [int]$FailedCount,
        [int]$TotalCount,
        [string[]]$FailedFiles = @(),
        [string[]]$DifferentFiles = @(),
        [string[]]$MissingFiles = @(),
        [string]$ResultsDir
    )

    $FailedFiles    = @($FailedFiles)
    $DifferentFiles = @($DifferentFiles)
    $MissingFiles   = @($MissingFiles)

    # ========================================================
    # Summary
    # ========================================================
    Write-Host ""
    Write-Host "==================================" -ForegroundColor Cyan
    Write-Host "       COMPARISON SUMMARY         " -ForegroundColor Cyan
    Write-Host "==================================" -ForegroundColor Cyan

    Write-Host "Identical : " -NoNewline
    Write-Host $identicalCount -ForegroundColor Green

    Write-Host "Different : " -NoNewline
    Write-Host $differentCount -ForegroundColor Yellow

    Write-Host "Missing   : " -NoNewline
    Write-Host $missingCount -ForegroundColor Yellow

    Write-Host "Failed    : " -NoNewline
    Write-Host $failedCount -ForegroundColor Red

    Write-Host "----------------------------------"

    Write-Host "Total     : " -NoNewline
    Write-Host $TotalCount -ForegroundColor Cyan

    Write-Host "==================================" -ForegroundColor Cyan

    # ========================================================
    # Failed and different filename lists
    # ========================================================
    Write-Host ""
    Write-Host "FAILED FILES ($($failedFiles.Count))" -ForegroundColor Yellow

    if ($failedFiles.Count -gt 0) {
        Write-Host "----------------------------------" -ForegroundColor Yellow

        for ($i = 0; $i -lt $failedFiles.Count; $i++) {
            Write-Host "$($i + 1). $($failedFiles[$i])" -ForegroundColor Yellow
        }
    }

    Write-Host ""
    Write-Host "DIFFERENT FILES ($($differentFiles.Count))" -ForegroundColor Yellow

    if ($differentFiles.Count -gt 0) {
        Write-Host "----------------------------------" -ForegroundColor Yellow

        for ($i = 0; $i -lt $differentFiles.Count; $i++) {
            Write-Host "$($i + 1). $($differentFiles[$i])" -ForegroundColor Yellow
        }
    }

    Write-Host ""
    Write-Host "MISSING FILES ($($missingFiles.Count))" -ForegroundColor Yellow

    if ($missingFiles.Count -gt 0) {
        Write-Host "----------------------------------" -ForegroundColor Yellow

        for ($i = 0; $i -lt $missingFiles.Count; $i++) {
            Write-Host "$($i + 1). $($missingFiles[$i])" -ForegroundColor Yellow
        }
    }

    Write-Host ""
    Write-Host "Results folder:" -ForegroundColor Cyan
    Write-Host $resultsDir
}


# ============================================================
# Compare PDF text using pdftotext + git diff
# ============================================================
function Invoke-PdfTextCompare {
    $testFolder = $Config.TestPath
    $liveFolder = $Config.LivePath
    $resultsDir = $Config.TextResultPath
    $pdfToText  = $Config.PdfToTextExe
    $gitExe     = $Config.GitExe
    $filter     = $Config.FileFilter
    $tempDir    = Join-Path $env:TEMP "pdf-textdiff"

    Write-Host ""

    # ========================================================
    # Validate directories
    # ========================================================
    if (-not (Test-Path -LiteralPath $testFolder -PathType Container)) {
        Write-Warning "Test folder not found: $testFolder"
        return
    }

    if (-not (Test-Path -LiteralPath $liveFolder -PathType Container)) {
        Write-Warning "Live folder not found: $liveFolder"
        return
    }

    # ========================================================
    # Validate pdftotext and git
    # ========================================================
    if (-not (Get-Command $pdfToText -ErrorAction SilentlyContinue)) {
        Write-Warning "pdftotext was not found. Install Poppler and add it to PATH."
        return
    }

    if (-not (Get-Command $gitExe -ErrorAction SilentlyContinue)) {
        Write-Warning "git was not found. Install Git and add it to PATH."
        return
    }

    # ========================================================
    # Create results and temp folders
    # ========================================================
    try {
        foreach ($dir in @($resultsDir, $tempDir)) {
            if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
                New-Item `
                    -ItemType Directory `
                    -Path $dir `
                    -Force `
                    -ErrorAction Stop |
                    Out-Null
            }
        }
    }
    catch {
        Write-Warning "Could not create folder: $($_.Exception.Message)"
        return
    }

    # ========================================================
    # Load test files
    # ========================================================
    try {
        $testFiles = @(
            Get-ChildItem `
                -LiteralPath $testFolder `
                -Filter $filter `
                -File `
                -ErrorAction Stop
        )
    }
    catch {
        Write-Warning "Could not read test files: $($_.Exception.Message)"
        return
    }

    if ($testFiles.Count -eq 0) {
        Write-Warning "No PDF files found in:"
        Write-Warning $testFolder
        return
    }

    # ========================================================
    # Counters
    # ========================================================
    $identicalCount = 0
    $differentCount = 0
    $missingCount   = 0
    $failedCount    = 0

    $failedFiles    = @()
    $differentFiles = @()
    $missingFiles   = @()

    # ========================================================
    # Compare PDF text
    # ========================================================
    foreach ($file in $testFiles) {

        $original = $file.FullName
        $revised  = Join-Path $liveFolder $file.Name
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $output   = Join-Path $resultsDir "textdiff_$baseName.diff"
        $testTxt  = Join-Path $tempDir "test_$baseName.txt"
        $liveTxt  = Join-Path $tempDir "live_$baseName.txt"

        Write-Host ""
        Write-Host "Comparing text: $($file.Name)" -ForegroundColor Cyan

        # ----------------------------------------------------
        # Matching Live file does not exist
        # ----------------------------------------------------
        if (-not (Test-Path -LiteralPath $revised -PathType Leaf)) {
            Write-Host "  MISSING   : No matching file in Live." `
                -ForegroundColor Yellow

            $missingCount++
            $missingFiles += $file.Name
            continue
        }

        # ----------------------------------------------------
        # Remove old diff result
        # ----------------------------------------------------
        if (Test-Path -LiteralPath $output -PathType Leaf) {
            try {
                Remove-Item `
                    -LiteralPath $output `
                    -Force `
                    -ErrorAction Stop
            }
            catch {
                Write-Warning "Failed to remove old result: $output"
                $failedCount++
                $failedFiles += $file.Name
                continue
            }
        }

        # ----------------------------------------------------
        # Extract text from both PDFs
        # ----------------------------------------------------
        & $pdfToText -raw -enc UTF-8 $original $testTxt
        $testExit = $LASTEXITCODE

        & $pdfToText -raw -enc UTF-8 $revised $liveTxt
        $liveExit = $LASTEXITCODE

        if ($testExit -ne 0 -or $liveExit -ne 0) {
            Write-Warning "FAILED: pdftotext could not read '$($file.Name)'."
            $failedCount++
            $failedFiles += $file.Name
            continue
        }

        # ----------------------------------------------------
        # Run git diff (character level)
        # ----------------------------------------------------
        & $gitExe diff `
            --no-index `
            --no-color `
            --text `
            --word-diff=plain `
            --word-diff-regex=. `
            "--output=$output" `
            $testTxt `
            $liveTxt

        $exitCode = $LASTEXITCODE

        switch ($exitCode) {

            # Exit code 0 = text is identical
            0 {
                Write-Host "  IDENTICAL" -ForegroundColor Green
                $identicalCount++
                Remove-Item -LiteralPath $output -Force -ErrorAction SilentlyContinue
            }

            # Exit code 1 = text is different
            1 {
                Write-Host "  DIFFERENT" -ForegroundColor Yellow
                $differentCount++
                $differentFiles += $file.Name
            }

            # Anything else = unexpected failure
            default {
                Write-Warning (
                    "FAILED: git diff returned unexpected exit code " +
                    "$exitCode for '$($file.Name)'."
                )

                $failedCount++
                $failedFiles += $file.Name
                Remove-Item -LiteralPath $output -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # ========================================================
    # Clean up temp text files
    # ========================================================
    Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue

    Show-CompareSummary `
        -IdenticalCount $identicalCount `
        -DifferentCount $differentCount `
        -MissingCount $missingCount `
        -FailedCount $failedCount `
        -TotalCount $testFiles.Count `
        -FailedFiles $failedFiles `
        -DifferentFiles $differentFiles `
        -MissingFiles $missingFiles `
        -ResultsDir $resultsDir
}


# ============================================================
# Delete contents of test, live, results, and textresult
# ============================================================
function Invoke-QuickDelete {
    $folders = @(
        $Config.TestPath
        $Config.LivePath
        $Config.ResultPath
        $Config.TextResultPath
    )

    Write-Host ""
    Write-Host "WARNING" -ForegroundColor Red
    Write-Host (
        "This will permanently delete all files and folders inside:"
    ) -ForegroundColor Yellow

    Write-Host ""

    foreach ($folder in $folders) {
        Write-Host "  $folder" -ForegroundColor Yellow
    }

    Write-Host ""

    $confirm = Read-Host "Type DELETE to continue"

    if ($confirm -cne "DELETE") {
        Write-Host ""
        Write-Host "Quick Delete cancelled." -ForegroundColor Cyan
        return
    }

    Write-Host ""
    Write-Host "Deleting files..." -ForegroundColor Yellow

    foreach ($folder in $folders) {

        if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
            Write-Host "Folder not found: $folder" -ForegroundColor DarkYellow
            continue
        }

        try {
            $items = @(
                Get-ChildItem `
                    -LiteralPath $folder `
                    -Force `
                    -ErrorAction Stop
            )

            if ($items.Count -eq 0) {
                Write-Host "Already empty: $folder" -ForegroundColor Cyan
                continue
            }

            $items |
                Remove-Item `
                    -Force `
                    -Recurse `
                    -ErrorAction Stop

            Write-Host "Cleared: $folder" -ForegroundColor Green
        }
        catch {
            Write-Warning (
                "Failed to clear '$folder': " +
                $_.Exception.Message
            )
        }
    }

    Write-Host ""
    Write-Host "Quick Delete completed." -ForegroundColor Cyan
}


# ============================================================
# Menu
# ============================================================
function Show-Menu {
    Write-Host ""
    Write-Host "==================================" -ForegroundColor Cyan
    Write-Host "       PDF BULK COMPARE           " -ForegroundColor Cyan
    Write-Host "==================================" -ForegroundColor Cyan
    Write-Host "1. Change File Extensions to PDF"
    Write-Host "2. Check Unmatched Files"
    Write-Host "3. Compare PDF Text"
    Write-Host "4. Compare PDFs"
    Write-Host "5. Quick Delete"
    Write-Host "6. Exit"
    Write-Host "Note : Make sure to Check Unmatched Files before comparing PDFs" -ForegroundColor Yellow
    Write-Host ""
}


# ============================================================
# Main loop
# ============================================================
do {
    Show-Menu

    Write-Host "Select an option: " `
        -ForegroundColor Cyan `
        -NoNewline

    $choice = Read-Host

    switch ($choice) {
        "1" {
            Rename-FilesToPdf 
        }

        "2" {
            Test-UnmatchedFiles
        }

        "3" {
            Invoke-PdfTextCompare
        }

        "4" {
            Invoke-PdfCompare
        }

        "5" {
            Invoke-QuickDelete
        }

        "6" {
            Write-Host ""
            Write-Host "Exit!" -ForegroundColor Cyan
        }

        default {
            Write-Host ""
            Write-Host "Invalid choice. Please enter 1, 2, 3, 4, 5, or 6." `
                -ForegroundColor Red
        }
    }

} while ($choice -ne "6")