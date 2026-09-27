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
        Write-Host "Both folders contain no matching files." -ForegroundColor Cyan
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

    # Validate directories
    if (-not (Test-Path -LiteralPath $testFolder -PathType Container)) {
        Write-Warning "Test folder not found: $testFolder"
        return
    }

    if (-not (Test-Path -LiteralPath $liveFolder -PathType Container)) {
        Write-Warning "Live folder not found: $liveFolder"
        return
    }

    # Validate diff-pdf.exe
    if (-not (Test-Path -LiteralPath $diffPdfExe -PathType Leaf)) {
        Write-Warning "diff-pdf.exe was not found:"
        Write-Warning $diffPdfExe
        return
    }

    # Create results folder
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

    # Load test files
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

    $comparedCount = 0
    $skippedCount  = 0
    $failedCount   = 0

    foreach ($file in $testFiles) {
        $original = $file.FullName
        $revised  = Join-Path $liveFolder $file.Name
        $output   = Join-Path $resultsDir "diff_$($file.Name)"

        # Skip unmatched file
        if (-not (Test-Path -LiteralPath $revised -PathType Leaf)) {
            Write-Warning "Skipped: $($file.Name) - no matching file in Live."
            $skippedCount++
            continue
        }

        Write-Host "Comparing: $($file.Name)" -ForegroundColor Yellow

        # Remove an old result for this file
        if (Test-Path -LiteralPath $output -PathType Leaf) {
            try {
                Remove-Item `
                    -LiteralPath $output `
                    -Force `
                    -ErrorAction Stop
            }
            catch {
                Write-Warning "Could not remove old result: $output"
                $failedCount++
                continue
            }
        }

        try {
            # Run diff-pdf
            & $diffPdfExe "--output-diff=$output" -s -m $original $revised

            $exitCode = $LASTEXITCODE

            if ($exitCode -eq 0) {
                Write-Host "Completed: $($file.Name)" -ForegroundColor Green
                $comparedCount++
            }
            else {
                Write-Warning (
                    "diff-pdf returned exit code $exitCode for '$($file.Name)'."
                )

                $failedCount++
            }
        }
        catch {
            Write-Warning (
                "Failed to compare '$($file.Name)': " +
                $_.Exception.Message
            )

            $failedCount++
        }
    }

    Write-Host ""
    Write-Host "Comparison finished." -ForegroundColor Cyan
    Write-Host "Compared : $comparedCount"
    Write-Host "Skipped  : $skippedCount"
    Write-Host "Failed   : $failedCount"
    Write-Host ""
    Write-Host "Results folder:" -ForegroundColor Cyan
    Write-Host $resultsDir
}


# ============================================================
# Delete contents of test, live, and results
# ============================================================
function Invoke-QuickDelete {
    $folders = @(
        $Config.TestPath
        $Config.LivePath
        $Config.ResultPath
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
    Write-Host "1. Check Unmatched Files"
    Write-Host "2. Compare PDFs"
    Write-Host "3. Quick Delete"
    Write-Host "4. Exit"
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
            Test-UnmatchedFiles
        }

        "2" {
            Invoke-PdfCompare
        }

        "3" {
            Invoke-QuickDelete
        }

        "4" {
            Write-Host ""
            Write-Host "Exit!" -ForegroundColor Cyan
        }

        default {
            Write-Host ""
            Write-Host "Invalid choice. Please enter 1, 2, 3, or 4." `
                -ForegroundColor Red
        }
    }

} while ($choice -ne "4")