# Configuration
$Config = @{
    FileFilter = "*.pdf"
    TestPath   = ".\test"
    LivePath   = ".\live"
    ResultPath = ".\result"
}

function Test-UnmatchedFiles {
    
    Compare-Object (Get-ChildItem ".\test\*.pdf") (Get-ChildItem ".\live\*.pdf") -Property Name |
    Select-Object Name, @{
        Name = "Location"
        Expression = { if ($_.SideIndicator -eq "<=") { "Only in .\test" } else { "Only in .\live" } }
    } | Format-Table -AutoSize


}

function Invoke-PdfCompare {
    # Create output directory
    New-Item -ItemType Directory -Force -Path ".\results" | Out-Null

    # Compare matching PDFs
    Get-ChildItem -Path ".\test\*.pdf" | ForEach-Object {
        $original = $_.FullName
        $revised = Join-Path ".\live" $_.Name
        $output = Join-Path ".\results" "diff_$($_.Name)"

        if (Test-Path $revised) {
            Write-Host "Comparing: $($_.Name)..." -ForegroundColor Yellow

            & ".\diff-pdf.exe" `
                --output-diff=$output `
                -s -m $original $revised
        }
        else {
            Write-Warning "Skipped: $($_.Name) (No matching file in .\live)"
        }
    }
}

function Invoke-QuickDelete {
    Write-Host "`nWARNING: This will permanently delete all files in .\test, .\live, and .\results." -ForegroundColor Yellow

    $confirm = Read-Host "Are you sure? (Y/N)"
    if ($confirm -ne "Y") {
        Write-Host "Quick Delete cancelled." -ForegroundColor Cyan
        return
    }


    Write-Host "`nDeleting files..." -ForegroundColor Yellow

    @(".\test", ".\live", ".\results") | ForEach-Object {
        if (Test-Path $_) {
            Remove-Item "$_\*" -Force -Recurse -ErrorAction SilentlyContinue
            Write-Host "Cleared: $_" -ForegroundColor Green
        }
        else {
            Write-Host "Folder not found: $_" -ForegroundColor Red
        }
    }

    Write-Host "`nQuick Delete completed." -ForegroundColor Cyan
}

function Show-Menu {
    Write-Host ""
    Write-Host "==================================" -ForegroundColor Cyan
    Write-Host "       PDF BULK COMPARE           " -ForegroundColor Cyan
    Write-Host "==================================" -ForegroundColor Cyan

    Write-Host "1. Check Unmatched Files" -ForegroundColor Cyan
    Write-Host "2. Compare PDFs" -ForegroundColor Cyan
    Write-Host "3. Quick Delete" -ForegroundColor Cyan
    Write-Host "4. Exit" -ForegroundColor Cyan
}

do {
    Show-Menu
    Write-Host "Select an option: " -ForegroundColor Cyan -NoNewline
    $choice = Read-Host

    switch ($choice) {
        '1' { Test-UnmatchedFiles }
        '2' { Invoke-PdfCompare }
        '3' { Invoke-QuickDelete }
        '4' { Write-Host "Exit!" -ForegroundColor Green }
        default { Write-Host "Invalid choice" }
    }

} while ($choice -ne '4')