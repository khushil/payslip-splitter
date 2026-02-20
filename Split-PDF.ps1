#Requires -Version 5.1
<#
.SYNOPSIS
    Splits a PDF into individual pages. No Acrobat Pro needed!

.DESCRIPTION
    Drag-and-drop a PDF onto this script, or run it from PowerShell.
    Uses the free PdfSharp library (auto-downloaded on first run).
    Output pages go into a subfolder next to the original PDF.

.EXAMPLE
    .\Split-PDF.ps1 "C:\Documents\payslips.pdf"
    .\Split-PDF.ps1  (will open a file picker dialog)
#>

param(
    [string]$PdfPath
)

# --- Config ---
$LibFolder = Join-Path $PSScriptRoot "lib"
$PdfSharpDll = Join-Path $LibFolder "PdfSharp.dll"
$NuGetUrl = "https://www.nuget.org/api/v2/package/PdfSharp/1.50.5147"

# --- Functions ---
function Install-PdfSharp {
    if (Test-Path $PdfSharpDll) { return }

    Write-Host "First run - downloading PdfSharp (free, open-source)..." -ForegroundColor Cyan
    New-Item -ItemType Directory -Path $LibFolder -Force | Out-Null

    $zipPath = Join-Path $env:TEMP "pdfsharp.nupkg.zip"
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $NuGetUrl -OutFile $zipPath -UseBasicParsing
    }
    catch {
        Write-Host "ERROR: Failed to download PdfSharp. Check your internet connection." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        pause
        exit 1
    }

    $extractPath = Join-Path $env:TEMP "pdfsharp_extract"
    if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force

    # Find the .NET Framework 4.x compatible DLL
    $dll = Get-ChildItem -Path $extractPath -Filter "PdfSharp.dll" -Recurse |
           Where-Object { $_.FullName -match "net[0-9]|netstandard" } |
           Select-Object -First 1

    if (-not $dll) {
        # Fallback: just grab any PdfSharp.dll
        $dll = Get-ChildItem -Path $extractPath -Filter "PdfSharp.dll" -Recurse |
               Select-Object -First 1
    }

    if (-not $dll) {
        Write-Host "ERROR: Could not find PdfSharp.dll in the downloaded package." -ForegroundColor Red
        pause
        exit 1
    }

    Copy-Item $dll.FullName -Destination $PdfSharpDll -Force

    # Cleanup
    Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
    Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue

    Write-Host "PdfSharp installed successfully." -ForegroundColor Green
}

function Show-FilePicker {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Select a PDF to split"
    $dialog.Filter = "PDF Files (*.pdf)|*.pdf"
    $dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }
    return $null
}

function Split-Pdf {
    param([string]$SourcePath)

    Add-Type -Path $PdfSharpDll

    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
    $parentDir = [System.IO.Path]::GetDirectoryName($SourcePath)
    $outputDir = Join-Path $parentDir "${fileName}_pages"

    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

    try {
        $sourceDoc = [PdfSharp.Pdf.IO.PdfReader]::Open($SourcePath, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::Import)
    }
    catch {
        Write-Host "ERROR: Could not open the PDF. It may be password-protected or corrupted." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        pause
        exit 1
    }

    $totalPages = $sourceDoc.PageCount
    Write-Host "`nSplitting '$fileName.pdf' ($totalPages pages)..." -ForegroundColor Yellow
    Write-Host "Output: $outputDir`n" -ForegroundColor Gray

    $padWidth = $totalPages.ToString().Length

    for ($i = 0; $i -lt $totalPages; $i++) {
        $pageNum = ($i + 1).ToString().PadLeft($padWidth, '0')
        $outFile = Join-Path $outputDir "${fileName}_page${pageNum}.pdf"

        $newDoc = New-Object PdfSharp.Pdf.PdfDocument
        $newDoc.AddPage($sourceDoc.Pages[$i]) | Out-Null
        $newDoc.Save($outFile)
        $newDoc.Close()

        $pct = [math]::Round((($i + 1) / $totalPages) * 100)
        Write-Progress -Activity "Splitting PDF" -Status "Page $($i+1) of $totalPages" -PercentComplete $pct
    }

    $sourceDoc.Close()
    Write-Progress -Activity "Splitting PDF" -Completed

    Write-Host "Done! $totalPages pages saved to:" -ForegroundColor Green
    Write-Host "  $outputDir" -ForegroundColor Cyan

    # Open the output folder in Explorer
    Start-Process explorer.exe $outputDir
}

# --- Main ---
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PDF Page Splitter (free, no Acrobat)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# If no path provided, open a file picker
if (-not $PdfPath -or -not (Test-Path $PdfPath)) {
    if ($PdfPath -and -not (Test-Path $PdfPath)) {
        Write-Host "File not found: $PdfPath" -ForegroundColor Red
    }
    Write-Host "Opening file picker..." -ForegroundColor Gray
    $PdfPath = Show-FilePicker

    if (-not $PdfPath) {
        Write-Host "No file selected. Exiting." -ForegroundColor Yellow
        pause
        exit 0
    }
}

# Validate it's a PDF
if ([System.IO.Path]::GetExtension($PdfPath).ToLower() -ne ".pdf") {
    Write-Host "ERROR: Please select a PDF file." -ForegroundColor Red
    pause
    exit 1
}

Install-PdfSharp
Split-Pdf -SourcePath $PdfPath

Write-Host ""
pause
