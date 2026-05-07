# Interactive: official uninstall first, optional PATH trim, optional folder delete for D:\anaconda3
# Run: powershell -ExecutionPolicy Bypass -File .\uninstall_anaconda.ps1
# Requires: run in elevated PowerShell if you need to modify Machine PATH or delete protected files.

param(
    [string]$InstallRoot = 'D:\anaconda3'
)

$ErrorActionPreference = 'Stop'

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

Write-Host ""
Write-Host "=== Anaconda / Miniconda helper ===" -ForegroundColor Cyan
Write-Host "Target root: $InstallRoot"
Write-Host "Running as admin: $(Test-Admin)"
Write-Host ""

if (-not (Test-Path -LiteralPath $InstallRoot)) {
    Write-Host "Path does not exist (nothing to do): $InstallRoot" -ForegroundColor Yellow
    exit 0
}

$uninstallers = @(
    (Join-Path $InstallRoot 'Uninstall-Anaconda3.exe'),
    (Join-Path $InstallRoot 'Uninstall-Miniconda3.exe')
)
$foundUn = $uninstallers | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1

Write-Host "Choose:" -ForegroundColor Yellow
Write-Host "  [1] Launch official uninstaller (RECOMMENDED) if found"
Write-Host "  [2] Remove 'anaconda3' entries from USER Path only (after uninstall or if PATH is messy)"
Write-Host "  [3] DELETE entire folder $InstallRoot (DANGEROUS — only if uninstaller already ran or you accept risk)"
Write-Host "  [Q] Quit"
$choice = Read-Host "Enter 1 / 2 / 3 / Q"

if ($choice -eq '1') {
    if (-not $foundUn) {
        Write-Host "No Uninstall-Anaconda3.exe / Uninstall-Miniconda3.exe under $InstallRoot" -ForegroundColor Red
        Write-Host "Use Windows Settings -> Apps -> uninstall 'Anaconda' / 'Miniconda', then rerun this script for step 2 or 3."
        exit 1
    }
    Write-Host "Starting: $foundUn" -ForegroundColor Green
    Start-Process -FilePath $foundUn -Wait
    Write-Host "Uninstaller finished. Reboot if prompted. Then run option 2 to clean USER Path, and 3 only if folder remains."
    exit 0
}

if ($choice -eq '2') {
    $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
    if (-not $userPath) {
        Write-Host "User Path is empty." -ForegroundColor Yellow
        exit 0
    }
    $parts = $userPath -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    # Only strip paths that clearly belong to Anaconda/Miniconda installs (narrow on purpose).
    $bad = $parts | Where-Object { $_ -imatch 'anaconda3|miniconda3' }
    $keep = $parts | Where-Object { $_ -inotmatch 'anaconda3|miniconda3' }
    if (-not $bad) {
        Write-Host "No Anaconda/Miniconda-like segments found in USER Path." -ForegroundColor Green
        exit 0
    }
    Write-Host "Will REMOVE these USER Path entries:" -ForegroundColor Yellow
    $bad | ForEach-Object { Write-Host "  $_" }
    $confirm = Read-Host "Type YES to rewrite USER Path (only user scope)"
    if ($confirm -ne 'YES') { Write-Host "Aborted."; exit 0 }
    $newPath = ($keep -join ';').TrimEnd(';')
    [Environment]::SetEnvironmentVariable('Path', $newPath, 'User')
    Write-Host "USER Path updated. Open a NEW terminal for PATH to refresh." -ForegroundColor Green
    exit 0
}

if ($choice -eq '3') {
    Write-Host "This will try to permanently delete: $InstallRoot" -ForegroundColor Red
    Write-Host "Close VS Code, Jupyter, conda terminals, and any Python using this folder first."
    $confirm = Read-Host "Type exactly: DELETE $InstallRoot"
    $expected = "DELETE $InstallRoot"
    if ($confirm -ne $expected) {
        Write-Host "Confirmation mismatch. Aborted." -ForegroundColor Yellow
        exit 1
    }
    try {
        Remove-Item -LiteralPath $InstallRoot -Recurse -Force -ErrorAction Stop
        Write-Host "Removed: $InstallRoot" -ForegroundColor Green
    } catch {
        Write-Host "Remove failed (file in use or need admin): $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Try: close all apps, reboot, run this script as Administrator, or delete from Explorer."
        exit 1
    }
    Write-Host "Done. Run option 2 to clean USER Path if needed. Install python.org Python for this project." -ForegroundColor Green
    exit 0
}

Write-Host "Quit." -ForegroundColor Gray
exit 0
