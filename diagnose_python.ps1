# Diagnose why Python prefix points at project folder (run from project root).
# Usage: powershell -ExecutionPolicy Bypass -File .\diagnose_python.ps1

$ErrorActionPreference = 'Continue'
Write-Host "=== Relevant environment variables ===" -ForegroundColor Cyan
foreach ($n in @('VIRTUAL_ENV', 'CONDA_PREFIX', 'CONDA_DEFAULT_ENV', 'CONDA_PYTHON_EXE', 'PYTHONHOME', 'PYTHONPATH')) {
    $v = [Environment]::GetEnvironmentVariable($n, 'Process')
    if (-not $v) { $v = [Environment]::GetEnvironmentVariable($n, 'User') }
    if (-not $v) { $v = [Environment]::GetEnvironmentVariable($n, 'Machine') }
    Write-Host ("  {0}={1}" -f $n, $(if ($v) { $v } else { '(empty)' }))
}

$condaPy = 'D:\anaconda3\python.exe'
if (-not (Test-Path $condaPy)) {
    Write-Host "Not found: $condaPy (edit script if your Anaconda path differs)" -ForegroundColor Yellow
    exit 1
}

Write-Host "`n=== Run Anaconda python from SYSTEMDRIVE (ignore cwd) ===" -ForegroundColor Cyan
Push-Location $env:SystemDrive
try {
    & $condaPy -c "import sys; print('prefix=', sys.prefix); print('base_prefix=', sys.base_prefix); print('executable=', sys.executable)"
} finally {
    Pop-Location
}

Write-Host "`n=== Run same from current project directory ===" -ForegroundColor Cyan
& $condaPy -c "import sys; print('prefix=', sys.prefix); print('base_prefix=', sys.base_prefix)"

Write-Host "`n=== Files to check manually (open in Notepad if they exist) ===" -ForegroundColor Cyan
$check = @(
    'D:\anaconda3\pyvenv.cfg',
    'D:\anaconda3\python311._pth',
    'D:\anaconda3\python._pth',
    (Join-Path $PSScriptRoot 'pyvenv.cfg')
)
foreach ($p in $check) {
    if (Test-Path $p) {
        Write-Host "  EXISTS: $p" -ForegroundColor Yellow
        if ($p -match 'pyvenv\.cfg$|_pth$') {
            Write-Host "    ---- first lines ----" -ForegroundColor DarkGray
            Get-Content -LiteralPath $p -TotalCount 25 -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "    $_" }
        }
    } else {
        Write-Host "  (no) $p"
    }
}

Write-Host "`n=== Windows Registry: Python install path (wrong path here can confuse tools) ===" -ForegroundColor Cyan
$regPaths = @(
    'HKLM:\SOFTWARE\Python\PythonCore\3.11\InstallPath',
    'HKLM:\SOFTWARE\WOW6432Node\Python\PythonCore\3.11\InstallPath',
    'HKCU:\SOFTWARE\Python\PythonCore\3.11\InstallPath'
)
foreach ($rp in $regPaths) {
    if (Test-Path $rp) {
        $d = (Get-ItemProperty -LiteralPath $rp -ErrorAction SilentlyContinue).'(default)'
        Write-Host "  $rp -> $d" -ForegroundColor Yellow
    } else {
        Write-Host "  (no) $rp"
    }
}

$proj = $PSScriptRoot
Write-Host "`n=== Suspicious files under project (wrong copies can confuse debugging) ===" -ForegroundColor Cyan
foreach ($rel in @('Lib\encodings', 'Lib', 'python311.dll', 'python.exe', 'pyvenv.cfg', 'Scripts\python.exe')) {
    $fp = Join-Path $proj $rel
    if (Test-Path $fp) { Write-Host "  EXISTS: $fp" -ForegroundColor Yellow } else { Write-Host "  (no) $rel" }
}

Write-Host "`nIf prefix is correct from C:\ but wrong from project folder, look for pyvenv.cfg / ._pth under the project or Anaconda root." -ForegroundColor Green
Write-Host "If prefix is always wrong: repair/reinstall Anaconda, or install python.org Python to a non-C path (e.g. D:\Python311, E:\Python311) and use that exe for pip." -ForegroundColor Green
