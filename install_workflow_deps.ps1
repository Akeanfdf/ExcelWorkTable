# Install workflow deps with a clean env (fixes PYTHONHOME pointing at project folder).
# Run from project root:  powershell -ExecutionPolicy Bypass -File .\install_workflow_deps.ps1

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

Remove-Item Env:PYTHONHOME -ErrorAction SilentlyContinue
Remove-Item Env:PYTHONPATH -ErrorAction SilentlyContinue

Write-Host "Using:" (Get-Command py -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source)
py -3.11 -c "import sys; print('executable:', sys.executable); print('prefix:', sys.prefix); print('version:', sys.version)"

py -3.11 -m pip install -r requirements-workflow.txt
Write-Host "Done."
