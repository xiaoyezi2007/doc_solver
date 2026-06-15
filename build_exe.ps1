$ErrorActionPreference = "Stop"

python -m pip install -r requirements.txt

if (Test-Path .\build) {
    Remove-Item .\build -Recurse -Force
}

python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name DocSolver `
    --hidden-import comtypes.stream `
    --hidden-import comtypes.client `
    .\app.py

New-Item -ItemType Directory -Force -Path .\dist\input, .\dist\output | Out-Null
Write-Host "Build complete: dist\DocSolver.exe"
