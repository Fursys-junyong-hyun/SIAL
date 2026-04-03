param(
    [string]$Python = "python",
    [string]$AppName = "유상사급 타처보관 확인서",
    [string]$Version = "0.1.0"
)

$ErrorActionPreference = "Stop"

Set-Location (Join-Path $PSScriptRoot "..")

& $Python -m pip install -U pyinstaller
& $Python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --name "유상사급타처보관" `
    --add-data "resources\\template.xlsx;resources" `
    --paths "src" `
    "src\\outsourced_inventory_confirmation\\__main__.py"

Write-Host "PyInstaller build complete. Output: dist\\유상사급타처보관"
Write-Host "Installer build can be run with Inno Setup using packaging\\installer.iss"
