param(
    [string]$Python = "python"
)

$ErrorActionPreference = "Stop"

Set-Location (Join-Path $PSScriptRoot "..")

& $Python -m pip install -U pyinstaller
& $Python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --name "outsourced_inventory_confirmation" `
    --add-data "resources\\template.xlsx;resources" `
    --paths "src" `
    "src\\outsourced_inventory_confirmation\\__main__.py"

Write-Host "PyInstaller build complete. Output: dist\\outsourced_inventory_confirmation"
Write-Host "Installer build requires Inno Setup with packaging\\installer.iss"
