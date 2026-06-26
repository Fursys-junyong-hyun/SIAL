param(
    [string]$Python = "python",
    [switch]$SkipInstaller
)

$ErrorActionPreference = "Stop"

Set-Location (Join-Path $PSScriptRoot "..")

# PyInstaller, ISCC 같은 네이티브 도구는 진행 로그를 stderr 로 흘리는데, PowerShell 의
# $ErrorActionPreference="Stop" 가 그걸 오류로 오인해 스크립트를 멈춘다.
# 네이티브 명령을 호출할 때는 직접 종료 코드만 확인하도록 헬퍼를 둔다.
function Invoke-Native {
    param([Parameter(Mandatory)] [string]$Command, [Parameter(ValueFromRemainingArguments)] $Args)
    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        & $Command @Args
        $code = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $prev
    }
    if ($code -ne 0) {
        throw "'$Command' exited with code $code"
    }
}

Write-Host "[1/4] Installing build dependencies..."
Invoke-Native $Python -m pip install -U pyinstaller reportlab pillow

Write-Host "[2/4] Generating user manual PDF..."
Invoke-Native $Python scripts\generate_user_manual_pdf.py

Write-Host "[3/4] Running PyInstaller..."
Invoke-Native $Python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --name "outsourced_inventory_confirmation" `
    --add-data "resources\template.xlsx;resources" `
    --paths "src" `
    "app_bootstrap.py"

Write-Host "PyInstaller build complete. Output: dist\outsourced_inventory_confirmation"

if ($SkipInstaller) {
    Write-Host "Skipping Inno Setup step (-SkipInstaller specified)."
    return
}

$isccCandidates = @(
    "${env:ProgramFiles}\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles}\Inno Setup 5\ISCC.exe",
    "${env:ProgramFiles(x86)}\Inno Setup 5\ISCC.exe"
)
$iscc = $isccCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if ($iscc) {
    Write-Host "[4/4] Building installer with Inno Setup: $iscc"
    Invoke-Native $iscc "packaging\installer.iss"
    Write-Host "Installer build complete. Output: dist_installer\outsourced_inventory_confirmation_setup.exe"
} else {
    Write-Host "[4/4] Skipped: Inno Setup (ISCC.exe) not found. Open packaging\installer.iss in the Inno Setup IDE to build manually."
}
