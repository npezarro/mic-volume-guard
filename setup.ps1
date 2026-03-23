# setup.ps1 — One-time setup for mic-volume-guard
# Sets execution policy, unblocks scripts, installs startup shortcut,
# and starts the guard. Runs from wherever the repo is cloned.
#
# Usage: powershell -ExecutionPolicy Bypass -File setup.ps1

$ErrorActionPreference = "Stop"

# ── Resolve script directory (works from any location) ───────────────
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$guardScript = Join-Path $scriptDir "mic-volume-guard.ps1"

if (!(Test-Path $guardScript)) {
    Write-Host "ERROR: mic-volume-guard.ps1 not found in $scriptDir" -ForegroundColor Red
    exit 1
}

Write-Host "Mic Volume Guard Setup" -ForegroundColor Cyan
Write-Host "  Location: $scriptDir" -ForegroundColor Gray
Write-Host ""

# ── Set execution policy for current user ────────────────────────────
Write-Host "Checking execution policy..." -ForegroundColor Cyan
$currentPolicy = Get-ExecutionPolicy
if ($currentPolicy -eq "Restricted") {
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction SilentlyContinue
    Write-Host "  Set to RemoteSigned." -ForegroundColor Green
} else {
    Write-Host "  Already set to $currentPolicy (OK)." -ForegroundColor Green
}

# ── Unblock all scripts (removes "downloaded from internet" flag) ────
Write-Host "Unblocking scripts..." -ForegroundColor Cyan
Get-ChildItem -Path $scriptDir -Filter "*.ps1" | Unblock-File
Write-Host "  Done." -ForegroundColor Green

# ── Install startup shortcut ─────────────────────────────────────────
Write-Host "Installing startup shortcut..." -ForegroundColor Cyan
$startupDir   = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
$shortcutPath = "$startupDir\MicVolumeGuard.lnk"

$shell    = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath       = "powershell.exe"
$shortcut.Arguments        = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$guardScript`""
$shortcut.WorkingDirectory = $scriptDir
$shortcut.Description      = "Mic Volume Guard - keeps recording input at 100%"
$shortcut.Save()
Write-Host "  Shortcut: $shortcutPath" -ForegroundColor Green

# ── Kill ALL existing guard processes (WMI for reliable CommandLine access) ─
Write-Host "Killing any existing guard instances..." -ForegroundColor Cyan
$killed = 0
Get-WmiObject Win32_Process -Filter "Name='powershell.exe' AND CommandLine LIKE '%mic-volume-guard%'" -ErrorAction SilentlyContinue |
    ForEach-Object {
        $_.Terminate() | Out-Null
        Write-Host "  Killed PID $($_.ProcessId)" -ForegroundColor Yellow
        $killed++
    }
if ($killed -eq 0) {
    Write-Host "  No existing instances found." -ForegroundColor Green
}

Write-Host "Starting mic-volume-guard..." -ForegroundColor Cyan

Start-Process powershell -ArgumentList "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$guardScript`"" -WindowStyle Hidden
Write-Host "  Guard is running in background." -ForegroundColor Green

# ── Summary ──────────────────────────────────────────────────────────
Write-Host ""
Write-Host "Setup complete." -ForegroundColor Green
Write-Host "  Scripts:  $scriptDir"
Write-Host "  Startup:  $shortcutPath"
Write-Host "  Log:      $scriptDir\mic-volume-guard.log"
Write-Host ""
Write-Host "To verify it's running:" -ForegroundColor Cyan
Write-Host "  Get-Process powershell | Where-Object { `$_.CommandLine -like '*mic-volume*' }"
Write-Host ""
Write-Host "To uninstall:" -ForegroundColor Yellow
Write-Host "  Remove-Item '$shortcutPath'"
Write-Host "  Get-Process powershell | Where-Object { `$_.CommandLine -like '*mic-volume*' } | Stop-Process"
