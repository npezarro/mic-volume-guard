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
$watchdogScript = Join-Path $scriptDir "mic-volume-guard-watchdog.ps1"
$shortcut.Arguments        = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$watchdogScript`""
$shortcut.WorkingDirectory = $scriptDir
$shortcut.Description      = "Mic Volume Guard - keeps recording input at 100%"
$shortcut.Save()
Write-Host "  Shortcut: $shortcutPath" -ForegroundColor Green

# ── Kill ALL existing guard/watchdog processes ─
Write-Host "Killing any existing guard/watchdog instances..." -ForegroundColor Cyan
$killed = 0
$myPid = $PID
Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -like '*mic-volume-guard*' -and $_.ProcessId -ne $myPid -and $_.CommandLine -notlike '*setup*' } |
    ForEach-Object {
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
        Write-Host "  Killed PID $($_.ProcessId)" -ForegroundColor Yellow
        $killed++
    }
if ($killed -eq 0) {
    Write-Host "  No existing instances found." -ForegroundColor Green
}
Start-Sleep -Seconds 2

Write-Host "Starting mic-volume-guard watchdog..." -ForegroundColor Cyan

Start-Process powershell -ArgumentList "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$watchdogScript`"" -WindowStyle Hidden
Write-Host "  Watchdog + guard running in background." -ForegroundColor Green

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
