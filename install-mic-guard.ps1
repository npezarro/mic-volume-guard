# install-mic-guard.ps1 — Adds mic-volume-guard to Windows startup (runs hidden)
# Run this once: powershell -ExecutionPolicy Bypass -File install-mic-guard.ps1

$scriptPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "mic-volume-guard.ps1"
$startupDir = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
$shortcutPath = "$startupDir\MicVolumeGuard.lnk"

if (!(Test-Path $scriptPath)) {
    Write-Host "ERROR: $scriptPath not found" -ForegroundColor Red
    exit 1
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = "powershell.exe"
$shortcut.Arguments = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
$shortcut.WorkingDirectory = Split-Path -Parent $scriptPath
$shortcut.Description = "Mic Volume Guard - keeps input at 100%"
$shortcut.Save()

Write-Host "Installed startup shortcut: $shortcutPath" -ForegroundColor Green
Write-Host "The mic volume guard will start automatically on next login."
Write-Host ""
Write-Host "To start it now:"
Write-Host "  Start-Process powershell -ArgumentList '-WindowStyle Hidden -ExecutionPolicy Bypass -File $scriptPath'" -ForegroundColor Cyan
Write-Host ""
Write-Host "To uninstall:"
Write-Host "  Remove-Item '$shortcutPath'" -ForegroundColor Yellow
