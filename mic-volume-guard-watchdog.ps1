# mic-volume-guard-watchdog.ps1 - Launcher that monitors the guard via heartbeat
# If the guard stops updating its heartbeat file for 60 seconds, kill and restart it.
# Usage: powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File mic-volume-guard-watchdog.ps1

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$guardScript = Join-Path $scriptDir "mic-volume-guard.ps1"
$heartbeatFile = Join-Path $scriptDir "mic-volume-guard.heartbeat"
$logFile = Join-Path $scriptDir "mic-volume-guard.log"
$maxHeartbeatAge = 60  # seconds before considering guard dead

function Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFile -Value "$ts  [watchdog] $msg"
}

function Kill-Guard {
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -EA SilentlyContinue |
        Where-Object { $_.CommandLine -like '*mic-volume-guard.ps1*' -and $_.CommandLine -notlike '*watchdog*' } |
        ForEach-Object {
            Stop-Process -Id $_.ProcessId -Force -EA SilentlyContinue
            Log "Killed guard PID $($_.ProcessId)"
        }
}

function Start-Guard {
    Start-Process powershell -ArgumentList "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$guardScript`"" -WindowStyle Hidden
    Log "Started new guard process"
}

# Kill any existing guard and start fresh
Kill-Guard
Start-Sleep -Seconds 2
Start-Guard

Log "Watchdog started, monitoring heartbeat every 30s (max age: ${maxHeartbeatAge}s)"

while ($true) {
    Start-Sleep -Seconds 30

    $guardAlive = $false

    if (Test-Path $heartbeatFile) {
        try {
            $lastBeat = [DateTime]::Parse([IO.File]::ReadAllText($heartbeatFile))
            $age = ((Get-Date) - $lastBeat).TotalSeconds
            if ($age -lt $maxHeartbeatAge) {
                $guardAlive = $true
            } else {
                Log "Heartbeat stale (${age}s old), restarting guard"
            }
        } catch {
            Log "Cannot parse heartbeat file, restarting guard"
        }
    } else {
        # No heartbeat file yet; give it 60s to create one
        $guardProcs = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -EA SilentlyContinue |
            Where-Object { $_.CommandLine -like '*mic-volume-guard.ps1*' -and $_.CommandLine -notlike '*watchdog*' }
        if ($guardProcs) {
            $guardAlive = $true  # guard is running, just hasn't written heartbeat yet
        } else {
            Log "No heartbeat file and no guard process, starting guard"
        }
    }

    if (-not $guardAlive) {
        Kill-Guard
        Start-Sleep -Seconds 2
        Start-Guard
    }
}
