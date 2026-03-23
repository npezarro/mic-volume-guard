# mic-volume-guard.ps1 - Background watchdog: keeps all recording devices at 100%
# Checks every 2 seconds. Runs silently, logs only when fixing.
# Respects intentional mute: if volume is set to exactly 0%, allows it for up to
# 120 seconds before restoring to 100%.
# Usage: powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File mic-volume-guard.ps1

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator {
    int EnumAudioEndpoints(int dataFlow, int stateMask, out IMMDeviceCollection devices);
    int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice device);
}

[Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceCollection {
    int GetCount(out int count);
    int Item(int index, out IMMDevice device);
}

[Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDevice {
    int Activate([MarshalAs(UnmanagedType.LPStruct)] Guid iid, int clsCtx, IntPtr activationParams, [MarshalAs(UnmanagedType.IUnknown)] out object endpoint);
    int OpenPropertyStore(int access, out IntPtr properties);
    int GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);
    int GetState(out int state);
}

[Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAudioEndpointVolume {
    int RegisterControlChangeNotify(IntPtr notify);
    int UnregisterControlChangeNotify(IntPtr notify);
    int GetChannelCount(out int count);
    int SetMasterVolumeLevel(float level, ref Guid eventContext);
    int SetMasterVolumeLevelScalar(float level, ref Guid eventContext);
    int GetMasterVolumeLevel(out float level);
    int GetMasterVolumeLevelScalar(out float level);
    int SetChannelVolumeLevel(int channel, float level, ref Guid eventContext);
    int SetChannelVolumeLevelScalar(int channel, float level, ref Guid eventContext);
    int GetChannelVolumeLevel(int channel, out float level);
    int GetChannelVolumeLevelScalar(int channel, out float level);
    int SetMute([MarshalAs(UnmanagedType.Bool)] bool mute, ref Guid eventContext);
    int GetMute([MarshalAs(UnmanagedType.Bool)] out bool mute);
}

public static class AudioHelper {
    // Returns parallel arrays: volumes[i] = scalar 0-1, muted[i] = mute flag
    public static void GetCaptureState(out float[] volumes, out bool[] muted) {
        volumes = new float[0];
        muted = new bool[0];
        try {
            var enumeratorType = Type.GetTypeFromCLSID(new Guid("BCDE0395-E52F-467C-8E3D-C4579291692E"));
            var enumerator = (IMMDeviceEnumerator)Activator.CreateInstance(enumeratorType);
            IMMDeviceCollection devices;
            enumerator.EnumAudioEndpoints(1, 0xB, out devices);
            int count;
            devices.GetCount(out count);
            var volumeGuid = new Guid("5CDF2C82-841E-4546-9722-0CF74078229A");
            volumes = new float[count];
            muted = new bool[count];
            for (int i = 0; i < count; i++) {
                try {
                    IMMDevice device;
                    devices.Item(i, out device);
                    object activated;
                    device.Activate(volumeGuid, 1, IntPtr.Zero, out activated);
                    if (activated == null) { volumes[i] = 1.0f; muted[i] = false; continue; }
                    var vol = (IAudioEndpointVolume)activated;
                    float current;
                    vol.GetMasterVolumeLevelScalar(out current);
                    volumes[i] = current;
                    bool isMuted;
                    vol.GetMute(out isMuted);
                    muted[i] = isMuted;
                } catch { volumes[i] = 1.0f; muted[i] = false; }
            }
        } catch {}
    }

    public static void SetCaptureVolume(int index, float level) {
        try {
            var enumeratorType = Type.GetTypeFromCLSID(new Guid("BCDE0395-E52F-467C-8E3D-C4579291692E"));
            var enumerator = (IMMDeviceEnumerator)Activator.CreateInstance(enumeratorType);
            IMMDeviceCollection devices;
            enumerator.EnumAudioEndpoints(1, 0xB, out devices);
            IMMDevice device;
            devices.Item(index, out device);
            var volumeGuid = new Guid("5CDF2C82-841E-4546-9722-0CF74078229A");
            var ctx = Guid.Empty;
            object activated;
            device.Activate(volumeGuid, 1, IntPtr.Zero, out activated);
            if (activated == null) return;
            var vol = (IAudioEndpointVolume)activated;
            vol.SetMasterVolumeLevelScalar(level, ref ctx);
        } catch {}
    }

    public static void SetCaptureUnmute(int index) {
        try {
            var enumeratorType = Type.GetTypeFromCLSID(new Guid("BCDE0395-E52F-467C-8E3D-C4579291692E"));
            var enumerator = (IMMDeviceEnumerator)Activator.CreateInstance(enumeratorType);
            IMMDeviceCollection devices;
            enumerator.EnumAudioEndpoints(1, 0xB, out devices);
            IMMDevice device;
            devices.Item(index, out device);
            var volumeGuid = new Guid("5CDF2C82-841E-4546-9722-0CF74078229A");
            var ctx = Guid.Empty;
            object activated;
            device.Activate(volumeGuid, 1, IntPtr.Zero, out activated);
            if (activated == null) return;
            var vol = (IAudioEndpointVolume)activated;
            vol.SetMute(false, ref ctx);
        } catch {}
    }
}
"@

$scriptPath = $MyInvocation.MyCommand.Path
$scriptDir = Split-Path -Parent $scriptPath
$logFile = Join-Path $scriptDir "mic-volume-guard.log"
$muteGraceSec = 120
$maxStaleSeconds = 300  # restart self if no successful poll for 5 minutes

# Kill any other mic-volume-guard instances (prevent stacking)
$myPid = $PID
Get-WmiObject Win32_Process -Filter "Name='powershell.exe' AND CommandLine LIKE '%mic-volume-guard%'" -ErrorAction SilentlyContinue |
    Where-Object { $_.ProcessId -ne $myPid } |
    ForEach-Object {
        $_.Terminate() | Out-Null
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path $logFile -Value "$ts  Killed stale guard instance (PID $($_.ProcessId))"
    }

# Track per-device mute timestamps: key = device index, value = DateTime when mute was first seen
# Covers both hardware mute (mute flag) and volume-at-0% mute
$muteTimers = @{}
# Track per-device previous state for change detection
$prevVolumes = @{}
$prevMuted = @{}

$consecutiveErrors = 0
$lastPollTime = Get-Date
$heartbeatFile = Join-Path $scriptDir "mic-volume-guard.heartbeat"

while ($true) {
    # Detect sleep/wake: if more than 30 seconds passed since last poll,
    # the system was likely asleep. Reset state so we re-read fresh.
    $now = Get-Date
    $gap = ($now - $lastPollTime).TotalSeconds
    if ($gap -gt 30) {
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path $logFile -Value "$ts  Wake detected (${gap}s gap) - resetting device state"
        $prevVolumes = @{}
        $prevMuted = @{}
        $muteTimers = @{}
    }
    $lastPollTime = $now

    try {
        $volumes = $null
        $muted = $null
        [AudioHelper]::GetCaptureState([ref]$volumes, [ref]$muted)

        if ($null -eq $volumes -or $volumes.Length -eq 0) {
            # COM may be stale after sleep/wake - wait and retry
            $consecutiveErrors++
            if ($consecutiveErrors -eq 1 -or $consecutiveErrors % 30 -eq 0) {
                $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Add-Content -Path $logFile -Value "$ts  No devices returned (attempt $consecutiveErrors) - COM may be stale, retrying..."
            }
            Start-Sleep -Seconds 2
            continue
        }

        if ($consecutiveErrors -gt 0) {
            $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Add-Content -Path $logFile -Value "$ts  Recovered after $consecutiveErrors failed attempts"
            $consecutiveErrors = 0
            $prevVolumes = @{}
            $prevMuted = @{}
            $muteTimers = @{}
        }

        for ($i = 0; $i -lt $volumes.Length; $i++) {
            $vol = $volumes[$i]
            $pct = [int]($vol * 100)
            $isMuted = $muted[$i]
            $prevPct = if ($prevVolumes.ContainsKey($i)) { $prevVolumes[$i] } else { -1 }
            $wasMuted = if ($prevMuted.ContainsKey($i)) { $prevMuted[$i] } else { $false }

            # Log volume changes
            if ($pct -ne $prevPct) {
                $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                if ($prevPct -eq -1) {
                    Add-Content -Path $logFile -Value "$ts  Device $i initial volume: ${pct}% muted: $isMuted"
                } else {
                    Add-Content -Path $logFile -Value "$ts  Device $i volume changed: ${prevPct}% -> ${pct}% muted: $isMuted"
                }
                $prevVolumes[$i] = $pct
            }

            # Log mute state changes
            if ($isMuted -ne $wasMuted -and $prevPct -ne -1) {
                $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                if ($isMuted) {
                    Add-Content -Path $logFile -Value "$ts  Device $i mute flag set (hardware mute)"
                } else {
                    Add-Content -Path $logFile -Value "$ts  Device $i mute flag cleared (hardware unmute)"
                }
            }
            $prevMuted[$i] = $isMuted

            # Check if device is muted (either via mute flag or volume at 0%)
            $effectivelyMuted = $isMuted -or ($vol -le 0.01)

            if ($effectivelyMuted) {
                # Device is muted; respect with grace period
                if (-not $muteTimers.ContainsKey($i)) {
                    $muteTimers[$i] = Get-Date
                    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    $reason = if ($isMuted) { "mute flag" } else { "0% volume" }
                    Add-Content -Path $logFile -Value "$ts  Device $i muted ($reason), grace period started (${muteGraceSec}s)"
                    continue
                }

                $elapsed = ((Get-Date) - $muteTimers[$i]).TotalSeconds
                if ($elapsed -lt $muteGraceSec) {
                    continue
                }

                # Grace period expired, restore
                if ($isMuted) {
                    [AudioHelper]::SetCaptureUnmute($i)
                }
                [AudioHelper]::SetCaptureVolume($i, 1.0)
                $muteTimers.Remove($i)
                $prevVolumes[$i] = 100
                $prevMuted[$i] = $false
                $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Add-Content -Path $logFile -Value "$ts  Device $i mute grace expired (${muteGraceSec}s), restored to 100%"
                continue
            }

            # Not muted; clear any mute timer
            if ($muteTimers.ContainsKey($i)) {
                $muteTimers.Remove($i)
            }

            if ($vol -ge 0.99) {
                # Volume is fine
                continue
            }

            # Volume drifted to something other than 100% - fix immediately
            [AudioHelper]::SetCaptureVolume($i, 1.0)
            $prevVolumes[$i] = 100
            $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Add-Content -Path $logFile -Value "$ts  Device $i was at ${pct}%, restored to 100%"
        }
    } catch {
        $consecutiveErrors++
        if ($consecutiveErrors -eq 1 -or $consecutiveErrors % 30 -eq 0) {
            $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Add-Content -Path $logFile -Value "$ts  ERROR (attempt $consecutiveErrors): $($_.Exception.Message)"
        }
    }

    # Write heartbeat so external watchdog can detect frozen process
    [IO.File]::WriteAllText($heartbeatFile, (Get-Date).ToString("o"))

    # Self-restart if stale for too long (COM objects gone bad after sleep/wake)
    $staleSec = $consecutiveErrors * 2  # each error = ~2 seconds
    if ($staleSec -ge $maxStaleSeconds) {
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path $logFile -Value "$ts  Stale for ${staleSec}s ($consecutiveErrors errors), restarting self..."
        Start-Process powershell -ArgumentList "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" -WindowStyle Hidden
        Start-Sleep -Seconds 2
        exit 0
    }

    Start-Sleep -Seconds 2
}
