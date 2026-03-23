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
    public static float[] GetCaptureVolumes() {
        try {
            var enumeratorType = Type.GetTypeFromCLSID(new Guid("BCDE0395-E52F-467C-8E3D-C4579291692E"));
            var enumerator = (IMMDeviceEnumerator)Activator.CreateInstance(enumeratorType);
            IMMDeviceCollection devices;
            // stateMask 0xB = ACTIVE|DISABLED|UNPLUGGED (skip NOT_PRESENT)
            enumerator.EnumAudioEndpoints(1, 0xB, out devices);
            int count;
            devices.GetCount(out count);
            var volumeGuid = new Guid("5CDF2C82-841E-4546-9722-0CF74078229A");
            float[] volumes = new float[count];
            for (int i = 0; i < count; i++) {
                try {
                    IMMDevice device;
                    devices.Item(i, out device);
                    object activated;
                    device.Activate(volumeGuid, 1, IntPtr.Zero, out activated);
                    if (activated == null) { volumes[i] = 1.0f; continue; }
                    var vol = (IAudioEndpointVolume)activated;
                    float current;
                    vol.GetMasterVolumeLevelScalar(out current);
                    volumes[i] = current;
                } catch { volumes[i] = 1.0f; }
            }
            return volumes;
        } catch { return new float[0]; }
    }

    public static void SetCaptureVolume(int index, float level) {
        try {
            var enumeratorType = Type.GetTypeFromCLSID(new Guid("BCDE0395-E52F-467C-8E3D-C4579291692E"));
            var enumerator = (IMMDeviceEnumerator)Activator.CreateInstance(enumeratorType);
            IMMDeviceCollection devices;
            // stateMask 0xB = ACTIVE|DISABLED|UNPLUGGED (skip NOT_PRESENT)
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
}
"@

$logFile = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "mic-volume-guard.log"
$muteGraceSec = 120

# Kill any other mic-volume-guard instances (prevent stacking)
$myPid = $PID
Get-WmiObject Win32_Process -Filter "Name='powershell.exe' AND CommandLine LIKE '%mic-volume-guard%'" -ErrorAction SilentlyContinue |
    Where-Object { $_.ProcessId -ne $myPid } |
    ForEach-Object {
        $_.Terminate() | Out-Null
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path $logFile -Value "$ts  Killed stale guard instance (PID $($_.ProcessId))"
    }

# Track per-device mute timestamps: key = device index, value = DateTime when 0% was first seen
$muteTimers = @{}
# Track per-device previous volume for change detection
$prevVolumes = @{}

$consecutiveErrors = 0
$lastPollTime = Get-Date

while ($true) {
    # Detect sleep/wake: if more than 30 seconds passed since last poll,
    # the system was likely asleep. Reset state so we re-read fresh.
    $now = Get-Date
    $gap = ($now - $lastPollTime).TotalSeconds
    if ($gap -gt 30) {
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Add-Content -Path $logFile -Value "$ts  Wake detected (${gap}s gap) - resetting device state"
        $prevVolumes = @{}
        $muteTimers = @{}
    }
    $lastPollTime = $now

    try {
        $volumes = [AudioHelper]::GetCaptureVolumes()

        if ($volumes.Length -eq 0) {
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
            # Reset state after recovery - device indices may have changed
            $prevVolumes = @{}
            $muteTimers = @{}
        }

        for ($i = 0; $i -lt $volumes.Length; $i++) {
            $vol = $volumes[$i]
            $pct = [int]($vol * 100)
            $prevPct = if ($prevVolumes.ContainsKey($i)) { $prevVolumes[$i] } else { -1 }

            # Log any volume change
            if ($pct -ne $prevPct) {
                $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                if ($prevPct -eq -1) {
                    Add-Content -Path $logFile -Value "$ts  Device $i initial volume: ${pct}%"
                } else {
                    Add-Content -Path $logFile -Value "$ts  Device $i volume changed: ${prevPct}% -> ${pct}%"
                }
                $prevVolumes[$i] = $pct
            }

            if ($vol -ge 0.99) {
                # Volume is fine, clear any mute timer
                if ($muteTimers.ContainsKey($i)) {
                    $muteTimers.Remove($i)
                }
                continue
            }

            if ($vol -le 0.01) {
                # Volume is at 0% - intentional mute (headset button)
                if (-not $muteTimers.ContainsKey($i)) {
                    # First time seeing 0%, start grace period
                    $muteTimers[$i] = Get-Date
                    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    Add-Content -Path $logFile -Value "$ts  Device $i muted (0%), grace period started (${muteGraceSec}s)"
                    continue
                }

                $elapsed = ((Get-Date) - $muteTimers[$i]).TotalSeconds
                if ($elapsed -lt $muteGraceSec) {
                    # Still within grace period, leave it muted
                    continue
                }

                # Grace period expired, restore to 100%
                [AudioHelper]::SetCaptureVolume($i, 1.0)
                $muteTimers.Remove($i)
                $prevVolumes[$i] = 100
                $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Add-Content -Path $logFile -Value "$ts  Device $i mute grace expired (${muteGraceSec}s), restored to 100%"
            } else {
                # Volume drifted to something other than 0% or 100% - fix immediately
                [AudioHelper]::SetCaptureVolume($i, 1.0)
                if ($muteTimers.ContainsKey($i)) {
                    $muteTimers.Remove($i)
                }
                $prevVolumes[$i] = 100
                $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Add-Content -Path $logFile -Value "$ts  Device $i was at ${pct}%, restored to 100%"
            }
        }
    } catch {
        $consecutiveErrors++
        if ($consecutiveErrors -eq 1 -or $consecutiveErrors % 30 -eq 0) {
            $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Add-Content -Path $logFile -Value "$ts  ERROR (attempt $consecutiveErrors): $($_.Exception.Message)"
        }
    }

    Start-Sleep -Seconds 2
}
