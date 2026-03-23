Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator {
    int EnumAudioEndpoints(int dataFlow, int stateMask, out IntPtr devices);
    int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice device);
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

public static class DefaultMicCheck {
    // role: 0=Console, 1=Multimedia, 2=Communications
    public static void Check() {
        var enumeratorType = Type.GetTypeFromCLSID(new Guid("BCDE0395-E52F-467C-8E3D-C4579291692E"));
        var enumerator = (IMMDeviceEnumerator)Activator.CreateInstance(enumeratorType);
        var volumeGuid = new Guid("5CDF2C82-841E-4546-9722-0CF74078229A");
        string[] roleNames = {"Console", "Multimedia", "Communications"};

        for (int role = 0; role < 3; role++) {
            try {
                IMMDevice device;
                // dataFlow=1 = eCapture
                int hr = enumerator.GetDefaultAudioEndpoint(1, role, out device);
                if (hr != 0 || device == null) {
                    Console.WriteLine(roleNames[role] + " default capture: NONE (hr=0x" + hr.ToString("X") + ")");
                    continue;
                }
                string id;
                device.GetId(out id);
                int state;
                device.GetState(out state);

                object activated;
                device.Activate(volumeGuid, 1, IntPtr.Zero, out activated);
                if (activated == null) {
                    Console.WriteLine(roleNames[role] + " default capture: " + id + " state=" + state + " (null activation)");
                    continue;
                }
                var vol = (IAudioEndpointVolume)activated;
                float scalar;
                vol.GetMasterVolumeLevelScalar(out scalar);
                bool muted;
                vol.GetMute(out muted);
                Console.WriteLine(roleNames[role] + " default capture: " + (int)(scalar*100) + "% muted=" + muted + " state=" + state + " " + id);
            } catch (Exception e) {
                Console.WriteLine(roleNames[role] + " default capture: ERROR " + e.Message);
            }
        }
    }
}
"@

Write-Host "=== Default capture devices ===" -ForegroundColor Cyan
[DefaultMicCheck]::Check()

Write-Host "`n=== PowerShell audio info ===" -ForegroundColor Cyan
# Check what Windows Settings would show
$wshell = New-Object -ComObject Shell.Application
# Also try the newer AudioDeviceCmdlets if available
try {
    Import-Module AudioDeviceCmdlets -ErrorAction Stop
    Get-AudioDevice -Recording | Format-Table Name, ID, Default, Volume -AutoSize
} catch {
    Write-Host "(AudioDeviceCmdlets not installed)" -ForegroundColor Gray
}
