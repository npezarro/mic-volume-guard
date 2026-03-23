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

public static class AudioDiag {
    public static void DumpAll() {
        var enumeratorType = Type.GetTypeFromCLSID(new Guid("BCDE0395-E52F-467C-8E3D-C4579291692E"));
        var enumerator = (IMMDeviceEnumerator)Activator.CreateInstance(enumeratorType);
        var volumeGuid = new Guid("5CDF2C82-841E-4546-9722-0CF74078229A");

        // Check all state masks individually
        string[] stateNames = { "ACTIVE", "DISABLED", "NOT_PRESENT", "UNPLUGGED" };
        int[] stateMasks = { 1, 2, 4, 8 };

        for (int s = 0; s < 4; s++) {
            IMMDeviceCollection devices;
            enumerator.EnumAudioEndpoints(1, stateMasks[s], out devices);
            int count;
            devices.GetCount(out count);
            if (count == 0) continue;

            Console.WriteLine("=== " + stateNames[s] + " (" + count + " devices) ===");
            for (int i = 0; i < count; i++) {
                try {
                    IMMDevice device;
                    devices.Item(i, out device);
                    string id;
                    device.GetId(out id);
                    int state;
                    device.GetState(out state);

                    Console.Write("[" + i + "] " + id + " (state=" + state + ")");

                    try {
                        object activated;
                        device.Activate(volumeGuid, 1, IntPtr.Zero, out activated);
                        var vol = (IAudioEndpointVolume)activated;
                        float scalar;
                        vol.GetMasterVolumeLevelScalar(out scalar);
                        bool muted;
                        vol.GetMute(out muted);
                        Console.WriteLine("  Vol=" + (int)(scalar * 100) + "%  Muted=" + muted);
                    } catch (Exception e) {
                        Console.WriteLine("  (cannot read volume: " + e.Message + ")");
                    }
                } catch (Exception e) {
                    Console.WriteLine("[" + i + "] ERROR: " + e.Message);
                }
            }
            Console.WriteLine();
        }
    }
}
"@

[AudioDiag]::DumpAll()
