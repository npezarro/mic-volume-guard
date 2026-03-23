# context.md
Last Updated: 2026-03-22 | Initial setup, guard working with error recovery

## Current State
- Guard script (`mic-volume-guard.ps1`) is running and actively correcting volume drift
- Startup shortcut installed at `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\MicVolumeGuard.lnk`
- Root cause identified: Krisp noise cancellation continuously resets HyperX mic to ~96%
- Guard recovers from sleep/wake via error handling and wake detection (time gap >30s)
- Mute grace period (120s) tested and working: 0% volume is held for 120s before auto-restore
- All scripts pushed to https://github.com/npezarro/mic-volume-guard

## Known Issues
- Guard process can silently die after extended sleep/wake cycles (COM objects go stale)
- Error recovery handles this by catching exceptions and retrying, but if the process is fully dead, only a reboot or manual restart recovers it
- No scheduled task for wake-recovery (requires admin elevation)
- Krisp fights the guard every 2s on the HyperX mic (96% vs 100%); this is cosmetic and harmless

## Scripts
- `mic-volume-guard.ps1` - main watchdog, polls every 2s, error recovery, wake detection, 120s mute grace
- `setup.ps1` - one-time installer: startup shortcut, kills stale instances, starts guard
- `fix-mic-volume.ps1` - one-shot volume fix to 100%
- `debug-default-mic.ps1` - shows default capture device roles with volume/mute status
- `diag-mic.ps1` - full inventory of all capture devices in all states
- `root-cause-check.ps1` - exists in Documents/Claude copy but not in this repo yet
- `install-mic-guard.ps1` - alternative installer

## Environment
- Windows 11 Pro
- Scripts live on I: drive at `I:\Stuff\Projects\2026\mic-volume-guard\mic-volume-guard\`
- Startup shortcut points to I: drive path
- Audio devices: HyperX Cloud Alpha Wireless (primary), NVIDIA Broadcast, Krisp Audio
- Realtek HD Audio driver v6.0.9768.1
- Krisp v3.11.4 installed at `%LOCALAPPDATA%\Programs\Krisp\`
- GitHub CLI installed, authenticated as npezarro

## Active Branch
main
