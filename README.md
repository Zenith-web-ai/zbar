# zbar

Process blocker for Windows. Kills gaming processes every 5 minutes.

## Install

1. Plug in USB, open the `zbar` folder
2. Double-click `Install.bat` (will prompt for admin)
3. Done. See the post-install verification in the terminal.

Install location: `C:\ProgramData\zbar\`
Install report saved to USB: `install-report.txt`

## Check status

Double-click `Check.bat` for a full diagnostic. Report saved to USB: `zbar-report.txt`

## Uninstall

Double-click `Uninstall.bat` (will prompt for admin). Report saved to USB: `uninstall-report.txt`

## Update the blocklist

Edit `C:\ProgramData\zbar\blocklist.txt` on the target PC. One process name per line (without `.exe`). Lines starting with `#` are comments. Changes take effect within 5 minutes.

## How it works

- Tries 5 persistence methods (Task Scheduler, schtasks, Startup folder, Registry)
- Runs as a background PowerShell process with no window
- Scans every 5 minutes and force-kills blocked processes
- Kills logged to `C:\ProgramData\zbar\zbar-log.txt`
- Auto-restarts if the process dies
