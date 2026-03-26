# zbar - Process blocker
# Checks every 5 minutes, kills processes matching blocklist

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$blocklistFile = Join-Path $scriptDir "blocklist.txt"
$logFile = Join-Path $scriptDir "zbar-log.txt"
$pidFile = Join-Path $scriptDir "zbar.pid"

# Write PID for diagnostics (ASCII so batch can read it)
[IO.File]::WriteAllText($pidFile, $PID.ToString())

while ($true) {
    # Reload blocklist each cycle so edits take effect without restart
    $blocklist = @(Get-Content $blocklistFile -ErrorAction SilentlyContinue |
        ForEach-Object { $_.Trim().ToLower() } |
        Where-Object { $_ -and $_ -notmatch '^\s*#' })

    if ($blocklist.Count -gt 0) {
        foreach ($proc in (Get-Process -ErrorAction SilentlyContinue)) {
            if ($blocklist -contains $proc.ProcessName.ToLower()) {
                $now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                try {
                    Stop-Process -Id $proc.Id -Force -ErrorAction Stop
                    "$now  KILLED  $($proc.ProcessName)  (PID $($proc.Id))" |
                        Out-File $logFile -Append -Encoding utf8
                } catch {
                    "$now  FAILED  $($proc.ProcessName)  (PID $($proc.Id))  $($_.Exception.Message)" |
                        Out-File $logFile -Append -Encoding utf8
                }
            }
        }
    }

    Start-Sleep -Seconds 300
}
