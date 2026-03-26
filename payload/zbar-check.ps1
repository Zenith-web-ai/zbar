# zbar - Comprehensive Diagnostic Check
# Writes full report to USB and console with colored output

param(
    [string]$UsbDir = ""
)

$installDir = "C:\ProgramData\zbar"

# --- Detect USB directory if not provided ---
if (-not $UsbDir -or -not (Test-Path $UsbDir)) {
    # Check if $PSScriptRoot parent has Install.bat
    if ($PSScriptRoot) {
        $parentDir = Split-Path $PSScriptRoot -Parent -ErrorAction SilentlyContinue
        if ($parentDir -and (Test-Path (Join-Path $parentDir "Install.bat"))) {
            $UsbDir = $parentDir
        } else {
            # Fallback: write report next to this script
            $UsbDir = $PSScriptRoot
        }
    }
}

# --- Report infrastructure ---
$report = [System.Collections.ArrayList]::new()
$pass = 0; $fail = 0; $warn = 0; $info = 0

function Log {
    param([string]$text, [string]$color = "White")
    Write-Host $text -ForegroundColor $color
    $script:report.Add($text) | Out-Null
}

function Tag {
    param([string]$tag, [string]$message, [string]$detail = "")
    switch ($tag) {
        "PASS" {
            $script:pass++
            $color = "Green"
        }
        "FAIL" {
            $script:fail++
            $color = "Red"
        }
        "WARN" {
            $script:warn++
            $color = "Yellow"
        }
        "INFO" {
            $script:info++
            $color = "Cyan"
        }
        default { $color = "White" }
    }
    $line = "  [$tag] $message"
    Write-Host $line -ForegroundColor $color
    $script:report.Add($line) | Out-Null
    if ($detail) {
        $detailLine = "        $detail"
        Write-Host $detailLine -ForegroundColor "DarkGray"
        $script:report.Add($detailLine) | Out-Null
    }
}

function Detail {
    param([string]$text)
    $line = "        $text"
    Write-Host $line -ForegroundColor "DarkGray"
    $script:report.Add($line) | Out-Null
}

function Section {
    param([string]$title)
    Log ""
    Log "  ========== $title ==========" "White"
    Log ""
}

function FileMD5 {
    param([string]$path)
    try {
        $hash = Get-FileHash -Path $path -Algorithm MD5 -ErrorAction Stop
        return $hash.Hash
    } catch {
        return "ERROR"
    }
}

# =====================================================================
#  HEADER
# =====================================================================
Log ""
Log "  ################################################################" "Cyan"
Log "  ##                                                            ##" "Cyan"
Log "  ##            zbar - Comprehensive Diagnostic Report          ##" "Cyan"
Log "  ##                                                            ##" "Cyan"
Log "  ################################################################" "Cyan"
Log ""
Log "  Report generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Log "  Script location:  $PSCommandPath"
Log "  USB directory:    $UsbDir"
Log ""

# =====================================================================
#  1. ENVIRONMENT
# =====================================================================
Section "1. ENVIRONMENT"

try {
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    $osCaption = $os.Caption
    $osVersion = $os.Version
    $osBuild = $os.BuildNumber
    $bootTime = $os.LastBootUpTime
    $uptime = (Get-Date) - $bootTime
    $uptimeStr = "$($uptime.Days)d $($uptime.Hours)h $($uptime.Minutes)m $($uptime.Seconds)s"
} catch {
    $osCaption = "Unknown"; $osVersion = "Unknown"; $osBuild = "Unknown"
    $uptimeStr = "Unknown"
}

Tag "INFO" "Computer name: $env:COMPUTERNAME"
Tag "INFO" "OS caption:    $osCaption"
Tag "INFO" "OS version:    $osVersion"
Tag "INFO" "OS build:      $osBuild"

$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
Tag "INFO" "Current user:  $($currentUser.Name)"
Tag "INFO" "User SID:      $($currentUser.User.Value)"

$isAdmin = ([Security.Principal.WindowsPrincipal]$currentUser).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin) {
    Tag "INFO" "Is admin:      True"
} else {
    Tag "WARN" "Is admin:      False (some checks may be limited)"
}

$psVer = $PSVersionTable.PSVersion
Tag "INFO" "PowerShell:    $($psVer.Major).$($psVer.Minor).$($psVer.Build)"

# .NET Framework version
try {
    $ndpKey = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
    if (Test-Path $ndpKey) {
        $release = (Get-ItemProperty $ndpKey -Name Release -ErrorAction SilentlyContinue).Release
        $dotnetVer = switch ($true) {
            ($release -ge 533320) { "4.8.1 or later (release $release)"; break }
            ($release -ge 528040) { "4.8 (release $release)"; break }
            ($release -ge 461808) { "4.7.2 (release $release)"; break }
            ($release -ge 461308) { "4.7.1 (release $release)"; break }
            ($release -ge 460798) { "4.7 (release $release)"; break }
            ($release -ge 394802) { "4.6.2 (release $release)"; break }
            ($release -ge 394254) { "4.6.1 (release $release)"; break }
            ($release -ge 393295) { "4.6 (release $release)"; break }
            default { "Pre-4.6 (release $release)" }
        }
        Tag "INFO" ".NET Framework: $dotnetVer"
    } else {
        Tag "INFO" ".NET Framework: v4 Full key not found"
    }
} catch {
    Tag "INFO" ".NET Framework: Could not determine"
}

Tag "INFO" "System uptime: $uptimeStr"

# =====================================================================
#  2. INSTALL DIRECTORY
# =====================================================================
Section "2. INSTALL DIRECTORY"

if (Test-Path $installDir) {
    Tag "PASS" "Install directory exists: $installDir"

    Log ""
    Log "  Files in install directory:" "White"
    Log "  $(('-' * 90))"
    Log "  $("{0,-30} {1,12} {2,20} {3}" -f 'Name','Size (bytes)','Last Modified','MD5')"
    Log "  $(('-' * 90))"

    $allFiles = Get-ChildItem $installDir -File -ErrorAction SilentlyContinue
    if ($allFiles) {
        foreach ($f in $allFiles) {
            $md5 = FileMD5 $f.FullName
            $size = "{0,12:N0}" -f $f.Length
            $date = $f.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
            Log "  $("{0,-30} {1} {2} {3}" -f $f.Name,$size,$date,$md5)"
        }
    } else {
        Log "  (directory is empty)" "Yellow"
    }
    Log "  $(('-' * 90))"
    Log ""

    # Check ACLs/permissions
    Log "  Directory ACL/Permissions:" "White"
    try {
        $acl = Get-Acl $installDir -ErrorAction Stop
        Tag "INFO" "Owner: $($acl.Owner)"
        foreach ($rule in $acl.Access) {
            $identity = $rule.IdentityReference
            $rights = $rule.FileSystemRights
            $type = $rule.AccessControlType
            $inherited = if ($rule.IsInherited) { "Inherited" } else { "Explicit" }
            Detail "$identity | $rights | $type | $inherited"
        }
    } catch {
        Tag "WARN" "Could not read ACLs: $($_.Exception.Message)"
    }
} else {
    Tag "FAIL" "Install directory NOT found: $installDir"
    Log ""
    Log "  zbar is not installed. Run Install.bat first." "Red"
}

# =====================================================================
#  3. CORE FILES
# =====================================================================
Section "3. CORE FILES"

# zbar.ps1
$zbarScript = Join-Path $installDir "zbar.ps1"
if (Test-Path $zbarScript) {
    $zbarSize = (Get-Item $zbarScript).Length
    if ($zbarSize -gt 0) {
        Tag "PASS" "zbar.ps1 exists and is non-empty ($zbarSize bytes)"
    } else {
        Tag "FAIL" "zbar.ps1 exists but is EMPTY (0 bytes)"
    }
} else {
    Tag "FAIL" "zbar.ps1 NOT found"
}

# blocklist.txt
$blFile = Join-Path $installDir "blocklist.txt"
if (Test-Path $blFile) {
    $blSize = (Get-Item $blFile).Length
    if ($blSize -gt 0) {
        Tag "PASS" "blocklist.txt exists and is non-empty ($blSize bytes)"

        # Count active entries
        $blContent = @(Get-Content $blFile -ErrorAction SilentlyContinue |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and $_ -notmatch '^\s*#' })
        Tag "INFO" "Active blocklist entries: $($blContent.Count)"

        if ($blContent.Count -gt 0) {
            Detail "Entries: $($blContent -join ', ')"
        }
    } else {
        Tag "FAIL" "blocklist.txt exists but is EMPTY (0 bytes)"
    }
} else {
    Tag "FAIL" "blocklist.txt NOT found"
}

# =====================================================================
#  4. ALL PERSISTENCE METHODS
# =====================================================================
Section "4. PERSISTENCE METHODS"

$activeMethods = @()

# --- 4a. Scheduled Task ---
Log "  -- Scheduled Task --" "White"
$task = Get-ScheduledTask -TaskName "zbar" -ErrorAction SilentlyContinue
if ($task) {
    $state = $task.State
    if ($state -eq "Ready" -or $state -eq "Running") {
        Tag "PASS" "Scheduled task 'zbar' exists (State: $state)"
    } else {
        Tag "WARN" "Scheduled task 'zbar' exists but State: $state"
    }
    $activeMethods += "ScheduledTask"

    try {
        $taskInfo = Get-ScheduledTaskInfo -TaskName "zbar" -ErrorAction SilentlyContinue
        if ($taskInfo) {
            Detail "Last run time:   $($taskInfo.LastRunTime)"
            Detail "Last result:     $($taskInfo.LastTaskResult)"
            Detail "Next run time:   $($taskInfo.NextRunTime)"
        }
    } catch {}

    try {
        foreach ($a in $task.Actions) {
            Detail "Action execute:  $($a.Execute)"
            Detail "Action args:     $($a.Arguments)"
        }
        foreach ($t in $task.Triggers) {
            Detail "Trigger type:    $($t.CimClass.CimClassName)"
            if ($t.Delay) { Detail "Trigger delay:   $($t.Delay)" }
            if ($t.Repetition -and $t.Repetition.Interval) {
                Detail "Repeat interval: $($t.Repetition.Interval)"
            }
        }
        Detail "Principal user:  $($task.Principal.UserId)"
        Detail "Principal logon: $($task.Principal.LogonType)"
        Detail "Run level:       $($task.Principal.RunLevel)"
    } catch {}
} else {
    Tag "FAIL" "Scheduled task 'zbar' NOT found"
}

Log ""

# --- 4b. Registry Run HKCU ---
Log "  -- Registry Run (HKCU) --" "White"
$regHKCU = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
try {
    $regValHKCU = Get-ItemProperty -Path $regHKCU -Name "zbar" -ErrorAction SilentlyContinue
    if ($regValHKCU) {
        Tag "PASS" "HKCU Run entry exists"
        Detail "Value: $($regValHKCU.zbar)"
        $activeMethods += "HKCU_Run"
    } else {
        Tag "INFO" "HKCU Run entry not present"
    }
} catch {
    Tag "INFO" "HKCU Run entry not present"
}

Log ""

# --- 4c. Registry Run HKLM ---
Log "  -- Registry Run (HKLM) --" "White"
$regHKLM = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"
try {
    $regValHKLM = Get-ItemProperty -Path $regHKLM -Name "zbar" -ErrorAction SilentlyContinue
    if ($regValHKLM) {
        Tag "PASS" "HKLM Run entry exists"
        Detail "Value: $($regValHKLM.zbar)"
        $activeMethods += "HKLM_Run"
    } else {
        Tag "INFO" "HKLM Run entry not present"
    }
} catch {
    Tag "INFO" "HKLM Run entry not present"
}

Log ""

# --- 4d. All Users Startup folder ---
Log "  -- All Users Startup Folder --" "White"
$allUsersStartup = [Environment]::GetFolderPath("CommonStartup")
if (-not $allUsersStartup) {
    $allUsersStartup = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
}
if (Test-Path $allUsersStartup) {
    $startupFiles = Get-ChildItem $allUsersStartup -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*zbar*" }
    if ($startupFiles) {
        Tag "PASS" "Found zbar files in All Users Startup"
        foreach ($sf in $startupFiles) {
            Detail "$($sf.Name) ($($sf.Length) bytes)"
        }
        $activeMethods += "AllUsersStartup"
    } else {
        Tag "INFO" "No zbar files in All Users Startup"
        Detail "Path: $allUsersStartup"
    }
} else {
    Tag "WARN" "All Users Startup folder not found: $allUsersStartup"
}

Log ""

# --- 4e. Current User Startup folder ---
Log "  -- Current User Startup Folder --" "White"
$userStartup = [Environment]::GetFolderPath("Startup")
if (-not $userStartup) {
    $userStartup = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
}
if (Test-Path $userStartup) {
    $userStartupFiles = Get-ChildItem $userStartup -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*zbar*" }
    if ($userStartupFiles) {
        Tag "PASS" "Found zbar files in Current User Startup"
        foreach ($sf in $userStartupFiles) {
            Detail "$($sf.Name) ($($sf.Length) bytes)"
        }
        $activeMethods += "UserStartup"
    } else {
        Tag "INFO" "No zbar files in Current User Startup"
        Detail "Path: $userStartup"
    }
} else {
    Tag "WARN" "Current User Startup folder not found: $userStartup"
}

Log ""

# --- 4f. Persistence Summary ---
Log "  -- Persistence Summary --" "White"
if ($activeMethods.Count -gt 0) {
    Tag "PASS" "Active persistence methods ($($activeMethods.Count)): $($activeMethods -join ', ')"
} else {
    Tag "FAIL" "NO persistence methods active! zbar will not survive a reboot."
}

# =====================================================================
#  5. PROCESS STATUS
# =====================================================================
Section "5. PROCESS STATUS"

$zbarProcs = @()
try {
    $allPSProcs = @(Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue)
    $zbarProcs = @($allPSProcs | Where-Object { $_.CommandLine -and $_.CommandLine -like "*zbar*" })
} catch {
    Tag "WARN" "Could not query Win32_Process: $($_.Exception.Message)"
}

if ($zbarProcs.Count -gt 0) {
    Tag "PASS" "zbar is RUNNING ($($zbarProcs.Count) matching process(es))"
    Log ""
    foreach ($p in $zbarProcs) {
        $procUptime = (Get-Date) - $p.CreationDate
        $procUptimeStr = ""
        if ($procUptime.Days -gt 0) { $procUptimeStr += "$($procUptime.Days)d " }
        $procUptimeStr += "$($procUptime.Hours)h $($procUptime.Minutes)m $($procUptime.Seconds)s"

        # Try to get session user
        $procUser = "Unknown"
        try {
            $owner = Invoke-CimMethod -InputObject $p -MethodName GetOwner -ErrorAction SilentlyContinue
            if ($owner -and $owner.User) {
                $procUser = "$($owner.Domain)\$($owner.User)"
            }
        } catch {}

        Detail "PID:         $($p.ProcessId)"
        Detail "Start time:  $($p.CreationDate)"
        Detail "Uptime:      $procUptimeStr"
        Detail "Session ID:  $($p.SessionId)"
        Detail "User:        $procUser"
        Detail "CommandLine: $($p.CommandLine)"
        Log ""
    }
} else {
    Tag "FAIL" "zbar is NOT running (no powershell.exe with 'zbar' in CommandLine)"
    Log ""

    if ($allPSProcs.Count -gt 0) {
        Tag "INFO" "All powershell.exe processes ($($allPSProcs.Count)) for debugging:"
        Log ""
        foreach ($p in $allPSProcs) {
            $cmdLine = if ($p.CommandLine) { $p.CommandLine } else { "(no CommandLine)" }
            Detail "PID $($p.ProcessId) | Session $($p.SessionId) | $cmdLine"
        }
    } else {
        Tag "INFO" "No powershell.exe processes running at all"
    }
}

# =====================================================================
#  6. PID FILE
# =====================================================================
Section "6. PID FILE"

$pidFile = Join-Path $installDir "zbar.pid"
if (Test-Path $pidFile) {
    Tag "PASS" "PID file exists: $pidFile"
    $storedPid = (Get-Content $pidFile -ErrorAction SilentlyContinue)
    if ($storedPid) {
        $storedPid = $storedPid.Trim()
        Detail "Stored PID: $storedPid"

        try {
            $pidProc = Get-Process -Id ([int]$storedPid) -ErrorAction SilentlyContinue
            if ($pidProc) {
                Tag "PASS" "PID $storedPid is alive (Process: $($pidProc.ProcessName), StartTime: $($pidProc.StartTime))"
            } else {
                Tag "WARN" "PID $storedPid is NOT running (stale PID file)"
            }
        } catch {
            Tag "WARN" "PID $storedPid is NOT running (stale PID file)"
        }

        # Also check what process actually has that PID (if any)
        try {
            $pidCim = Get-CimInstance Win32_Process -Filter "ProcessId=$storedPid" -ErrorAction SilentlyContinue
            if ($pidCim) {
                Detail "Process at PID $storedPid : $($pidCim.Name)"
                Detail "CommandLine: $($pidCim.CommandLine)"
            }
        } catch {}
    } else {
        Tag "WARN" "PID file is empty"
    }
} else {
    Tag "WARN" "No PID file found at $pidFile"
}

# =====================================================================
#  7. KILL LOG
# =====================================================================
Section "7. KILL LOG"

$logFile = Join-Path $installDir "zbar-log.txt"
$kills = @()
$failLines = @()

if (Test-Path $logFile) {
    $logSize = (Get-Item $logFile).Length
    $logLines = @(Get-Content $logFile -ErrorAction SilentlyContinue)
    $kills = @($logLines | Where-Object { $_ -match "KILLED" })
    $failLines = @($logLines | Where-Object { $_ -match "FAILED" })

    Tag "PASS" "Log file exists: $logFile"
    Detail "Size:     $("{0:N0}" -f $logSize) bytes"
    Detail "Lines:    $($logLines.Count)"
    Detail "Kills:    $($kills.Count)"
    Detail "Failures: $($failLines.Count)"

    # Kills by process name (top 10)
    if ($kills.Count -gt 0) {
        Log ""
        Log "  Top 10 most killed processes:" "White"
        $killNames = @{}
        foreach ($k in $kills) {
            # Try to extract process name from log line
            # Expected format: "2025-01-01 12:00:00  KILLED  processname.exe  (PID 1234)"
            if ($k -match "KILLED\s+(.+?)\s+\(PID") {
                $procName = $Matches[1].Trim()
                if ($killNames.ContainsKey($procName)) {
                    $killNames[$procName]++
                } else {
                    $killNames[$procName] = 1
                }
            }
        }
        $sorted = $killNames.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10
        foreach ($entry in $sorted) {
            Detail "$("{0,6}" -f $entry.Value) kills  $($entry.Key)"
        }
    }

    # Last 50 kills
    if ($kills.Count -gt 0) {
        Log ""
        Log "  Last 50 kills:" "White"
        Log "  $(('-' * 80))"
        $recent = $kills | Select-Object -Last 50
        foreach ($line in $recent) {
            if ($line -match "^(\S+ \S+)\s+KILLED\s+(.+?)\s+\(PID (\d+)\)") {
                Detail "$($Matches[1])  $($Matches[2])  PID $($Matches[3])"
            } else {
                Detail $line
            }
        }
        Log "  $(('-' * 80))"
    }
} else {
    Tag "WARN" "No log file found at $logFile (no kills recorded yet)"
}

# =====================================================================
#  8. POWERSHELL EXECUTION POLICY
# =====================================================================
Section "8. POWERSHELL EXECUTION POLICY"

$scopes = @("MachinePolicy", "UserPolicy", "Process", "CurrentUser", "LocalMachine")
foreach ($scope in $scopes) {
    try {
        $pol = Get-ExecutionPolicy -Scope $scope -ErrorAction SilentlyContinue
        $polStr = "$pol"
        if ($polStr -eq "Restricted") {
            Tag "WARN" "ExecutionPolicy ($scope): $polStr"
        } elseif ($polStr -eq "AllSigned") {
            Tag "WARN" "ExecutionPolicy ($scope): $polStr (may block unsigned scripts)"
        } elseif ($polStr -eq "Undefined") {
            Tag "INFO" "ExecutionPolicy ($scope): $polStr"
        } else {
            Tag "PASS" "ExecutionPolicy ($scope): $polStr"
        }
    } catch {
        Tag "WARN" "ExecutionPolicy ($scope): Could not query"
    }
}

# =====================================================================
#  9. POTENTIAL BLOCKERS
# =====================================================================
Section "9. POTENTIAL BLOCKERS"

# --- 9a. Language Mode ---
Log "  -- PowerShell Language Mode --" "White"
$langMode = $ExecutionContext.SessionState.LanguageMode
if ($langMode -eq "FullLanguage") {
    Tag "PASS" "Language mode: $langMode"
} elseif ($langMode -eq "ConstrainedLanguage") {
    Tag "FAIL" "Language mode: $langMode (Constrained Language Mode blocks many operations)"
} else {
    Tag "WARN" "Language mode: $langMode"
}

Log ""

# --- 9b. Windows Defender ---
Log "  -- Windows Defender --" "White"
try {
    $defenderStatus = Get-MpComputerStatus -ErrorAction Stop
    $rtProtection = $defenderStatus.RealTimeProtectionEnabled
    if ($rtProtection) {
        Tag "INFO" "Defender Real-Time Protection: Enabled"
    } else {
        Tag "INFO" "Defender Real-Time Protection: Disabled"
    }

    # Check exclusions
    try {
        $prefs = Get-MpPreference -ErrorAction Stop
        $pathExclusions = $prefs.ExclusionPath
        $procExclusions = $prefs.ExclusionProcess

        if ($pathExclusions -and ($pathExclusions -contains $installDir -or ($pathExclusions | Where-Object { $installDir -like "$_*" }))) {
            Tag "PASS" "Install directory is in Defender exclusion paths"
        } else {
            if ($rtProtection) {
                Tag "WARN" "Install directory is NOT excluded from Defender (may interfere with PowerShell scripts)"
            } else {
                Tag "INFO" "Install directory is NOT in Defender exclusions (but RT protection is off)"
            }
        }

        if ($pathExclusions) {
            Detail "Path exclusions: $($pathExclusions -join '; ')"
        }
        if ($procExclusions) {
            Detail "Process exclusions: $($procExclusions -join '; ')"
        }
    } catch {
        Tag "INFO" "Could not query Defender exclusions (may need admin)"
    }
} catch {
    Tag "INFO" "Windows Defender status not available (may not be installed or need admin)"
}

Log ""

# --- 9c. Group Policy script restrictions ---
Log "  -- Group Policy Script Restrictions --" "White"
try {
    $gpScriptKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell"
    if (Test-Path $gpScriptKey) {
        $gpProps = Get-ItemProperty $gpScriptKey -ErrorAction SilentlyContinue
        Tag "WARN" "PowerShell Group Policy key exists"
        if ($gpProps.EnableScripts -eq 0) {
            Tag "FAIL" "Group Policy DISABLES PowerShell scripts (EnableScripts=0)"
        } elseif ($gpProps.EnableScripts -eq 1) {
            Tag "PASS" "Group Policy allows PowerShell scripts (EnableScripts=1)"
            Detail "ExecutionPolicy via GP: $($gpProps.ExecutionPolicy)"
        } else {
            Tag "INFO" "EnableScripts value: $($gpProps.EnableScripts)"
        }
    } else {
        Tag "PASS" "No PowerShell Group Policy restrictions found"
    }
} catch {
    Tag "INFO" "Could not check Group Policy keys"
}

Log ""

# --- 9d. AppLocker ---
Log "  -- AppLocker / WDAC --" "White"
try {
    $appLockerKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\SrpV2"
    if (Test-Path $appLockerKey) {
        $subKeys = Get-ChildItem $appLockerKey -ErrorAction SilentlyContinue
        if ($subKeys) {
            Tag "WARN" "AppLocker policies detected (may restrict script execution)"
            foreach ($sk in $subKeys) {
                Detail "Policy category: $($sk.PSChildName)"
                $rules = Get-ChildItem $sk.PSPath -ErrorAction SilentlyContinue
                if ($rules) {
                    Detail "  Rules count: $($rules.Count)"
                }
            }
        } else {
            Tag "PASS" "AppLocker key exists but no policy sub-keys found"
        }
    } else {
        Tag "PASS" "No AppLocker policies found"
    }
} catch {
    Tag "INFO" "Could not check AppLocker policies"
}

# WDAC / Device Guard
try {
    $dgKey = "HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard"
    if (Test-Path $dgKey) {
        $dgProps = Get-ItemProperty $dgKey -ErrorAction SilentlyContinue
        if ($dgProps.EnableVirtualizationBasedSecurity -eq 1) {
            Tag "WARN" "Device Guard / WDAC is enabled (VBS active)"
        } else {
            Tag "INFO" "Device Guard key exists but VBS not enabled"
        }

        $ciKey = "HKLM:\SYSTEM\CurrentControlSet\Control\CI"
        if (Test-Path $ciKey) {
            $ciProps = Get-ItemProperty $ciKey -ErrorAction SilentlyContinue
            if ($ciProps.UMCIAuditMode -eq 1) {
                Tag "INFO" "Code Integrity in audit mode"
            }
        }
    } else {
        Tag "PASS" "No Device Guard / WDAC configuration found"
    }
} catch {
    Tag "INFO" "Could not check WDAC / Device Guard"
}

# =====================================================================
#  10. TASK SCHEDULER DEEP DIVE
# =====================================================================
Section "10. TASK SCHEDULER DEEP DIVE"

# --- Check Task Scheduler service ---
Log "  -- Task Scheduler Service --" "White"
try {
    $schedSvc = Get-Service -Name "Schedule" -ErrorAction Stop
    if ($schedSvc.Status -eq "Running") {
        Tag "PASS" "Task Scheduler service is running"
    } else {
        Tag "FAIL" "Task Scheduler service status: $($schedSvc.Status)"
    }
    Detail "Service name: $($schedSvc.Name)"
    Detail "Display name: $($schedSvc.DisplayName)"
    Detail "Start type:   $($schedSvc.StartType)"
} catch {
    Tag "FAIL" "Could not query Task Scheduler service"
}

Log ""

# --- Search ALL tasks for anything zbar-related ---
Log "  -- All zbar-related Scheduled Tasks --" "White"
try {
    $allTasks = Get-ScheduledTask -ErrorAction SilentlyContinue |
        Where-Object {
            $_.TaskName -like "*zbar*" -or
            $_.TaskPath -like "*zbar*" -or
            ($_.Actions | ForEach-Object { "$($_.Execute) $($_.Arguments)" }) -like "*zbar*"
        }

    if ($allTasks) {
        Tag "INFO" "Found $($allTasks.Count) zbar-related task(s) in Task Scheduler"
        Log ""
        foreach ($t in $allTasks) {
            Log "  Task: $($t.TaskPath)$($t.TaskName)" "White"
            Detail "State:   $($t.State)"
            Detail "URI:     $($t.URI)"

            foreach ($a in $t.Actions) {
                Detail "Action:  $($a.Execute) $($a.Arguments)"
            }
            foreach ($tr in $t.Triggers) {
                Detail "Trigger: $($tr.CimClass.CimClassName)"
            }
            Detail "User:    $($t.Principal.UserId)"
            Detail "LogonType: $($t.Principal.LogonType)"
            Detail "RunLevel:  $($t.Principal.RunLevel)"

            try {
                $ti = Get-ScheduledTaskInfo -TaskName $t.TaskName -TaskPath $t.TaskPath -ErrorAction SilentlyContinue
                if ($ti) {
                    Detail "LastRun:    $($ti.LastRunTime)"
                    Detail "LastResult: $($ti.LastTaskResult)"
                    Detail "NextRun:    $($ti.NextRunTime)"
                    Detail "MissedRuns: $($ti.NumberOfMissedRuns)"
                }
            } catch {}

            # Full XML definition
            Log ""
            Log "  Full XML definition:" "White"
            try {
                $xml = Export-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -ErrorAction Stop
                foreach ($xmlLine in ($xml -split "`n")) {
                    Detail $xmlLine.TrimEnd()
                }
            } catch {
                Detail "Could not export XML: $($_.Exception.Message)"
            }
            Log ""
        }
    } else {
        Tag "INFO" "No zbar-related tasks found in Task Scheduler"
    }
} catch {
    Tag "WARN" "Could not enumerate scheduled tasks: $($_.Exception.Message)"
}

# =====================================================================
#  SUMMARY
# =====================================================================
Log ""
Log "  ################################################################" "Cyan"
Log "  ##                        SUMMARY                            ##" "Cyan"
Log "  ################################################################" "Cyan"
Log ""
$totalChecks = $pass + $fail + $warn + $info

$passColor = if ($pass -gt 0) { "Green" } else { "White" }
$failColor = if ($fail -gt 0) { "Red" } else { "White" }
$warnColor = if ($warn -gt 0) { "Yellow" } else { "White" }
$infoColor = if ($info -gt 0) { "Cyan" } else { "White" }

Log "  Total checks: $totalChecks"
Write-Host "  [PASS] $pass" -ForegroundColor $passColor -NoNewline; Write-Host "  |  " -NoNewline
Write-Host "[FAIL] $fail" -ForegroundColor $failColor -NoNewline; Write-Host "  |  " -NoNewline
Write-Host "[WARN] $warn" -ForegroundColor $warnColor -NoNewline; Write-Host "  |  " -NoNewline
Write-Host "[INFO] $info" -ForegroundColor $infoColor
$report.Add("  [PASS] $pass  |  [FAIL] $fail  |  [WARN] $warn  |  [INFO] $info") | Out-Null

Log ""

if ($fail -eq 0 -and $warn -eq 0) {
    Log "  VERDICT: ALL CLEAR - zbar is fully installed, running, and healthy." "Green"
} elseif ($fail -eq 0) {
    Log "  VERDICT: OK with $warn warning(s) - zbar is functional but review warnings above." "Yellow"
} elseif ($fail -le 2) {
    Log "  VERDICT: ISSUES DETECTED - $fail failure(s) and $warn warning(s). Review failures above." "Red"
} else {
    Log "  VERDICT: CRITICAL - $fail failure(s) and $warn warning(s). zbar may not be functional." "Red"
}

Log ""
Log "  ################################################################" "Cyan"

# =====================================================================
#  WRITE REPORT TO FILE
# =====================================================================
$reportPath = ""
if ($UsbDir -and (Test-Path $UsbDir)) {
    $reportPath = $UsbDir
}

if ($reportPath) {
    $reportFile = Join-Path $reportPath "zbar-report.txt"
    try {
        $report | Out-File $reportFile -Encoding utf8 -Force
        Log ""
        Log "  Report saved to: $reportFile" "Green"
    } catch {
        Log ""
        Log "  ERROR: Could not save report to $reportFile : $($_.Exception.Message)" "Red"
    }
} else {
    Log ""
    Log "  Report not saved to file (no writable USB/directory detected)." "Yellow"
}

Log ""
Read-Host "  Press Enter to exit"
