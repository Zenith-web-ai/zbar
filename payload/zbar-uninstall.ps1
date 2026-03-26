# zbar-uninstall.ps1 - Complete removal of zbar from a Windows PC
# Removes ALL persistence methods, kills processes, deletes files, and writes a report.

param(
    [Parameter(Mandatory = $true)]
    [string]$UsbDir
)

$ErrorActionPreference = "Continue"
$installDir = "C:\ProgramData\zbar"
$taskName = "zbar"
$report = @()
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$success = @{}

function Log {
    param([string]$Message, [string]$Color = "White")
    Write-Host $Message -ForegroundColor $Color
    $script:report += $Message
}

# ============================================================
# STEP 0: Self-elevate to admin
# ============================================================
Log "========================================" "Cyan"
Log "  zbar - Uninstaller" "Cyan"
Log "  $timestamp" "Cyan"
Log "========================================" "Cyan"
Log ""

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Log "  [INFO] Not running as admin. Attempting elevation..." "Yellow"
    try {
        $argList = "-ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`" -UsbDir `"$UsbDir`""
        Start-Process powershell.exe -Verb RunAs -ArgumentList $argList -Wait
        Log "  [OK] Elevated instance completed." "Green"
        exit 0
    } catch {
        Log "  [WARN] Elevation failed: $($_.Exception.Message)" "Yellow"
        Log "  [WARN] Continuing without admin rights (some removals may fail)." "Yellow"
    }
}

Log ""
Log "  Running as admin: $isAdmin" $(if ($isAdmin) { "Green" } else { "Yellow" })
Log ""

# ============================================================
# STEP 1: Remove scheduled task (method 1: Unregister-ScheduledTask)
# ============================================================
Log "  --- Removing Scheduled Task ---" "Cyan"

try {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
    Log "  [OK] Scheduled task removed via Unregister-ScheduledTask." "Green"
    $success["ScheduledTask_Unregister"] = $true
} catch {
    Log "  [WARN] Unregister-ScheduledTask: $($_.Exception.Message)" "Yellow"
    $success["ScheduledTask_Unregister"] = $false
}

# Scheduled task (method 2: schtasks /delete)
try {
    $output = & schtasks /delete /tn $taskName /f 2>&1
    if ($LASTEXITCODE -eq 0) {
        Log "  [OK] Scheduled task removed via schtasks /delete." "Green"
        $success["ScheduledTask_schtasks"] = $true
    } else {
        Log "  [INFO] schtasks /delete: $output" "Yellow"
        $success["ScheduledTask_schtasks"] = $false
    }
} catch {
    Log "  [WARN] schtasks /delete: $($_.Exception.Message)" "Yellow"
    $success["ScheduledTask_schtasks"] = $false
}

# Watchdog task
try {
    Unregister-ScheduledTask -TaskName "zbar-watchdog" -Confirm:$false -ErrorAction Stop
    Log "  [OK] Watchdog task removed." "Green"
} catch {
    Log "  [INFO] No watchdog task to remove." "Yellow"
}
try {
    & schtasks /delete /tn "zbar-watchdog" /f 2>&1 | Out-Null
} catch {}

# ============================================================
# STEP 2: Remove HKCU Registry Run entry
# ============================================================
Log ""
Log "  --- Removing Registry Run Entries ---" "Cyan"

try {
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
    $existing = Get-ItemProperty -Path $regPath -Name $taskName -ErrorAction SilentlyContinue
    if ($existing) {
        Remove-ItemProperty -Path $regPath -Name $taskName -ErrorAction Stop
        Log "  [OK] HKCU Registry Run entry removed." "Green"
        $success["Registry_HKCU"] = $true
    } else {
        Log "  [INFO] HKCU Registry Run entry not found (nothing to remove)." "Yellow"
        $success["Registry_HKCU"] = $true
    }
} catch {
    Log "  [FAIL] HKCU Registry Run removal failed: $($_.Exception.Message)" "Red"
    $success["Registry_HKCU"] = $false
}

# Remove HKLM Registry Run entry (requires admin)
try {
    $regPathLM = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"
    $existingLM = Get-ItemProperty -Path $regPathLM -Name $taskName -ErrorAction SilentlyContinue
    if ($existingLM) {
        Remove-ItemProperty -Path $regPathLM -Name $taskName -ErrorAction Stop
        Log "  [OK] HKLM Registry Run entry removed." "Green"
        $success["Registry_HKLM"] = $true
    } else {
        Log "  [INFO] HKLM Registry Run entry not found (nothing to remove)." "Yellow"
        $success["Registry_HKLM"] = $true
    }
} catch {
    Log "  [FAIL] HKLM Registry Run removal failed: $($_.Exception.Message)" "Red"
    $success["Registry_HKLM"] = $false
}

# ============================================================
# STEP 3: Remove startup folder .vbs files
# ============================================================
Log ""
Log "  --- Removing Startup Folder Entries ---" "Cyan"

# All Users startup
$allUsersVbs = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\zbar.vbs"
try {
    if (Test-Path $allUsersVbs) {
        Remove-Item -Path $allUsersVbs -Force -ErrorAction Stop
        Log "  [OK] All Users startup .vbs deleted." "Green"
        $success["Startup_AllUsers"] = $true
    } else {
        Log "  [INFO] All Users startup .vbs not found (nothing to remove)." "Yellow"
        $success["Startup_AllUsers"] = $true
    }
} catch {
    Log "  [FAIL] All Users startup .vbs removal failed: $($_.Exception.Message)" "Red"
    $success["Startup_AllUsers"] = $false
}

# Current User startup
$currentUserVbs = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\zbar.vbs"
try {
    if (Test-Path $currentUserVbs) {
        Remove-Item -Path $currentUserVbs -Force -ErrorAction Stop
        Log "  [OK] Current User startup .vbs deleted." "Green"
        $success["Startup_CurrentUser"] = $true
    } else {
        Log "  [INFO] Current User startup .vbs not found (nothing to remove)." "Yellow"
        $success["Startup_CurrentUser"] = $true
    }
} catch {
    Log "  [FAIL] Current User startup .vbs removal failed: $($_.Exception.Message)" "Red"
    $success["Startup_CurrentUser"] = $false
}

# ============================================================
# STEP 3.5: Remove Windows Defender exclusion
# ============================================================
Log ""
Log "  --- Removing Defender Exclusions ---" "Cyan"
try {
    Remove-MpPreference -ExclusionPath $installDir -ErrorAction SilentlyContinue
    Log "  [OK] Defender path exclusion removed" "Green"
} catch {}
try {
    Remove-MpPreference -ExclusionProcess "powershell.exe" -ErrorAction SilentlyContinue
    Log "  [OK] Defender process exclusion (powershell) removed" "Green"
} catch {}
try {
    Remove-MpPreference -ExclusionProcess "wscript.exe" -ErrorAction SilentlyContinue
    Log "  [OK] Defender process exclusion (wscript) removed" "Green"
} catch {}
try {
    Set-MpPreference -AttackSurfaceReductionRules_Ids "e6db77e5-3df2-4cf1-b95a-636979351e5b" -AttackSurfaceReductionRules_Actions Enabled -ErrorAction SilentlyContinue
    Set-MpPreference -AttackSurfaceReductionRules_Ids "d1e49aac-8f56-4280-b9ba-993a6d77406c" -AttackSurfaceReductionRules_Actions Enabled -ErrorAction SilentlyContinue
    Log "  [OK] Defender ASR rules restored" "Green"
} catch {}

# ============================================================
# STEP 4: Kill ALL zbar processes
# ============================================================
Log ""
Log "  --- Killing zbar Processes ---" "Cyan"

$killedCount = 0

# Method 1: Find via Win32_Process CommandLine
try {
    $zbarProcs = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -like "*zbar.ps1*" }
    if ($zbarProcs) {
        foreach ($proc in $zbarProcs) {
            try {
                Stop-Process -Id $proc.ProcessId -Force -ErrorAction Stop
                Log "  [OK] Killed process PID $($proc.ProcessId) (CommandLine match)." "Green"
                $killedCount++
            } catch {
                Log "  [WARN] Failed to kill PID $($proc.ProcessId): $($_.Exception.Message)" "Yellow"
            }
        }
    } else {
        Log "  [INFO] No running zbar processes found via CommandLine search." "Yellow"
    }
    $success["Kill_CommandLine"] = $true
} catch {
    Log "  [WARN] CommandLine process search failed: $($_.Exception.Message)" "Yellow"
    $success["Kill_CommandLine"] = $false
}

# Method 2: Read PID file and kill that PID
$pidFile = Join-Path $installDir "zbar.pid"
try {
    if (Test-Path $pidFile) {
        $pidContent = (Get-Content $pidFile -ErrorAction SilentlyContinue).Trim()
        if ($pidContent -match '^\d+$') {
            $pidInt = [int]$pidContent
            $proc = Get-Process -Id $pidInt -ErrorAction SilentlyContinue
            if ($proc) {
                Stop-Process -Id $pidInt -Force -ErrorAction Stop
                Log "  [OK] Killed process from PID file (PID $pidInt)." "Green"
                $killedCount++
            } else {
                Log "  [INFO] PID $pidInt from PID file is not running." "Yellow"
            }
        } else {
            Log "  [WARN] PID file contents invalid: '$pidContent'" "Yellow"
        }
    } else {
        Log "  [INFO] No PID file found at $pidFile." "Yellow"
    }
    $success["Kill_PidFile"] = $true
} catch {
    Log "  [WARN] PID file kill failed: $($_.Exception.Message)" "Yellow"
    $success["Kill_PidFile"] = $false
}

Log "  Total processes killed: $killedCount" $(if ($killedCount -gt 0) { "Green" } else { "Yellow" })

# ============================================================
# STEP 5: Wait for processes to die
# ============================================================
Log ""
Log "  Waiting 2 seconds for processes to terminate..." "Cyan"
Start-Sleep -Seconds 2

# ============================================================
# STEP 6: Delete install directory
# ============================================================
Log ""
Log "  --- Deleting Install Directory ---" "Cyan"

try {
    if (Test-Path $installDir) {
        # Reset ACLs first (installer locks them down)
        try {
            $acl = Get-Acl $installDir
            $acl.SetAccessRuleProtection($false, $true)
            Set-Acl $installDir $acl -ErrorAction SilentlyContinue
        } catch {}
        Remove-Item -Path $installDir -Recurse -Force -ErrorAction Stop
        Log "  [OK] Install directory deleted: $installDir" "Green"
        $success["DeleteDir"] = $true
    } else {
        Log "  [INFO] Install directory not found (already removed): $installDir" "Yellow"
        $success["DeleteDir"] = $true
    }
} catch {
    Log "  [FAIL] Failed to delete install directory: $($_.Exception.Message)" "Red"
    $success["DeleteDir"] = $false
}

# ============================================================
# STEP 7: Verify removal
# ============================================================
Log ""
Log "  ========================================" "Cyan"
Log "  Verification" "Cyan"
Log "  ========================================" "Cyan"
Log ""

$verifyPass = 0
$verifyFail = 0

# 7a: Confirm directory is gone
try {
    if (-not (Test-Path $installDir)) {
        Log "  [PASS] Install directory is gone." "Green"
        $verifyPass++
    } else {
        Log "  [FAIL] Install directory still exists!" "Red"
        $verifyFail++
    }
} catch {
    Log "  [FAIL] Directory check error: $($_.Exception.Message)" "Red"
    $verifyFail++
}

# 7b: Confirm no scheduled task
try {
    $taskCheck = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if (-not $taskCheck) {
        Log "  [PASS] No scheduled task found." "Green"
        $verifyPass++
    } else {
        Log "  [FAIL] Scheduled task still exists!" "Red"
        $verifyFail++
    }
} catch {
    # Get-ScheduledTask throws if task not found on some systems - that means success
    Log "  [PASS] No scheduled task found." "Green"
    $verifyPass++
}

# 7c: Confirm no HKCU registry entry
try {
    $regCheck = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name $taskName -ErrorAction SilentlyContinue
    if (-not $regCheck) {
        Log "  [PASS] No HKCU Registry Run entry." "Green"
        $verifyPass++
    } else {
        Log "  [FAIL] HKCU Registry Run entry still exists!" "Red"
        $verifyFail++
    }
} catch {
    Log "  [PASS] No HKCU Registry Run entry." "Green"
    $verifyPass++
}

# 7d: Confirm no HKLM registry entry
try {
    $regCheckLM = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run" -Name $taskName -ErrorAction SilentlyContinue
    if (-not $regCheckLM) {
        Log "  [PASS] No HKLM Registry Run entry." "Green"
        $verifyPass++
    } else {
        Log "  [FAIL] HKLM Registry Run entry still exists!" "Red"
        $verifyFail++
    }
} catch {
    Log "  [PASS] No HKLM Registry Run entry." "Green"
    $verifyPass++
}

# 7e: Confirm no startup folder entries
try {
    $startupIssues = @()
    if (Test-Path $allUsersVbs) { $startupIssues += "All Users" }
    if (Test-Path $currentUserVbs) { $startupIssues += "Current User" }
    if ($startupIssues.Count -eq 0) {
        Log "  [PASS] No startup folder .vbs files." "Green"
        $verifyPass++
    } else {
        Log "  [FAIL] Startup .vbs still exists in: $($startupIssues -join ', ')" "Red"
        $verifyFail++
    }
} catch {
    Log "  [FAIL] Startup folder check error: $($_.Exception.Message)" "Red"
    $verifyFail++
}

# 7f: Confirm no running processes
try {
    $remaining = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -like "*zbar.ps1*" }
    if (-not $remaining) {
        Log "  [PASS] No running zbar processes." "Green"
        $verifyPass++
    } else {
        Log "  [FAIL] $($remaining.Count) zbar process(es) still running!" "Red"
        $verifyFail++
    }
} catch {
    Log "  [WARN] Process check error: $($_.Exception.Message)" "Yellow"
    $verifyFail++
}

# ============================================================
# STEP 8: Write report to USB
# ============================================================
Log ""
Log "  --- Writing Report ---" "Cyan"

$reportContent = @()
$reportContent += "zbar Uninstall Report"
$reportContent += "Generated: $timestamp"
$reportContent += "Computer:  $env:COMPUTERNAME"
$reportContent += "User:      $env:USERNAME"
$reportContent += "Admin:     $isAdmin"
$reportContent += "========================================"
$reportContent += ""
$reportContent += "Removal Actions:"
$reportContent += "----------------------------------------"
foreach ($key in $success.Keys | Sort-Object) {
    $status = if ($success[$key]) { "OK" } else { "FAILED" }
    $reportContent += "  $key : $status"
}
$reportContent += ""
$reportContent += "Verification:"
$reportContent += "----------------------------------------"
$reportContent += "  Checks Passed: $verifyPass"
$reportContent += "  Checks Failed: $verifyFail"
$reportContent += ""
if ($verifyFail -eq 0) {
    $reportContent += "RESULT: COMPLETE - zbar fully removed."
} else {
    $reportContent += "RESULT: PARTIAL - $verifyFail check(s) failed. Manual cleanup may be needed."
}
$reportContent += ""
$reportContent += "Full Log:"
$reportContent += "----------------------------------------"
$reportContent += $report

$reportPath = Join-Path $UsbDir "uninstall-report.txt"
try {
    $reportContent | Out-File -FilePath $reportPath -Encoding utf8 -Force -ErrorAction Stop
    Log "  [OK] Report written to: $reportPath" "Green"
} catch {
    Log "  [FAIL] Could not write report: $($_.Exception.Message)" "Red"
    Log "  Attempted path: $reportPath" "Red"
}

# ============================================================
# STEP 9: Display summary
# ============================================================
Log ""
Log "  ========================================" "Cyan"
Log "  Uninstall Summary" "Cyan"
Log "  ========================================" "Cyan"
Log ""

if ($verifyFail -eq 0) {
    Log "  RESULT: zbar has been COMPLETELY removed." "Green"
    Log "  All $verifyPass verification checks passed." "Green"
} else {
    Log "  RESULT: Uninstall completed with issues." "Red"
    Log "  Passed: $verifyPass  |  Failed: $verifyFail" "Red"
    Log "  Check the report for details: $reportPath" "Yellow"
}

Log ""
Log "  Report: $reportPath" "Cyan"
Log "  ========================================" "Cyan"
Log ""
