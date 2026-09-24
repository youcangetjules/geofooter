# Sync AES VBA modules from disk into the running Outlook VBA project.
# Replaces File > Import File for all MSCAN*.bas / MSCANEmailWatcher.cls.
#
# Prerequisites:
#   - Outlook classic (desktop) is running
#   - Trust access to the VBA project object model (AccessVBOM=1)
#     This script can set that registry value with -EnableAccessVBOM
#
# Usage:
#   .\Import_VBA_to_Outlook.ps1
#   .\Import_VBA_to_Outlook.ps1 -SyncThisOutlookSession
#   .\Import_VBA_to_Outlook.ps1 -DryRun
#   .\Import_VBA_to_Outlook.ps1 -Root "D:\Suite\GeoFooter"

[CmdletBinding()]
param(
    [string]$Root = "",
    [switch]$EnableAccessVBOM,
    [switch]$SyncThisOutlookSession,
    [switch]$DryRun,
    [switch]$SkipCleanup
)

$ErrorActionPreference = "Stop"

# Import order matches VBA\IMPORT.txt (ThisOutlookSession handled separately).
$ModuleFiles = @(
    "MSCANPaths.bas",
    "MSCANModLogging.bas",
    "MSCANModQueueManager.bas",
    "MSCANModWatchers.bas",
    "MSCANModStatus.bas",
    "MSCANSettings.bas",
    "MSCANAppBootstrap.bas",
    "MSCANClassificationDialog.bas",
    "MSCANReadNotify.bas",
    "MSCANToolbar.bas",
    "MSCANModSenderRules.bas",
    "MSCANModule1.bas",
    "MSCANCore.bas",
    "MSCANEventHandlers.bas",
    "MSCANDiagnostics.bas",
    "MSCANEmailWatcher.cls",
    "MSCANSelfUpdate.bas"
)

$StaleNames = @(
    "UserForm1",
    "frmClassification",
    "MSCANSecurityRatingForm",
    "MSCANmodLogging",
    "MSCANmodQueueManager",
    "MSCANmodWatchers",
    "MSCANclsEmailWatcher"
)

function Resolve-InstallRoot {
    param([string]$Hint)

    $candidates = @()
    if ($Hint) { $candidates += $Hint.Trim().Trim('"') }
    $envRoot = ($env:GEOFOOTER_ROOT -replace '"', '').Trim()
    if ($envRoot) { $candidates += $envRoot }

    $ptr = Join-Path $env:LOCALAPPDATA "GeoFooter\install_root.txt"
    if (Test-Path $ptr) {
        $line = (Get-Content $ptr -TotalCount 1 -ErrorAction SilentlyContinue)
        if ($line) { $candidates += ($line.Trim().Trim('"')) }
    }

    $scriptParent = Split-Path -Parent $PSScriptRoot
    $candidates += $scriptParent

    foreach ($c in $candidates) {
        if (-not $c) { continue }
        $p = $c.TrimEnd("\", "/")
        if ((Test-Path (Join-Path $p "VBA")) -and (
                (Test-Path (Join-Path $p "VERSION")) -or
                (Test-Path (Join-Path $p "aes\geolocate_headers.py"))
            )) {
            return (Resolve-Path $p).Path
        }
    }
    throw "Could not resolve install root. Pass -Root or set GEOFOOTER_ROOT / GURI Install root."
}

function Get-OutlookOfficeVersionKey {
    param($OutlookApp)
    try {
        $ver = [string]$OutlookApp.Version
        if ($ver -match "^(\d+)\.") {
            return "$($Matches[1]).0"
        }
    } catch {}
    return "16.0"
}

function Ensure-AccessVBOM {
    param(
        [string]$OfficeVer,
        [switch]$Enable
    )
    $key = "HKCU:\Software\Microsoft\Office\$OfficeVer\Outlook\Security"
    $cur = $null
    try {
        $cur = (Get-ItemProperty -Path $key -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
    } catch {}

    if ($cur -eq 1) {
        Write-Host "AccessVBOM: already enabled (Office $OfficeVer)."
        return $true
    }

    if (-not $Enable) {
        Write-Host "AccessVBOM: NOT set (current=$cur)."
        Write-Host "  Re-run with -EnableAccessVBOM, or set:"
        Write-Host "  $key  AccessVBOM = 1 (DWORD)"
        Write-Host "  Then fully quit Outlook (tray too) and reopen before syncing."
        return $false
    }

    if ($DryRun) {
        Write-Host "DryRun: would set AccessVBOM=1 under $key"
        return $false
    }

    New-Item -Path $key -Force | Out-Null
    Set-ItemProperty -Path $key -Name AccessVBOM -Type DWord -Value 1
    Write-Host "AccessVBOM: set to 1 under $key"
    Write-Host "IMPORTANT: fully quit Outlook (tray too), reopen, then run this script again."
    return $false
}

function Test-OutlookProcessRunning {
    return $null -ne (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
}

function Test-HasGetActiveObject {
    $flags = [Reflection.BindingFlags]::Public -bor [Reflection.BindingFlags]::Static
    return $null -ne [Runtime.InteropServices.Marshal].GetMethod("GetActiveObject", $flags)
}

function Test-IsElevated {
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        $p = New-Object Security.Principal.WindowsPrincipal($id)
        return [bool]$p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function Get-OutlookAttachAttempts {
    # Returns @{ App = $comOrNull; Errors = @(strings) }
    $errors = New-Object System.Collections.Generic.List[string]
    $app = $null
    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = "Continue"

    try {
        # 1) ROT via Marshal.GetActiveObject (Windows PowerShell / .NET Framework)
        try {
            $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
            if ($app) {
                return @{ App = $app; Errors = @() }
            }
            $errors.Add("GetActiveObject: returned null")
        } catch {
            $errors.Add("GetActiveObject: $($_.Exception.Message)")
        }

        # 2) VB GetObject(, "Outlook.Application")
        try {
            Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
            $app = [Microsoft.VisualBasic.Interaction]::GetObject($null, "Outlook.Application")
            if ($app) {
                return @{ App = $app; Errors = $errors.ToArray() }
            }
            $errors.Add("VB.GetObject: returned null")
        } catch {
            $errors.Add("VB.GetObject: $($_.Exception.Message)")
        }

        # 3) CreateObject - on Windows PowerShell this often binds to the running instance
        try {
            $app = New-Object -ComObject Outlook.Application
            if ($app) {
                return @{ App = $app; Errors = $errors.ToArray() }
            }
            $errors.Add("New-Object: returned null")
        } catch {
            $errors.Add("New-Object Outlook.Application: $($_.Exception.Message)")
        }

        return @{ App = $null; Errors = $errors.ToArray() }
    } finally {
        $ErrorActionPreference = $prevEap
    }
}

function Get-OutlookApplication {
    $result = Get-OutlookAttachAttempts
    if ($result.App) {
        $ver = ""
        try { $ver = [string]$result.App.Version } catch {}
        if ($ver) { Write-Host "Attached to Outlook $ver via COM." }
        else { Write-Host "Attached to Outlook via COM." }
        return $result.App
    }

    $running = Test-OutlookProcessRunning
    $outlookPath = ""
    try {
        $outlookPath = (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue |
            Select-Object -First 1 -ExpandProperty Path)
    } catch {}

    if ($running) {
        $errText = if ($result.Errors.Count) { ($result.Errors -join "`n  ") } else { "(no detail)" }
        $elev = Test-IsElevated
        $elevationHint = ""
        if ($elev) {
            $elevationHint = (
                "`n*** This PowerShell is ELEVATED (Administrator) but Outlook usually is not.`n" +
                "*** Close this window and re-run the bat from a normal (non-admin) prompt.`n"
            )
        }
        throw (
            "Outlook.exe is running but COM attach failed from this host.`n" +
            "  Host: Windows PowerShell $($PSVersionTable.PSVersion) | 64-bit process=$([Environment]::Is64BitProcess) | elevated=$elev`n" +
            "  Outlook: $outlookPath`n" +
            $elevationHint +
            "  Attempts:`n  $errText`n" +
            "Common fixes:`n" +
            "  - Start Outlook normally (not 'Run as administrator'), then re-run this bat (also not elevated) - or elevate BOTH.`n" +
            "  - Press Alt+F11 once in Outlook, close the VBA window, re-run.`n" +
            "  - Fully quit Outlook (tray too), start it again, wait until the inbox loads, re-run.`n" +
            "  - Your interactive shell can be PowerShell 7; this bat must keep using SysWOW64 Windows PowerShell 5.1."
        )
    }

    Write-Host "Outlook not running - starting it..."
    try {
        $app = New-Object -ComObject Outlook.Application
    } catch {
        throw (
            "Failed to start Outlook via COM ($($_.Exception.Message)).`n" +
            "Start Outlook manually, then re-run this script."
        )
    }
    $null = $app.GetNamespace("MAPI")
    Start-Sleep -Seconds 2
    return $app
}

function Get-WindowsPowerShell32 {
    $p = Join-Path $env:SystemRoot "SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    if (Test-Path $p) { return $p }
    return $null
}

function Ensure-CompatiblePowerShellHost {
    # Outlook AES installs are almost always 32-bit Office. PowerShell 7 also
    # lacks Marshal.GetActiveObject. Re-launch once under WinPS 5.1 x86.
    if ($env:GEOFOOTER_VBA_IMPORT_RELAUNCHED -eq "1") { return }

    $needRelaunch = $false
    $reason = ""
    if (-not (Test-HasGetActiveObject)) {
        $needRelaunch = $true
        $reason = "this host has no Marshal.GetActiveObject (likely PowerShell 7+)"
    } elseif ([Environment]::Is64BitProcess -and (Get-WindowsPowerShell32)) {
        $outlookPath = $null
        try {
            $outlookPath = (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue |
                Select-Object -First 1 -ExpandProperty Path)
        } catch {}
        if ($outlookPath -and ($outlookPath -match '(?i)Program Files \(x86\)')) {
            $needRelaunch = $true
            $reason = "Outlook is 32-bit; current PowerShell is 64-bit"
        }
    }

    if (-not $needRelaunch) { return }

    $ps32 = Get-WindowsPowerShell32
    if (-not $ps32) {
        Write-Host "WARNING: need Windows PowerShell 5.1 x86 but SysWOW64 powershell.exe not found ($reason)."
        return
    }

    Write-Host "Re-launching under 32-bit Windows PowerShell 5.1 ($reason)..."
    $env:GEOFOOTER_VBA_IMPORT_RELAUNCHED = "1"
    $argList = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $PSCommandPath
    )
    if ($Root) { $argList += @("-Root", $Root) }
    if ($EnableAccessVBOM) { $argList += "-EnableAccessVBOM" }
    if ($SyncThisOutlookSession) { $argList += "-SyncThisOutlookSession" }
    if ($DryRun) { $argList += "-DryRun" }
    if ($SkipCleanup) { $argList += "-SkipCleanup" }

    $p = Start-Process -FilePath $ps32 -ArgumentList $argList -Wait -PassThru -NoNewWindow
    exit $p.ExitCode
}

function Get-VbaProject {
    param($OutlookApp)

    $vbeErrors = New-Object System.Collections.Generic.List[string]
    $vbe = $null

    try {
        $vbe = $OutlookApp.VBE
        if (-not $vbe) { $vbeErrors.Add("Application.VBE returned null/empty") }
    } catch {
        $vbeErrors.Add("Application.VBE: $($_.Exception.Message)")
    }

    if (-not $vbe) {
        try {
            $vbe = $OutlookApp.Application.VBE
            if (-not $vbe) { $vbeErrors.Add("Application.Application.VBE returned null/empty") }
        } catch {
            $vbeErrors.Add("Application.Application.VBE: $($_.Exception.Message)")
        }
    }

    if (-not $vbe) {
        # Outlook often refuses Application.VBE until the IDE has been opened once.
        try {
            Write-Host "Nudge: sending Alt+F11 to Outlook to load the VBA project..."
            $shell = New-Object -ComObject WScript.Shell
            $activated = $false
            foreach ($title in @("Inbox", "Outlook", "Mail")) {
                if ($shell.AppActivate($title)) { $activated = $true; break }
            }
            if (-not $activated) {
                $oid = (Get-Process OUTLOOK -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty Id)
                if ($oid) { [void]$shell.AppActivate($oid) }
            }
            Start-Sleep -Milliseconds 400
            $shell.SendKeys("%{F11}")
            Start-Sleep -Seconds 2
        } catch {
            Write-Host "Nudge failed: $($_.Exception.Message)"
        }

        try {
            $vbe = $OutlookApp.VBE
            if (-not $vbe) { $vbeErrors.Add("Application.VBE after Alt+F11 returned null/empty") }
        } catch {
            $vbeErrors.Add("Application.VBE after Alt+F11: $($_.Exception.Message)")
        }
    }

    if (-not $vbe) {
        $access = $null
        try {
            $access = (Get-ItemProperty "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security" -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
        } catch {}
        $detail = ($vbeErrors -join "`n  ")
        $restartHint = ""
        if ($access -eq 1) {
            $restartHint = (
                "`nAccessVBOM is already 1 in the registry, but this Outlook process still blocks VBE.`n" +
                "That almost always means Outlook was running when AccessVBOM was set.`n" +
                "Fully quit Outlook (system tray too), start it again, press Alt+F11 once, then re-run this bat.`n"
            )
        }
        throw (
            "Cannot open the Outlook VBA project.`n" +
            "  AccessVBOM (HKCU)= $access`n" +
            $restartHint +
            "  VBE errors:`n  $detail`n" +
            "Fix:`n" +
            "  1. Fully quit Outlook (tray icon too) and reopen`n" +
            "  2. Press Alt+F11 once (VBA editor can then be closed)`n" +
            "  3. Re-run scripts\Import_VBA_to_Outlook.bat from a non-admin prompt`n" +
            "If it still fails, update from inside Outlook instead:`n" +
            "  Alt+F11 > File > Import File > VBA\MSCANSelfUpdate.bas (one time)`n" +
            "  Alt+F8 > UpdateAesFromDisk > Run, then restart Outlook"
        )
    }

    try { $null = $vbe.MainWindow } catch {}

    $project = $null
    try { $project = $vbe.ActiveVBProject } catch {}
    if (-not $project) {
        try {
            if ($vbe.VBProjects.Count -ge 1) {
                $project = $vbe.VBProjects.Item(1)
            }
        } catch {}
    }
    if (-not $project) {
        foreach ($p in @($vbe.VBProjects)) {
            if ($p.Name -match '(?i)VbaProject|Project') {
                $project = $p
                break
            }
        }
    }
    if (-not $project) {
        throw "No VBA project found after opening VBE. Press Alt+F11 in Outlook, then re-run."
    }
    return $project
}

function Remove-ComponentByName {
    param($Project, [string]$Name)
    foreach ($comp in @($Project.VBComponents)) {
        if ($comp.Name -eq $Name) {
            if ($DryRun) {
                Write-Host "  DryRun: remove $Name (type $($comp.Type))"
            } else {
                $Project.VBComponents.Remove($comp)
                Write-Host "  Removed $Name"
            }
            return $true
        }
    }
    return $false
}

function Import-ModuleFile {
    param($Project, [string]$FilePath)
    $leaf = Split-Path -Leaf $FilePath
    $name = [IO.Path]::GetFileNameWithoutExtension($leaf)

    if (-not (Test-Path $FilePath)) {
        throw "Missing module file: $FilePath"
    }

    # Remove existing standard/class module with this name (not the built-in document).
    foreach ($comp in @($Project.VBComponents)) {
        if ($comp.Name -ne $name) { continue }
        # Type 100 = document (ThisOutlookSession) - never remove that here.
        if ([int]$comp.Type -eq 100) {
            Write-Host "  Skip remove of document component $name"
            continue
        }
        if ($DryRun) {
            Write-Host "  DryRun: remove existing $name before import"
        } else {
            $Project.VBComponents.Remove($comp)
        }
    }

    if ($DryRun) {
        Write-Host "  DryRun: import $leaf"
        return
    }

    $Project.VBComponents.Import($FilePath) | Out-Null
    Write-Host "  Imported $leaf"
}

function Get-ClsBodyWithoutHeader {
    param([string]$ClsPath)
    $lines = Get-Content -LiteralPath $ClsPath -Encoding UTF8
    $out = New-Object System.Collections.Generic.List[string]
    $inBegin = $false
    foreach ($line in $lines) {
        if ($line -match '^\s*VERSION\s+') { continue }
        if ($line -match '^\s*BEGIN\s*$') { $inBegin = $true; continue }
        if ($inBegin) {
            if ($line -match '^\s*END\s*$') { $inBegin = $false }
            continue
        }
        if ($line -match '^\s*Attribute\s+VB_') { continue }
        $out.Add($line)
    }
    # Drop leading blank lines
    while ($out.Count -gt 0 -and [string]::IsNullOrWhiteSpace($out[0])) {
        $out.RemoveAt(0)
    }
    return ($out -join "`r`n")
}

function Sync-ThisOutlookSessionDocument {
    param($Project, [string]$ClsPath)

    $doc = $null
    foreach ($comp in @($Project.VBComponents)) {
        if ($comp.Name -eq "ThisOutlookSession" -and [int]$comp.Type -eq 100) {
            $doc = $comp
            break
        }
    }
    if (-not $doc) {
        Write-Host "WARNING: built-in ThisOutlookSession document not found; skip sync."
        return
    }

    $body = Get-ClsBodyWithoutHeader -ClsPath $ClsPath
    if ([string]::IsNullOrWhiteSpace($body)) {
        throw "ThisOutlookSession.cls produced empty body after stripping headers."
    }

    if ($DryRun) {
        Write-Host "  DryRun: replace built-in ThisOutlookSession CodeModule ($($body.Length) chars)"
        return
    }

    $cm = $doc.CodeModule
    if ($cm.CountOfLines -gt 0) {
        $cm.DeleteLines(1, $cm.CountOfLines)
    }
    $cm.AddFromString($body)
    Write-Host "  Synced built-in ThisOutlookSession from disk"
}

#----- main -----
Ensure-CompatiblePowerShellHost

Write-Host "========================================"
Write-Host "  Import AES VBA into Outlook"
Write-Host "========================================"
Write-Host ""
Write-Host ("PowerShell: {0} | 64-bit process: {1}" -f $PSVersionTable.PSVersion, [Environment]::Is64BitProcess)

$installRoot = Resolve-InstallRoot -Hint $Root
$vbaDir = Join-Path $installRoot "VBA"
Write-Host "Install root: $installRoot"
Write-Host "VBA folder:   $vbaDir"
Write-Host ""

if (-not (Test-Path $vbaDir)) {
    throw "VBA folder not found: $vbaDir"
}

$outlook = Get-OutlookApplication
$officeVer = Get-OutlookOfficeVersionKey -OutlookApp $outlook
Write-Host "Outlook: $($outlook.Version) (registry key $officeVer)"

$ok = Ensure-AccessVBOM -OfficeVer $officeVer -Enable:$EnableAccessVBOM
if (-not $ok) {
    if ($EnableAccessVBOM) {
        exit 2
    }
    # Still try - maybe AccessVBOM was set another way / Excel-style trust.
}

$project = Get-VbaProject -OutlookApp $outlook
Write-Host "VBProject: $($project.Name)"
Write-Host ""

if (-not $SkipCleanup) {
    Write-Host "Cleaning stale / duplicate components..."
    foreach ($n in $StaleNames) {
        [void](Remove-ComponentByName -Project $project -Name $n)
    }
    # Imported class copies of ThisOutlookSession (not the built-in document).
    foreach ($comp in @($project.VBComponents)) {
        $n = $comp.Name
        if ($n -match '^ThisOutlookSession\d*$' -and [int]$comp.Type -ne 100) {
            if ($DryRun) {
                Write-Host "  DryRun: remove duplicate class $n"
            } else {
                $project.VBComponents.Remove($comp)
                Write-Host "  Removed duplicate class $n"
            }
        }
    }
    Write-Host ""
}

Write-Host "Importing modules (IMPORT.txt order)..."
foreach ($file in $ModuleFiles) {
    $path = Join-Path $vbaDir $file
    Import-ModuleFile -Project $project -FilePath $path
}
Write-Host ""

if ($SyncThisOutlookSession) {
    Write-Host "Syncing built-in ThisOutlookSession..."
    $cls = Join-Path $vbaDir "ThisOutlookSession.cls"
    if (-not (Test-Path $cls)) {
        throw "Missing $cls"
    }
    Sync-ThisOutlookSessionDocument -Project $project -ClsPath $cls
    Write-Host ""
} else {
    Write-Host "ThisOutlookSession: left unchanged (pass -SyncThisOutlookSession to overwrite the built-in object)."
    Write-Host ""
}

function Invoke-VbeCommand {
    param($Vbe, [int]$ControlId, [string]$Label)
    try {
        $ctl = $Vbe.CommandBars.FindControl(1, $ControlId)
        if (-not $ctl) {
            Write-Host "  ${Label}: VBE control $ControlId not found - do it manually."
            return $false
        }
        $ctl.Execute()
        Write-Host "  ${Label}: done"
        return $true
    } catch {
        Write-Host "  ${Label}: failed ($($_.Exception.Message)) - do it manually."
        return $false
    }
}

$saved = $false
if (-not $DryRun) {
    # Without a save, Outlook reloads the old VbaProject.OTM on the next restart.
    Write-Host "Compiling and saving the VBA project..."
    $vbe = $project.VBE
    [void](Invoke-VbeCommand -Vbe $vbe -ControlId 578 -Label "Debug > Compile")
    $saved = Invoke-VbeCommand -Vbe $vbe -ControlId 3 -Label "File > Save"
    Write-Host ""
}

Write-Host "Done."
if (-not $saved) {
    Write-Host "IMPORTANT: Alt+F11 -> Debug -> Compile VBAProject, then File -> Save (Ctrl+S)."
    Write-Host "Without Save, restarting Outlook brings the OLD code back."
}
Write-Host "Then fully quit Outlook (tray too) and reopen so the new modules run from startup."
exit 0
