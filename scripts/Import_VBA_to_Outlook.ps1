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
    "MSCANEmailWatcher.cls"
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

function Get-OutlookApplication {
    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        Write-Host "Outlook not running - starting it..."
        $app = New-Object -ComObject Outlook.Application
        # Force MAPI init
        $null = $app.GetNamespace("MAPI")
        Start-Sleep -Seconds 2
        return $app
    }
}

function Get-VbaProject {
    param($OutlookApp)
    try {
        return $OutlookApp.Application.VBE.ActiveVBProject
    } catch {
        $detail = $_.Exception.Message
        throw (
            "Cannot open the Outlook VBA project (AccessVBOM / Trust access).`n" +
            "Fix:`n" +
            "  1. Run:  .\Import_VBA_to_Outlook.ps1 -EnableAccessVBOM`n" +
            "  2. Fully quit Outlook (including tray) and reopen`n" +
            "  3. Run this script again without -EnableAccessVBOM`n" +
            "Underlying error: $detail"
        )
    }
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
Write-Host "========================================"
Write-Host "  Import AES VBA into Outlook"
Write-Host "========================================"
Write-Host ""

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

Write-Host "Done."
Write-Host "Next: Alt+F11 -> Debug -> Compile VBAProject, then File -> Save."
Write-Host "If Outlook was already running with old code in memory, restart Outlook once."
exit 0
