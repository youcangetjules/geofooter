# Import AES VBA modules by driving the Outlook VBA editor UI
# (File > Import File). Use when Application.VBE is blocked even with
# AccessVBOM=1 (BUG-003 / BUG-040 on this machine).
#
# Prerequisites:
#   - Classic Outlook running
#   - Macros enabled
#   - Prefer Alt+F11 open before starting (script will try to open it)
#
# Usage:
#   .\Import_VBA_ViaUi.ps1
#   .\Import_VBA_ViaUi.ps1 -Root "C:\GeoFooter"
#   .\Import_VBA_ViaUi.ps1 -DryRun

[CmdletBinding()]
param(
    [string]$Root = "",
    [switch]$DryRun,
    [int]$PauseMs = 450
)

$ErrorActionPreference = "Stop"

$ModuleFiles = @(
    "MSCANPaths.bas",
    "MSCANModLogging.bas",
    "MSCANHealth.bas",
    "MSCANIdle.bas",
    "MSCANModQueueManager.bas",
    "MSCANModWatchers.bas",
    "MSCANModStatus.bas",
    "MSCANSettings.bas",
    "MSCANAppBootstrap.bas",
    "MSCANClassificationDialog.bas",
    "MSCANReadNotify.bas",
    "MSCANPixel.bas",
    "MSCANToolbar.bas",
    "MSCANModSenderRules.bas",
    "MSCANModule1.bas",
    "MSCANCore.bas",
    "MSCANEventHandlers.bas",
    "MSCANDiagnostics.bas",
    "MSCANEmailWatcher.cls",
    "MSCANSelfUpdate.bas"
)

function Resolve-InstallRoot {
    param([string]$Hint)
    $candidates = @()
    if ($Hint) { $candidates += $Hint.Trim().Trim('"') }
    if ($env:GEOFOOTER_ROOT) { $candidates += $env:GEOFOOTER_ROOT.Trim().Trim('"') }
    $ptr = Join-Path $env:LOCALAPPDATA "GeoFooter\install_root.txt"
    if (Test-Path $ptr) {
        $line = (Get-Content $ptr -TotalCount 1 -ErrorAction SilentlyContinue)
        if ($line) { $candidates += $line.Trim().Trim('"') }
    }
    $candidates += (Split-Path -Parent $PSScriptRoot)
    foreach ($c in $candidates) {
        if (-not $c) { continue }
        $p = $c.TrimEnd("\", "/")
        if ((Test-Path (Join-Path $p "VBA")) -and (Test-Path (Join-Path $p "VERSION"))) {
            return (Resolve-Path $p).Path
        }
    }
    throw "Could not resolve install root. Pass -Root or set install_root.txt."
}

function Get-VbeProcess {
    Get-Process -ErrorAction SilentlyContinue |
        Where-Object {
            $_.MainWindowHandle -ne 0 -and
            $_.MainWindowTitle -match 'Visual Basic'
        } |
        Select-Object -First 1
}

function Wait-Vbe {
    param([int]$TimeoutSec = 20)
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        $p = Get-VbeProcess
        if ($p) { return $p }
        Start-Sleep -Milliseconds 400
    }
    return $null
}

function Focus-Vbe {
    $p = Get-VbeProcess
    if (-not $p) { return $false }
    try {
        Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue
        [Microsoft.VisualBasic.Interaction]::AppActivate($p.Id) | Out-Null
    } catch {
        try { [Microsoft.VisualBasic.Interaction]::AppActivate($p.MainWindowTitle) | Out-Null } catch { return $false }
    }
    Start-Sleep -Milliseconds $PauseMs
    return $true
}

function Ensure-VbeOpen {
    $p = Get-VbeProcess
    if ($p) { return $p }

    $ol = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $ol) { throw "Outlook is not running. Start Outlook, then re-run." }

    Write-Host "Opening VBA editor (Alt+F11)..."
    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue
    try { [Microsoft.VisualBasic.Interaction]::AppActivate($ol.Id) | Out-Null } catch {}
    Start-Sleep -Milliseconds 500
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait("%{F11}")
    $p = Wait-Vbe -TimeoutSec 25
    if (-not $p) {
        throw "VBA editor did not open. Press Alt+F11 in Outlook, then re-run this script."
    }
    return $p
}

function Import-OneFile {
    param([string]$Path)

    if (-not (Focus-Vbe)) { throw "Lost VBA editor focus." }

    Add-Type -AssemblyName System.Windows.Forms
    # File > Import File (English VBE). Paste path — avoids SendKeys escaping.
    [System.Windows.Forms.SendKeys]::SendWait("%f")
    Start-Sleep -Milliseconds 280
    [System.Windows.Forms.SendKeys]::SendWait("i")
    Start-Sleep -Milliseconds 700

    Set-Clipboard -Value $Path
    Start-Sleep -Milliseconds 120
    [System.Windows.Forms.SendKeys]::SendWait("^v")
    Start-Sleep -Milliseconds 250
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Start-Sleep -Milliseconds ([Math]::Max(500, $PauseMs))

    # Dismiss unexpected message boxes (duplicate name, etc.)
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Start-Sleep -Milliseconds 200
}

function Invoke-VbeCompileAndSave {
    if (-not (Focus-Vbe)) { return }
    Add-Type -AssemblyName System.Windows.Forms
    Write-Host "Compile (Alt+D, C)..."
    [System.Windows.Forms.SendKeys]::SendWait("%d")
    Start-Sleep -Milliseconds 250
    [System.Windows.Forms.SendKeys]::SendWait("c")
    Start-Sleep -Milliseconds 900
    # If a compile-error dialog appeared, leave it for the user.
    Write-Host "Save (Ctrl+S)..."
    [System.Windows.Forms.SendKeys]::SendWait("^s")
    Start-Sleep -Milliseconds 600
}

# --- main ---
Add-Type -AssemblyName System.Windows.Forms

$rootPath = Resolve-InstallRoot -Hint $Root
$vbaDir = Join-Path $rootPath "VBA"
Write-Host "Install root: $rootPath"
Write-Host "VBA folder:   $vbaDir"
Write-Host ""

$missing = @()
foreach ($f in $ModuleFiles) {
    if (-not (Test-Path (Join-Path $vbaDir $f))) { $missing += $f }
}
if ($missing.Count) {
    throw "Missing module file(s): $($missing -join ', ')"
}

if ($DryRun) {
    Write-Host "DryRun: would UI-import $($ModuleFiles.Count) files."
    $ModuleFiles | ForEach-Object { Write-Host "  $_" }
    exit 0
}

$null = Ensure-VbeOpen
Write-Host "Importing $($ModuleFiles.Count) modules via File > Import File..."
Write-Host "Do not use the keyboard/mouse until this finishes."
Write-Host ""

$n = 0
foreach ($f in $ModuleFiles) {
    $n++
    $path = Join-Path $vbaDir $f
    Write-Host ("[{0}/{1}] {2}" -f $n, $ModuleFiles.Count, $f)
    Import-OneFile -Path $path
}

Invoke-VbeCompileAndSave

$tos = Join-Path $vbaDir "ThisOutlookSession.cls"
Write-Host ""
Write-Host "DONE with MSCAN* modules."
Write-Host "Still required once:"
Write-Host "  1. Open $tos"
Write-Host "  2. Copy its code into Microsoft Outlook Objects > ThisOutlookSession"
Write-Host "  3. Remove any duplicate ThisOutlookSession under Modules"
Write-Host "  4. Debug > Compile, File > Save"
Write-Host "  5. Fully quit Outlook (tray) and reopen"
Write-Host "  6. Alt+F8 > RecoverAesUi"
if (Test-Path $tos) {
    try { Start-Process notepad.exe -ArgumentList $tos } catch {}
}
exit 0
