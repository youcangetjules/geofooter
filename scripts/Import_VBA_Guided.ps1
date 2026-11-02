# Norton-safe guided import: NO SendKeys / NO UI automation.
# Puts each module path on the clipboard and waits for you to
# File > Import File in Outlook's VBA editor (Alt+F11).
#
# Usage:
#   .\Import_VBA_Guided.ps1
#   .\Import_VBA_Guided.ps1 -Root "C:\GeoFooter"

[CmdletBinding()]
param([string]$Root = "")

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
    throw "Could not resolve install root."
}

Add-Type -AssemblyName System.Windows.Forms

$rootPath = Resolve-InstallRoot -Hint $Root
$vbaDir = Join-Path $rootPath "VBA"
explorer.exe $vbaDir

$r = [System.Windows.Forms.MessageBox]::Show(
    "Norton-safe guided import (no keystroke injection).`r`n`r`n" +
    "1. In Outlook press Alt+F11 (VBA editor).`r`n" +
    "2. Remove any partial MSCAN* modules from a failed run.`r`n" +
    "3. Click OK here, then for EACH file:`r`n" +
    "     File > Import File > Ctrl+V > Enter`r`n" +
    "4. Path is placed on the clipboard automatically.`r`n`r`n" +
    "Close Excel VBA if it is open. Continue?",
    "AES guided VBA import",
    [System.Windows.Forms.MessageBoxButtons]::OKCancel,
    [System.Windows.Forms.MessageBoxIcon]::Information
)
if ($r -ne [System.Windows.Forms.DialogResult]::OK) { exit 1 }

$n = 0
foreach ($f in $ModuleFiles) {
    $n++
    $path = Join-Path $vbaDir $f
    if (-not (Test-Path $path)) { throw "Missing $path" }
    Set-Clipboard -Value $path
    $msg = "[{0}/{1}] {2}`r`n`r`nClipboard has:`r`n{3}`r`n`r`nIn Outlook VBA: File > Import File > Ctrl+V > Enter`r`n`r`nClick OK when that file is imported.`r`nCancel to stop." -f $n, $ModuleFiles.Count, $f, $path
    $r = [System.Windows.Forms.MessageBox]::Show(
        $msg,
        "AES import step",
        [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::None
    )
    if ($r -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "Stopped at $f"
        exit 2
    }
}

$tos = Join-Path $vbaDir "ThisOutlookSession.cls"
Set-Clipboard -Value (Get-Content $tos -Raw)
[System.Windows.Forms.MessageBox]::Show(
    "All MSCAN modules imported.`r`n`r`n" +
    "ThisOutlookSession.cls is now on the clipboard.`r`n" +
    "Paste into: Microsoft Outlook Objects > ThisOutlookSession`r`n" +
    "(replace existing). Remove any duplicate under Modules.`r`n`r`n" +
    "Then: Debug > Compile > File > Save > quit Outlook fully > reopen > Alt+F8 > RecoverAesUi",
    "AES import almost done",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)
if (Test-Path $tos) { Start-Process notepad.exe -ArgumentList $tos }
Write-Host "Guided import finished."
exit 0
