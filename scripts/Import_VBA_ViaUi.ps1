# Import AES VBA modules by driving the Outlook VBA editor UI
# (File > Import File). Use when Application.VBE is blocked even with
# AccessVBOM=1 (BUG-003 / BUG-040 on this machine).
#
# Prerequisites:
#   - Classic Outlook running
#   - Macros enabled
#   - Prefer Alt+F11 open before starting (script will try to open it)
#   - Close Excel / Word VBA editors first (or leave them; we target OUTLOOK PID only)
#
# Usage:
#   .\Import_VBA_ViaUi.ps1
#   .\Import_VBA_ViaUi.ps1 -Root "C:\GeoFooter"
#   .\Import_VBA_ViaUi.ps1 -DryRun

[CmdletBinding()]
param(
    [string]$Root = "",
    [switch]$DryRun,
    [int]$PauseMs = 550
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

# Win32 helpers: Outlook's VBE is a top-level window owned by OUTLOOK.EXE
# (Excel/Word each have their own). Matching only on title "Visual Basic"
# grabs the wrong IDE when multiple Office apps are open.
$win32 = @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public class OutlookVbeWin {
  public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);

  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc lp, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder sb, int max);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int nCmdShow);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();

  public static IntPtr FindVbeHwnd(uint outlookPid) {
    IntPtr found = IntPtr.Zero;
    EnumWindows((h, l) => {
      if (!IsWindowVisible(h)) return true;
      uint pid; GetWindowThreadProcessId(h, out pid);
      if (pid != outlookPid) return true;
      var sb = new StringBuilder(512);
      GetWindowText(h, sb, 512);
      string t = sb.ToString();
      if (t.IndexOf("Visual Basic", StringComparison.OrdinalIgnoreCase) >= 0) {
        found = h;
        return false;
      }
      return true;
    }, IntPtr.Zero);
    return found;
  }

  public static bool FocusHwnd(IntPtr h) {
    if (h == IntPtr.Zero) return false;
    if (IsIconic(h)) ShowWindow(h, 9); // SW_RESTORE
    ShowWindow(h, 5); // SW_SHOW
    return SetForegroundWindow(h);
  }

  public static string Title(IntPtr h) {
    var sb = new StringBuilder(512);
    GetWindowText(h, sb, 512);
    return sb.ToString();
  }
}
"@
Add-Type -TypeDefinition $win32 -ErrorAction Stop

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

function Get-OutlookPid {
    $ol = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $ol) { throw "Outlook is not running. Start Outlook, then re-run." }
    return [uint32]$ol.Id
}

function Get-OutlookVbeHwnd {
    param([uint32]$OutlookPid)
    return [OutlookVbeWin]::FindVbeHwnd($OutlookPid)
}

function Wait-OutlookVbe {
    param([uint32]$OutlookPid, [int]$TimeoutSec = 25)
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        $h = Get-OutlookVbeHwnd -OutlookPid $OutlookPid
        if ($h -ne [IntPtr]::Zero) { return $h }
        Start-Sleep -Milliseconds 400
    }
    return [IntPtr]::Zero
}

function Focus-OutlookVbe {
    param([uint32]$OutlookPid)
    $h = Get-OutlookVbeHwnd -OutlookPid $OutlookPid
    if ($h -eq [IntPtr]::Zero) { return $false }
    $ok = [OutlookVbeWin]::FocusHwnd($h)
    Start-Sleep -Milliseconds $PauseMs
    # Retry once — SetForegroundWindow often fails if another app just took focus
    if (-not $ok) {
        Start-Sleep -Milliseconds 300
        $ok = [OutlookVbeWin]::FocusHwnd($h)
        Start-Sleep -Milliseconds $PauseMs
    }
    return ($h -ne [IntPtr]::Zero)
}

function Ensure-OutlookVbeOpen {
    param([uint32]$OutlookPid)

    $h = Get-OutlookVbeHwnd -OutlookPid $OutlookPid
    if ($h -ne [IntPtr]::Zero) {
        Write-Host ("Using Outlook VBE: {0}" -f [OutlookVbeWin]::Title($h))
        [void](Focus-OutlookVbe -OutlookPid $OutlookPid)
        return $h
    }

    Write-Host "Opening Outlook VBA editor (Alt+F11 on Outlook)..."
    $ol = Get-Process -Id $OutlookPid
    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue
    try { [Microsoft.VisualBasic.Interaction]::AppActivate($ol.Id) | Out-Null } catch {}
    Start-Sleep -Milliseconds 600
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait("%{F11}")
    $h = Wait-OutlookVbe -OutlookPid $OutlookPid -TimeoutSec 25
    if ($h -eq [IntPtr]::Zero) {
        throw "Outlook VBA editor did not open. Close Excel/Word VBA if open, press Alt+F11 in Outlook, then re-run."
    }
    Write-Host ("Opened Outlook VBE: {0}" -f [OutlookVbeWin]::Title($h))
    [void](Focus-OutlookVbe -OutlookPid $OutlookPid)
    return $h
}

function Import-OneFile {
    param(
        [string]$Path,
        [uint32]$OutlookPid
    )

    if (-not (Focus-OutlookVbe -OutlookPid $OutlookPid)) {
        throw "Lost Outlook VBA editor (is Excel VBA in front?). Close other Office VBA windows and re-run."
    }

    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait("%f")
    Start-Sleep -Milliseconds 300
    [System.Windows.Forms.SendKeys]::SendWait("i")
    Start-Sleep -Milliseconds 800

    Set-Clipboard -Value $Path
    Start-Sleep -Milliseconds 150
    [System.Windows.Forms.SendKeys]::SendWait("^v")
    Start-Sleep -Milliseconds 280
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Start-Sleep -Milliseconds ([Math]::Max(650, $PauseMs))

    # Only dismiss if a dialog is up — Enter on the VBE itself can open things.
    # Re-focus Outlook VBE; if Import dialog still open, title often contains Import.
    [void](Focus-OutlookVbe -OutlookPid $OutlookPid)
}

function Invoke-VbeCompileAndSave {
    param([uint32]$OutlookPid)
    if (-not (Focus-OutlookVbe -OutlookPid $OutlookPid)) { return }
    Add-Type -AssemblyName System.Windows.Forms
    Write-Host "Compile (Alt+D, C)..."
    [System.Windows.Forms.SendKeys]::SendWait("%d")
    Start-Sleep -Milliseconds 280
    [System.Windows.Forms.SendKeys]::SendWait("c")
    Start-Sleep -Milliseconds 1000
    Write-Host "Save (Ctrl+S)..."
    [System.Windows.Forms.SendKeys]::SendWait("^s")
    Start-Sleep -Milliseconds 700
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
    Write-Host "DryRun: would UI-import $($ModuleFiles.Count) files into OUTLOOK VBE only."
    $ModuleFiles | ForEach-Object { Write-Host "  $_" }
    exit 0
}

$excelVba = Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object {
    $_.MainWindowTitle -match 'Visual Basic'
}
if ($excelVba) {
    Write-Host "WARNING: Excel is running. This script targets OUTLOOK's VBE only."
    Write-Host "Prefer closing Excel's VBA editor (or Excel) if focus fights occur."
    Write-Host ""
}

$outlookPid = Get-OutlookPid
Write-Host "Outlook PID: $outlookPid"
$null = Ensure-OutlookVbeOpen -OutlookPid $outlookPid

Write-Host "Importing $($ModuleFiles.Count) modules via File > Import File..."
Write-Host "Do not use the keyboard/mouse until this finishes."
Write-Host ""

$n = 0
foreach ($f in $ModuleFiles) {
    $n++
    $path = Join-Path $vbaDir $f
    Write-Host ("[{0}/{1}] {2}" -f $n, $ModuleFiles.Count, $f)
    Import-OneFile -Path $path -OutlookPid $outlookPid
}

Invoke-VbeCompileAndSave -OutlookPid $outlookPid

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
