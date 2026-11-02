# Preflight for AES VBA restore. Does not import anything.
# Checks install root, disk modules, AccessVBOM, Outlook/Excel/VBE, and prints
# the exact steps before Import_VBA_ViaUi / Guided.
#
# Usage:
#   .\Prepare_VBA_Restore.ps1
#   .\Prepare_VBA_Restore.ps1 -Root "C:\GeoFooter"

[CmdletBinding()]
param([string]$Root = "")

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms

$win32 = @"
using System; using System.Text; using System.Runtime.InteropServices;
public class PrepVbe {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc lp, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder sb, int max);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  public static string Find(uint pid) {
    string found = "";
    EnumWindows((h,l) => {
      if (!IsWindowVisible(h)) return true;
      uint p; GetWindowThreadProcessId(h, out p);
      if (p != pid) return true;
      var sb = new StringBuilder(512); GetWindowText(h, sb, 512);
      string t = sb.ToString();
      if (t.IndexOf("Visual Basic", StringComparison.OrdinalIgnoreCase) >= 0) { found = t; return false; }
      return true;
    }, IntPtr.Zero);
    return found;
  }
}
"@
Add-Type -TypeDefinition $win32 -ErrorAction SilentlyContinue

function Resolve-InstallRoot([string]$Hint) {
    $candidates = @()
    if ($Hint) { $candidates += $Hint.Trim().Trim('"') }
    if ($env:GEOFOOTER_ROOT) { $candidates += $env:GEOFOOTER_ROOT.Trim().Trim('"') }
    $ptr = Join-Path $env:LOCALAPPDATA "GeoFooter\install_root.txt"
    if (Test-Path $ptr) {
        $line = (Get-Content $ptr -TotalCount 1 -EA SilentlyContinue)
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

$root = Resolve-InstallRoot $Root
$vba = Join-Path $root "VBA"
$scripts = Join-Path $root "scripts"
$sec = Get-ItemProperty "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security" -EA SilentlyContinue
$ol = Get-Process OUTLOOK -EA SilentlyContinue | Select-Object -First 1
$excelWb = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -EA SilentlyContinue |
    Where-Object { $_.CommandLine -match '\.xls' })
$vbeTitle = if ($ol) { [PrepVbe]::Find([uint32]$ol.Id) } else { "" }

$checks = New-Object System.Collections.Generic.List[string]
$ok = $true
function Add-Check($pass, $msg) {
    if ($pass) { $checks.Add("[OK]  $msg") }
    else { $script:ok = $false; $checks.Add("[!!] $msg") }
}

Add-Check (Test-Path $vba) "Install root VBA folder: $vba"
Add-Check ($sec.AccessVBOM -eq 1) "AccessVBOM=$($sec.AccessVBOM) (need 1)"
Add-Check ($sec.Level -eq 1 -or $null -eq $sec.Level) "Macro Level=$($sec.Level) (1 = enable all)"
Add-Check ($null -ne $ol) "Outlook running"
Add-Check ([string]::IsNullOrEmpty($vbeTitle) -eq $false) $(if ($vbeTitle) { "Outlook VBE open: $vbeTitle" } else { "Outlook VBE not open — press Alt+F11 in Outlook" })
Add-Check ($excelWb.Count -eq 0) $(if ($excelWb.Count) { "Excel workbook still open ($($excelWb.Count)) — close Excel fully" } else { "No Excel workbooks open" })
Add-Check (Test-Path (Join-Path $scripts "Import_VBA_ViaUi.bat")) "ViaUi importer present"
Add-Check (Test-Path (Join-Path $scripts "Import_VBA_Guided.bat")) "Guided importer present"

$importerLive = @(Get-CimInstance Win32_Process -EA SilentlyContinue |
    Where-Object { $_.CommandLine -match 'Import_VBA_(ViaUi|Guided)' })
Add-Check ($importerLive.Count -eq 0) $(if ($importerLive.Count) { "Importer already running — stop it first" } else { "No importer currently running" })

$text = @"
AES VBA restore — set up first

$($checks -join "`r`n")

Before running the importer:
  1. Norton exception already set for GeoFooter (scripts + VBA).
  2. In Outlook VBA Project Explorer: remove EVERY MSCAN* module
     from failed attempts (right-click > Remove > No).
     Leave Microsoft Outlook Objects > ThisOutlookSession alone.
  3. Close Excel completely (not just the window — check Task Manager).
  4. Outlook open + Alt+F11 so VBE title shows VbaProject.OTM.
  5. Hands off keyboard until ViaUi finishes.

Then run:
  $scripts\Import_VBA_ViaUi.bat

If Norton blocks SendKeys again:
  $scripts\Import_VBA_Guided.bat

After import:
  Paste VBA\ThisOutlookSession.cls into built-in ThisOutlookSession
  Debug > Compile > File > Save
  Quit Outlook (tray too) > reopen > Alt+F8 > RecoverAesUi
  Then set AesRibbonHost LoadBehavior=3 if you want the ribbon back.
"@

Write-Host $text
[System.Windows.Forms.MessageBox]::Show(
    $text,
    $(if ($ok) { "AES restore — ready when Project Explorer is clean" } else { "AES restore — fix the [!!] items first" }),
    [System.Windows.Forms.MessageBoxButtons]::OK,
    $(if ($ok) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning })
)

if (-not $ok) { exit 2 }
exit 0
