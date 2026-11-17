# Import AES VBA modules into Outlook with no keystrokes and no Application.VBE.
#
# Application.VBE does not exist on this Outlook build (DISP_E_UNKNOWNNAME, both
# from add-ins and external COM, AccessVBOM=1), and Norton blocks SendKeys. The
# VBA editor's command bars and project tree are still reachable through MSAA
# (oleacc), so this script invokes the real menu items directly:
#
#   ExecuteMso VisualBasic  -> open the editor
#   File > Remove / Import  -> replace each MSCAN* module
#   View > Code + Paste     -> write the built-in ThisOutlookSession
#   Debug > Compile, File > Save -> persist %APPDATA%\Microsoft\Outlook\VbaProject.OTM
#
# It only ever acts on dialogs raised by the VBA editor itself. Other Outlook
# prompts (certificate warnings, reminders) are left for the user.
#
# Usage:
#   .\Import_VBA_ViaUi.ps1
#   .\Import_VBA_ViaUi.ps1 -Root "D:\Suite" -KeepEditor
#   .\Import_VBA_ViaUi.ps1 -DryRun
#   .\Import_VBA_ViaUi.ps1 -IfStale -Notify   (what the ribbon host runs at startup)
#
# Norton quarantines VbaProject.OTM as MO97:Downloader-YI (BUG-040). That is what
# empties the project, so this script reports it instead of re-importing forever.
#
# Exit codes: 0 ok | 1 failed | 3 imported and saved, but VBA does not compile
#             4 skipped: Norton quarantined the project | 5 saved, but Norton removed the backup

[CmdletBinding()]
param(
    [string]$Root = "",
    [switch]$DryRun,
    [switch]$KeepEditor,
    [switch]$IfStale,
    [switch]$Notify
)

$ErrorActionPreference = "Stop"

# Add-Type below is C# 5 + the desktop Accessibility assembly: Windows PowerShell only.
if ($PSVersionTable.PSEdition -eq "Core") {
    $winPs = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    $relay = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $PSCommandPath)
    if ($Root) { $relay += @("-Root", $Root) }
    if ($DryRun) { $relay += "-DryRun" }
    if ($KeepEditor) { $relay += "-KeepEditor" }
    if ($IfStale) { $relay += "-IfStale" }
    if ($Notify) { $relay += "-Notify" }
    & $winPs @relay
    exit $LASTEXITCODE
}

$vbeUi = @"
using System;
using System.Text;
using System.Threading;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Accessibility;

public static class VbeUi {
  delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc p, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, EnumProc p, IntPtr l);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] static extern bool IsWindow(IntPtr h);
  [DllImport("user32.dll")] static extern int GetDlgCtrlID(IntPtr h);
  [DllImport("user32.dll")] static extern IntPtr GetParent(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, string l);
  [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr h, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IUnknown)] out object o);
  [DllImport("oleacc.dll")] static extern int AccessibleChildren(IAccessible p, int start, int count, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=2)] object[] kids, out int got);

  const uint WM_SETTEXT = 0x000C, WM_COMMAND = 0x0111;
  const uint OBJID_CLIENT = 0xFFFFFFFC;
  const int STATE_UNAVAILABLE = 0x1;
  const int SELFLAG_TAKEFOCUS = 1, SELFLAG_TAKESELECTION = 2;
  public const int IDOK = 1;

  static string Cls(IntPtr h) { var s = new StringBuilder(256); GetClassName(h, s, 256); return s.ToString(); }
  static string Txt(IntPtr h) { var s = new StringBuilder(1024); GetWindowText(h, s, 1024); return s.ToString(); }
  static string Plain(string s) { return (s ?? "").Replace("&", "").Trim(); }

  static List<IntPtr> Tops(int pid) {
    var r = new List<IntPtr>();
    EnumWindows((h, l) => {
      uint p; GetWindowThreadProcessId(h, out p);
      if (p == (uint)pid && IsWindowVisible(h)) r.Add(h);
      return true;
    }, IntPtr.Zero);
    return r;
  }
  static List<IntPtr> Kids(IntPtr parent) {
    var r = new List<IntPtr>();
    EnumChildWindows(parent, (h, l) => { r.Add(h); return true; }, IntPtr.Zero);
    return r;
  }
  static IntPtr Kid(IntPtr parent, string cls, string text) {
    foreach (var h in Kids(parent))
      if (Cls(h) == cls && (text == null || Txt(h) == text)) return h;
    return IntPtr.Zero;
  }
  static IAccessible Acc(IntPtr h) {
    Guid iid = new Guid("618736E0-3C3D-11CF-810C-00AA00389B71");
    object o;
    int hr = AccessibleObjectFromWindow(h, OBJID_CLIENT, ref iid, out o);
    if (hr != 0) throw new Exception("AccessibleObjectFromWindow hr=0x" + hr.ToString("X"));
    return (IAccessible)o;
  }
  static object[] Children(IAccessible a) {
    int n = a.accChildCount;
    if (n == 0) return new object[0];
    var k = new object[n]; int got;
    AccessibleChildren(a, 0, n, k, out got);
    var r = new object[got]; Array.Copy(k, r, got);
    return r;
  }
  static IAccessible Named(IAccessible a, string name, bool prefix) {
    foreach (var c in Children(a)) {
      var x = c as IAccessible;
      if (x == null) continue;
      string n = "";
      try { n = Plain(x.get_accName(0)); } catch { continue; }
      if (prefix ? n.StartsWith(name, StringComparison.OrdinalIgnoreCase)
                 : string.Equals(n, name, StringComparison.OrdinalIgnoreCase)) return x;
    }
    return null;
  }
  static IAccessible MenuItem(IntPtr vbe, string menu, string item) {
    var bar = Kid(vbe, "MsoCommandBar", "Menu Bar");
    if (bar == IntPtr.Zero) throw new Exception("VBA editor menu bar not found");
    var m = Named(Acc(bar), menu, false);
    if (m == null) throw new Exception("Menu not found: " + menu);
    var it = Named(m, item, true);
    if (it == null) throw new Exception("Menu item not found: " + menu + " > " + item);
    return it;
  }

  public static IntPtr FindEditor(int pid) {
    foreach (var h in Tops(pid)) if (Cls(h) == "wndclass_desked_gsk") return h;
    return IntPtr.Zero;
  }

  public static bool MenuEnabled(IntPtr vbe, string menu, string item) {
    return (Convert.ToInt32(MenuItem(vbe, menu, item).get_accState(0)) & STATE_UNAVAILABLE) == 0;
  }

  // A menu command that opens a modal dialog does not return until the dialog
  // closes, so it runs on its own thread with its own MSAA objects.
  public static void Invoke(IntPtr vbe, string menu, string item) {
    if (!MenuEnabled(vbe, menu, item)) throw new Exception("Menu item is disabled: " + menu + " > " + item);
    Exception failure = null;
    var t = new Thread(() => {
      try { MenuItem(vbe, menu, item).accDoDefaultAction(0); }
      catch (Exception ex) { failure = ex; }
    });
    t.IsBackground = true;
    t.SetApartmentState(ApartmentState.STA);
    t.Start();
    t.Join(250);
    if (failure != null) throw new Exception(menu + " > " + item + ": " + failure.Message);
  }

  // Runs a command that may take a while (Compile). Returns the editor message box
  // it raised, or Zero once the command has finished without one.
  public static IntPtr InvokeWait(IntPtr vbe, string menu, string item, int pid, string titlePrefix, int ms) {
    if (!MenuEnabled(vbe, menu, item)) throw new Exception("Menu item is disabled: " + menu + " > " + item);
    Exception failure = null;
    var t = new Thread(() => {
      try { MenuItem(vbe, menu, item).accDoDefaultAction(0); }
      catch (Exception ex) { failure = ex; }
    });
    t.IsBackground = true;
    t.SetApartmentState(ApartmentState.STA);
    t.Start();
    var until = DateTime.UtcNow.AddMilliseconds(ms);
    while (DateTime.UtcNow < until) {
      var h = FindDialog(pid, titlePrefix);
      if (h != IntPtr.Zero) return h;
      if (!t.IsAlive) {
        if (failure != null) throw new Exception(menu + " > " + item + ": " + failure.Message);
        Thread.Sleep(300);
        return FindDialog(pid, titlePrefix);
      }
      Thread.Sleep(50);
    }
    throw new Exception(menu + " > " + item + " did not finish in " + (ms / 1000) + " seconds");
  }

  // --- dialogs ---------------------------------------------------------------
  public static IntPtr FindDialog(int pid, string titlePrefix) {
    foreach (var h in Tops(pid))
      if (Cls(h) == "#32770" && Txt(h).StartsWith(titlePrefix, StringComparison.OrdinalIgnoreCase)) return h;
    return IntPtr.Zero;
  }
  public static IntPtr WaitDialog(int pid, string titlePrefix, int ms) {
    var until = DateTime.UtcNow.AddMilliseconds(ms);
    do {
      var h = FindDialog(pid, titlePrefix);
      if (h != IntPtr.Zero) return h;
      Thread.Sleep(50);
    } while (DateTime.UtcNow < until);
    return IntPtr.Zero;
  }
  public static bool WaitClosed(IntPtr dlg, int ms) {
    var until = DateTime.UtcNow.AddMilliseconds(ms);
    while (IsWindow(dlg) && IsWindowVisible(dlg)) {
      if (DateTime.UtcNow > until) return false;
      Thread.Sleep(50);
    }
    return true;
  }
  public static string DialogText(IntPtr dlg) {
    var b = new StringBuilder();
    foreach (var h in Kids(dlg)) {
      if (Cls(h) == "Static" && IsWindowVisible(h)) {
        string t = Txt(h).Trim();
        if (t.Length > 0) b.Append(t).Append(' ');
      }
    }
    return b.ToString().Trim();
  }
  // WM_COMMAND rather than BM_CLICK: it does not need the dialog to be active.
  public static void ClickId(IntPtr dlg, int id) {
    foreach (var h in Kids(dlg)) {
      if (Cls(h) == "Button" && GetDlgCtrlID(h) == id) {
        PostMessage(dlg, WM_COMMAND, (IntPtr)id, h);
        return;
      }
    }
    PostMessage(dlg, WM_COMMAND, (IntPtr)id, IntPtr.Zero);
  }
  public static bool ClickText(IntPtr dlg, string caption) {
    foreach (var h in Kids(dlg)) {
      if (Cls(h) == "Button" && string.Equals(Plain(Txt(h)), caption, StringComparison.OrdinalIgnoreCase)) {
        PostMessage(dlg, WM_COMMAND, (IntPtr)GetDlgCtrlID(h), h);
        return true;
      }
    }
    return false;
  }
  public static bool SetFileName(IntPtr dlg, string path) {
    foreach (var h in Kids(dlg)) {
      if (Cls(h) == "Edit" && IsWindowVisible(h) && Cls(GetParent(h)) == "ComboBox") {
        SendMessage(h, WM_SETTEXT, IntPtr.Zero, path);
        return true;
      }
    }
    return false;
  }

  // --- project tree ----------------------------------------------------------
  static IntPtr Tree(IntPtr vbe) {
    var t = Kid(vbe, "SysTreeView32", null);
    if (t == IntPtr.Zero) throw new Exception("Project Explorer not found (View > Project Explorer)");
    return t;
  }
  static string NodeName(IAccessible t, object id) {
    string n = "";
    try { n = t.get_accName(id); } catch { }
    n = (n ?? "").Trim();
    // Folder view shows "Name"; object view shows "CodeName (Name)".
    int p = n.IndexOf(" (", StringComparison.Ordinal);
    return p > 0 ? n.Substring(0, p) : n;
  }
  public static void ExpandTree(IntPtr vbe) {
    for (int pass = 0; pass < 8; pass++) {
      var t = Acc(Tree(vbe));
      bool any = false;
      foreach (var id in Children(t)) {
        string act = "";
        try { act = t.get_accDefaultAction(id); } catch { }
        if (string.Equals(act, "Expand", StringComparison.OrdinalIgnoreCase)) {
          t.accDoDefaultAction(id);
          any = true;
          Thread.Sleep(120);
          break;
        }
      }
      if (!any) return;
    }
  }
  public static string[] TreeNames(IntPtr vbe) {
    var t = Acc(Tree(vbe));
    var r = new List<string>();
    foreach (var id in Children(t)) r.Add(NodeName(t, id));
    return r.ToArray();
  }
  public static void SelectNode(IntPtr vbe, string name) {
    var t = Acc(Tree(vbe));
    foreach (var id in Children(t)) {
      if (string.Equals(NodeName(t, id), name, StringComparison.OrdinalIgnoreCase)) {
        t.accSelect(SELFLAG_TAKEFOCUS | SELFLAG_TAKESELECTION, id);
        Thread.Sleep(150);
        return;
      }
    }
    throw new Exception("Not in Project Explorer: " + name);
  }
}
"@
Add-Type -TypeDefinition $vbeUi -ReferencedAssemblies Accessibility -ErrorAction Stop

$script:LogPath = Join-Path $env:LOCALAPPDATA "GeoFooter\Logs\VbaImport.log"
function Log {
    param([string]$Message)
    Write-Host $Message
    try {
        New-Item -ItemType Directory -Force -Path (Split-Path $script:LogPath) | Out-Null
        Add-Content -LiteralPath $script:LogPath -Value ("{0:yyyy-MM-dd HH:mm:ss} | {1}" -f (Get-Date), $Message) -Encoding UTF8
    } catch {}
}

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

# IMPORT.txt is the single source of truth for the module list and its order.
function Get-ModuleFiles {
    param([string]$VbaDir)
    $files = @()
    foreach ($line in Get-Content (Join-Path $VbaDir "IMPORT.txt")) {
        # A file listed twice would import twice, as MSCANPaths and MSCANPaths1.
        if ($line -match '^\s*VBA\\(MSCAN[^\s\\]+\.(?:bas|cls))\s*$' -and $files -notcontains $Matches[1]) { $files += $Matches[1] }
    }
    if (-not $files) { throw "No VBA\MSCAN* modules listed in IMPORT.txt" }
    return $files
}

# The editor mangles LF-only files on import, so always hand it CRLF. Bytes are
# otherwise untouched: VBA source is ANSI, not UTF-8.
function New-CrlfCopy {
    param([string]$Source, [string]$TempDir)
    $bytes = [IO.File]::ReadAllBytes($Source)
    $out = New-Object System.Collections.Generic.List[byte] ($bytes.Length + 2048)
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        $b = $bytes[$i]
        if ($b -eq 10 -and ($i -eq 0 -or $bytes[$i - 1] -ne 13)) { $out.Add(13) }
        $out.Add($b)
    }
    $dest = Join-Path $TempDir (Split-Path $Source -Leaf)
    [IO.File]::WriteAllBytes($dest, $out.ToArray())
    return $dest
}

function Get-OutlookPath {
    $key = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"
    $p = (Get-ItemProperty $key -ErrorAction SilentlyContinue).'(default)'
    if ($p) { return $p }
    throw "Could not find OUTLOOK.EXE (App Paths registry key missing)."
}

# A headless COM instance (-Embedding) has no explorer, so it cannot host the editor.
function Get-InteractiveOutlook {
    param([int]$TimeoutSec = 120)
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $started = $false
    while ((Get-Date) -lt $deadline) {
        $ol = Get-Process OUTLOOK -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } | Select-Object -First 1
        if ($ol) { return $ol }
        if (-not $started) {
            Log "Starting Outlook..."
            Start-Process (Get-OutlookPath)
            $started = $true
        }
        Start-Sleep -Seconds 1
    }
    throw "Outlook did not open a window within $TimeoutSec seconds."
}

function Open-Editor {
    param([int]$OutlookPid)
    $h = [VbeUi]::FindEditor($OutlookPid)
    if ($h -ne [IntPtr]::Zero) { return $h }
    $deadline = (Get-Date).AddSeconds(90)
    $lastErr = ""
    while ((Get-Date) -lt $deadline) {
        try {
            $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
            $exp = $app.ActiveExplorer()
            if ($exp) { $exp.CommandBars.ExecuteMso("VisualBasic") }
        } catch { $lastErr = $_.Exception.Message }
        Start-Sleep -Seconds 2
        $h = [VbeUi]::FindEditor($OutlookPid)
        if ($h -ne [IntPtr]::Zero) { return $h }
    }
    throw "VBA editor did not open: $lastErr"
}

# Report and clear any message box the editor raised; returns its text or "".
function Read-EditorMessage {
    param([int]$OutlookPid, [int]$WaitMs = 0)
    $d = [VbeUi]::WaitDialog($OutlookPid, "Microsoft Visual Basic", $WaitMs)
    if ($d -eq [IntPtr]::Zero) { return "" }
    $text = [VbeUi]::DialogText($d)
    [VbeUi]::ClickId($d, [VbeUi]::IDOK)
    [void][VbeUi]::WaitClosed($d, 3000)
    return $(if ($text) { $text } else { "(editor message with no text)" })
}

function Remove-Module {
    param([IntPtr]$Vbe, [int]$OutlookPid, [string]$Name)
    [VbeUi]::SelectNode($Vbe, $Name)
    [VbeUi]::Invoke($Vbe, "File", "Remove")
    $prompt = [VbeUi]::WaitDialog($OutlookPid, "Microsoft Visual Basic", 5000)
    if ($prompt -eq [IntPtr]::Zero) { throw "No confirmation appeared when removing $Name" }
    $text = [VbeUi]::DialogText($prompt)
    if ($text -notmatch 'export') { throw "Unexpected prompt removing ${Name}: $text" }
    if (-not [VbeUi]::ClickText($prompt, "No")) { throw "No 'No' button removing $Name" }
    if (-not [VbeUi]::WaitClosed($prompt, 5000)) { throw "Remove prompt did not close for $Name" }
    Start-Sleep -Milliseconds 150
}

function Import-Module-File {
    param([IntPtr]$Vbe, [int]$OutlookPid, [string]$Path)
    # The command can be swallowed if Outlook is busy for a moment (a mail
    # arriving, the user clicking). Ask again rather than fail the whole import
    # with the modules already removed.
    $dlg = [IntPtr]::Zero
    foreach ($attempt in 1..4) {
        try { [VbeUi]::Invoke($Vbe, "File", "Import File") } catch { Start-Sleep -Seconds 1; continue }
        $dlg = [VbeUi]::WaitDialog($OutlookPid, "Import File", 8000)
        if ($dlg -ne [IntPtr]::Zero) { break }
    }
    if ($dlg -eq [IntPtr]::Zero) { throw "Import File dialog did not open after 4 attempts" }
    if (-not [VbeUi]::SetFileName($dlg, $Path)) {
        [VbeUi]::ClickId($dlg, 2)
        throw "File name box not found in Import File dialog"
    }
    [VbeUi]::ClickId($dlg, [VbeUi]::IDOK)
    if (-not [VbeUi]::WaitClosed($dlg, 20000)) {
        $why = Read-EditorMessage -OutlookPid $OutlookPid
        [VbeUi]::ClickId($dlg, 2)
        throw "Import dialog stayed open for $Path $why"
    }
    $msg = Read-EditorMessage -OutlookPid $OutlookPid -WaitMs 200
    if ($msg) { throw "Editor refused ${Path}: $msg" }
}

function Get-ThisOutlookSessionCode {
    param([string]$Path)
    $lines = [IO.File]::ReadAllLines($Path, [Text.Encoding]::Default)
    $i = 0
    while ($i -lt $lines.Length -and $lines[$i] -match '^(VERSION |BEGIN\b|\s+MultiUse|END\b|Attribute )') { $i++ }
    return (($lines[$i..($lines.Length - 1)]) -join "`r`n") + "`r`n"
}

# Text, images and copied files on the clipboard, so they can be put back.
# Only these three, through their typed calls. Copying every format generically
# crashed PowerShell (access violation in clr.dll, 2026-11-17 11:15) when the
# clipboard held something handle-based, and that left the send hold in place.
function Save-ClipboardData {
    $saved = @{}
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $clip = [Windows.Forms.Clipboard]
        $text = @{}
        foreach ($format in "UnicodeText", "Rtf", "Html") {
            $kind = [Windows.Forms.TextDataFormat]$format
            if ($clip::ContainsText($kind)) { $text[$format] = $clip::GetText($kind) }
        }
        if ($text.Count) { $saved.Text = $text }
        if ($clip::ContainsFileDropList()) { $saved.Files = $clip::GetFileDropList() }
        if ($clip::ContainsImage()) { $saved.Image = $clip::GetImage() }
    } catch {}
    return $saved
}

function Restore-ClipboardData {
    param($Data)
    try {
        $clip = [Windows.Forms.Clipboard]
        if (-not $Data -or $Data.Count -eq 0) { $clip::Clear(); return }
        $object = New-Object Windows.Forms.DataObject
        if ($Data.Text) {
            foreach ($format in $Data.Text.Keys) {
                $object.SetText($Data.Text[$format], [Windows.Forms.TextDataFormat]$format)
            }
        }
        if ($Data.Files) { $object.SetFileDropList($Data.Files) }
        if ($Data.Image) { $object.SetImage($Data.Image) }
        $clip::SetDataObject($object, $true)
    } catch {}
}

# While modules are being replaced, the real ThisOutlookSession would call code
# that is not there. This stand-in calls nothing and holds any send, so mail cannot
# leave unclassified. -Failed adds the explanation once no automation is running:
# a MsgBox during the import would block the editor and fail it.
function Get-GuardCode {
    param([switch]$Failed)
    $say = ""
    if ($Failed) {
        $say = "    MsgBox ""The AES update did not finish, so this send was cancelled."" & vbCrLf & vbCrLf & _`r`n" +
               "           ""Close and reopen Outlook, then send again."", vbExclamation, ""AES""`r`n"
    }
    return "Option Explicit`r`n" +
        "' AES modules are being replaced (scripts\Import_VBA_ViaUi.ps1). No module calls here.`r`n" +
        "Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)`r`n" +
        "    Cancel = True`r`n" + $say +
        "End Sub`r`n"
}

# Edit commands are greyed out until the code pane has focus, which can take
# seconds after a large import. Fixed sleeps here failed twice in normal use.
function Invoke-EditCommand {
    param([IntPtr]$Vbe, [string]$Item, [int]$TimeoutSec = 30)
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $nextOpen = Get-Date
    $lastError = "stayed disabled"
    while ((Get-Date) -le $deadline) {
        if ([VbeUi]::MenuEnabled($Vbe, "Edit", $Item)) {
            # The pane can lose focus between the check and the command (the user
            # clicks, a mail arrives). That failed an import on 2026-11-17 and left
            # sends held, so a failed command goes round again instead of throwing.
            try { [VbeUi]::Invoke($Vbe, "Edit", $Item); return }
            catch { $lastError = $_.Exception.Message }
        }
        if ((Get-Date) -ge $nextOpen) {
            try { [VbeUi]::SelectNode($Vbe, "ThisOutlookSession"); [VbeUi]::Invoke($Vbe, "View", "Code") } catch {}
            $nextOpen = (Get-Date).AddSeconds(3)
        }
        Start-Sleep -Milliseconds 200
    }
    throw "Edit > $Item did not run in $TimeoutSec seconds: $lastError"
}

function Set-ThisOutlookSessionCode {
    param([IntPtr]$Vbe, [string]$Code)
    $saved = Save-ClipboardData
    try {
        Invoke-EditCommand -Vbe $Vbe -Item "Select All"
        Start-Sleep -Milliseconds 200
        Set-Clipboard -Value $Code
        Start-Sleep -Milliseconds 150
        Invoke-EditCommand -Vbe $Vbe -Item "Paste"
        Start-Sleep -Milliseconds 500
        # Read the pane back: a silently failed paste must not count as success.
        Invoke-EditCommand -Vbe $Vbe -Item "Select All"
        Start-Sleep -Milliseconds 150
        Set-Clipboard -Value " "
        Invoke-EditCommand -Vbe $Vbe -Item "Copy"
        Start-Sleep -Milliseconds 300
        $back = Get-Clipboard -Raw
        if (($back -replace '\s+', '') -ne ($Code -replace '\s+', '')) {
            throw "ThisOutlookSession code did not match after paste"
        }
    } finally {
        Restore-ClipboardData $saved
    }
}

# Content hash of VbaProject.OTM, read in place while Outlook has it open. Proves a
# Save reached disk. Not the timestamp: Outlook's saves leave it unchanged. Not a
# copy either: Norton blocks copies of this file outside the excluded folders.
function Get-OtmStamp {
    param([string]$Otm)
    if (-not (Test-Path -LiteralPath $Otm)) { return "missing" }
    for ($try = 0; $try -lt 5; $try++) {
        $stream = $null
        try {
            $stream = [IO.File]::Open($Otm, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
            return [BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash($stream))
        } catch { Start-Sleep -Milliseconds 300 }
        finally { if ($stream) { $stream.Close() } }
    }
    throw "Could not read $Otm to check the save."
}

# Fingerprint of what was imported, so the ribbon host can tell when VBA\ is newer.
function Get-VbaFingerprint {
    param([string]$VbaDir, [string[]]$Files)
    $sha = [Security.Cryptography.SHA256]::Create()
    $sb = New-Object Text.StringBuilder
    foreach ($f in ($Files | Sort-Object)) {
        $hash = [BitConverter]::ToString($sha.ComputeHash([IO.File]::ReadAllBytes((Join-Path $VbaDir $f))))
        [void]$sb.Append($f).Append('=').Append($hash).Append(';')
    }
    return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($sb.ToString()))) -replace '-', ''
}

function Test-VbaCurrent {
    param([string]$VbaDir, [string[]]$Files)
    $otm = Join-Path $env:APPDATA "Microsoft\Outlook\VbaProject.OTM"
    $state = Join-Path $env:LOCALAPPDATA "GeoFooter\vba_sync_state.json"
    if (-not (Test-Path $otm) -or -not (Test-Path $state)) { return $false }
    try { $saved = Get-Content -LiteralPath $state -Raw | ConvertFrom-Json } catch { return $false }
    return ($saved.fingerprint -eq (Get-VbaFingerprint -VbaDir $VbaDir -Files (@($Files) + "ThisOutlookSession.cls")))
}

# Newest time Norton quarantined a VbaProject*.otm, from its readable index; else $null.
function Get-NortonQuarantineTime {
    $index = Join-Path $env:ProgramData "Norton\Antivirus\chest\index.xml"
    if (-not (Test-Path $index)) { return $null }
    try { $xml = [IO.File]::ReadAllText($index) } catch { return $null }
    $newest = $null
    foreach ($m in [regex]::Matches($xml, '(?s)<ChestEntry>(.*?)</ChestEntry>')) {
        $e = $m.Groups[1].Value
        if ($e -notmatch '<OrigFileName>VbaProject[^<]*\.otm</OrigFileName>') { continue }
        if ($e -match '<FileTime>(\d+)</FileTime>') {
            $t = [DateTimeOffset]::FromUnixTimeSeconds([long]$Matches[1]).LocalDateTime
            if (-not $newest -or $t -gt $newest) { $newest = $t }
        }
    }
    return $newest
}

function Get-NortonAdvice {
    return "Norton quarantined the AES VBA project as MO97:Downloader-YI, which is why AES modules keep disappearing.`n`n" +
        "In Norton, add exclusions for:`n" +
        "  $env:APPDATA\Microsoft\Outlook\VbaProject.OTM`n" +
        "  $env:LOCALAPPDATA\GeoFooter\vba_otm_backups`n`n" +
        "Then click Repair VBA on the AES ribbon."
}

# True when the project is gone and Norton took it after our last import. Importing
# again would only be undone, so startup stops here until the exclusion is in place.
function Test-QuarantineBlocksImport {
    if (Test-Path (Join-Path $env:APPDATA "Microsoft\Outlook\VbaProject.OTM")) { return $false }
    $quarantined = Get-NortonQuarantineTime
    if (-not $quarantined) { return $false }
    $state = Join-Path $env:LOCALAPPDATA "GeoFooter\vba_sync_state.json"
    if (-not (Test-Path $state)) { return $true }
    try { $imported = [datetime](Get-Content -LiteralPath $state -Raw | ConvertFrom-Json).imported_at } catch { return $true }
    return ($quarantined -ge $imported)
}

function Show-QuarantineOncePerDay {
    if (-not $Notify) { return }
    $flag = Join-Path $env:LOCALAPPDATA "GeoFooter\vba_quarantine_notice_day.txt"
    $today = (Get-Date).ToString("yyyy-MM-dd")
    if ((Test-Path $flag) -and ((Get-Content -LiteralPath $flag -TotalCount 1) -eq $today)) { return }
    Set-Content -LiteralPath $flag -Value $today -Encoding ASCII
    Show-Result -Text (Get-NortonAdvice) -Icon 48
}

# Queued, then shown once the import lock is released: a popup left open must
# not make the next run report "already running".
$script:PendingPopup = $null
function Show-Result {
    param([string]$Text, [int]$Icon)
    if ($Notify) { $script:PendingPopup = @{ Text = $Text; Icon = $Icon } }
}

function Show-PendingPopup {
    $p = $script:PendingPopup
    $script:PendingPopup = $null
    if (-not $p) { return }
    try { [void](New-Object -ComObject WScript.Shell).Popup($p.Text, 0, "Aliniant AES - VBA update", $p.Icon) } catch {}
}

# A macro running (AES startup, say) greys out the editor's menus until it ends.
function Wait-EditorIdle {
    param([IntPtr]$Vbe, [int]$TimeoutSec = 120)
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        if ([VbeUi]::MenuEnabled($Vbe, "File", "Import File")) { return }
        Start-Sleep -Milliseconds 500
    }
    throw "VBA editor stayed busy for $TimeoutSec seconds (a macro is still running)."
}

# --- main ---
$mutex = New-Object Threading.Mutex($false, "Local\GeoFooter_VbaImport")
$haveLock = $false
try { $haveLock = $mutex.WaitOne(0) } catch [Threading.AbandonedMutexException] { $haveLock = $true }
if (-not $haveLock) {
    Log "Another VBA import is already running."
    # Someone pressed Repair VBA or ran the .bat: tell them. The startup check stays quiet.
    if (-not $IfStale) {
        Show-Result -Text "An AES VBA update is already running. You will get a message when it finishes." -Icon 64
        Show-PendingPopup
    }
    exit 1
}

$tempDir = $null
$exitCode = 1
$guardOn = $false
$vbe = [IntPtr]::Zero
try {
    $rootPath = Resolve-InstallRoot -Hint $Root
    $vbaDir = Join-Path $rootPath "VBA"
    $files = @(Get-ModuleFiles -VbaDir $vbaDir)
    $missing = @($files | Where-Object { -not (Test-Path (Join-Path $vbaDir $_)) })
    if ($missing) { throw "Missing module file(s): $($missing -join ', ')" }
    $tos = Join-Path $vbaDir "ThisOutlookSession.cls"
    if (-not (Test-Path $tos)) { throw "Missing $tos" }

    if ($DryRun) {
        Write-Host "DryRun: would import $($files.Count) modules from $vbaDir"
        $files | ForEach-Object { Write-Host "  $_" }
        exit 0
    }

    if ($IfStale -and (Test-VbaCurrent -VbaDir $vbaDir -Files $files)) {
        Log "VBA project already matches disk; nothing to do."
        exit 0
    }
    if ($IfStale -and (Test-QuarantineBlocksImport)) {
        Log "Skipped: Norton quarantined VbaProject.OTM at $(Get-NortonQuarantineTime). Add the exclusions, then Repair VBA."
        Show-QuarantineOncePerDay
        $exitCode = 4
        exit 4
    }

    Log "=== VBA import start: root=$rootPath modules=$($files.Count)"
    $statePath = Join-Path $env:LOCALAPPDATA "GeoFooter\vba_sync_state.json"
    $fingerprint = Get-VbaFingerprint -VbaDir $vbaDir -Files (@($files) + "ThisOutlookSession.cls")
    $realCode = Get-ThisOutlookSessionCode -Path $tos
    # If this run fails and the half-built project gets saved, the next start must
    # not see the old fingerprint and call it current.
    if (Test-Path $statePath) {
        try {
            $old = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
            @{ fingerprint = "import-in-progress"; version = $old.version; imported_at = $old.imported_at; compiled = $false } |
                ConvertTo-Json | Set-Content -LiteralPath $statePath -Encoding ASCII
        } catch { Remove-Item -LiteralPath $statePath -Force -ErrorAction SilentlyContinue }
    }
    $ol = Get-InteractiveOutlook
    $olPid = [int]$ol.Id
    $vbe = Open-Editor -OutlookPid $olPid
    Log "Editor open (Outlook pid $olPid)"
    Start-Sleep -Milliseconds 500
    Wait-EditorIdle -Vbe $vbe

    $tempDir = Join-Path $env:TEMP ("aes_vba_" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path $tempDir | Out-Null

    [VbeUi]::ExpandTree($vbe)
    Log "Holding sends while modules are replaced"
    Set-ThisOutlookSessionCode -Vbe $vbe -Code (Get-GuardCode)
    $guardOn = $true
    $stale = @("UserForm1", "frmClassification", "MSCANSecurityRatingForm", "ThisOutlookSession1", "ThisOutlookSession2")
    $wanted = @($files | ForEach-Object { [IO.Path]::GetFileNameWithoutExtension($_) })
    $existing = @([VbeUi]::TreeNames($vbe) | Where-Object { $_ -like "MSCAN*" -or $stale -contains $_ })
    foreach ($name in $existing) {
        Log "Removing $name"
        Remove-Module -Vbe $vbe -OutlookPid $olPid -Name $name
    }

    $n = 0
    foreach ($f in $files) {
        $n++
        Log ("[{0}/{1}] {2}" -f $n, $files.Count, $f)
        Import-Module-File -Vbe $vbe -OutlookPid $olPid -Path (New-CrlfCopy -Source (Join-Path $vbaDir $f) -TempDir $tempDir)
    }

    [VbeUi]::ExpandTree($vbe)
    $present = @([VbeUi]::TreeNames($vbe))
    $absent = @($wanted | Where-Object { $present -notcontains $_ })
    if ($absent) { throw "Not in the project after import: $($absent -join ', ')" }
    Log "All $($wanted.Count) modules are in the project"

    Log "Writing built-in ThisOutlookSession"
    $written = $false
    foreach ($attempt in 1..3) {
        try { Set-ThisOutlookSessionCode -Vbe $vbe -Code $realCode; $written = $true; break }
        catch {
            Log "ThisOutlookSession write attempt $attempt failed: $($_.Exception.Message)"
            if ($attempt -eq 3) { throw }
            Start-Sleep -Seconds 2
        }
    }
    $guardOn = $false

    Log "Compiling"
    $compileError = ""
    if ([VbeUi]::MenuEnabled($vbe, "Debug", "Compile")) {
        # Waits for the compile itself, however long: a fixed wait reported a slow
        # error as "Compile OK".
        if ([VbeUi]::InvokeWait($vbe, "Debug", "Compile", $olPid, "Microsoft Visual Basic", 180000) -ne [IntPtr]::Zero) {
            $compileError = Read-EditorMessage -OutlookPid $olPid
        }
        if ($compileError) { Log "COMPILE ERROR: $compileError" } else { Log "Compile OK" }
    }

    # Save even after a compile error: the source is still worth keeping on disk.
    $otm = Join-Path $env:APPDATA "Microsoft\Outlook\VbaProject.OTM"
    $stampBefore = Get-OtmStamp -Otm $otm
    $stampAfter = $stampBefore
    foreach ($attempt in 1, 2) {
        Log "Saving$(if ($attempt -gt 1) { ' (second attempt)' })"
        [VbeUi]::Invoke($vbe, "File", "Save")
        $msg = Read-EditorMessage -OutlookPid $olPid -WaitMs 1500
        if ($msg) { throw "Save failed: $msg" }
        $deadline = (Get-Date).AddSeconds(20)
        do {
            Start-Sleep -Milliseconds 500
            $stampAfter = Get-OtmStamp -Otm $otm
        } while ($stampAfter -eq $stampBefore -and (Get-Date) -lt $deadline)
        if ($stampAfter -ne $stampBefore) { break }
    }
    if ($stampAfter -eq "missing") { throw "Save ran but $otm was not written." }
    if ($stampAfter -eq $stampBefore) { throw "Save ran twice but $otm did not change on disk." }
    $len = (Get-Item $otm).Length
    Log "VbaProject.OTM written: $len bytes"

    $backupDir = Join-Path $env:LOCALAPPDATA "GeoFooter\vba_otm_backups"
    New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
    $backup = Join-Path $backupDir ("VbaProject_{0:yyyyMMdd_HHmmss}_{1}.otm" -f (Get-Date), $len)
    try { Copy-Item -LiteralPath $otm -Destination $backup -Force; Log "Backup: $backup" }
    catch { Log "Backup skipped (OTM in use; next shutdown backup will cover it): $($_.Exception.Message)" }

    # Norton took the last backup 7 seconds after it was written, so this is an early
    # test of whether the exclusions are in place. The OTM itself goes when Outlook exits.
    $backupRemoved = $false
    if (Test-Path -LiteralPath $backup) {
        Start-Sleep -Seconds 15
        $backupRemoved = -not (Test-Path -LiteralPath $backup)
    }

    @{
        fingerprint = $fingerprint
        version     = (Get-Content (Join-Path $rootPath "VERSION") -TotalCount 1).Trim()
        imported_at = (Get-Date).ToString("s")
        compiled    = (-not $compileError)
    } | ConvertTo-Json | Set-Content -LiteralPath $statePath -Encoding ASCII

    if (-not $KeepEditor) {
        try { [VbeUi]::Invoke($vbe, "File", "Close and Return") } catch { Log "Editor left open: $($_.Exception.Message)" }
    }

    if ($backupRemoved) {
        Log "=== VBA import saved, but Norton removed the backup. The project will be quarantined when Outlook exits."
        Show-Result -Text (Get-NortonAdvice) -Icon 48
        $exitCode = 5
    } elseif ($compileError) {
        Log "=== VBA import finished, but the project does not compile: $compileError"
        Show-Result -Text "AES modules were imported and saved, but they do not compile:`n`n$compileError`n`nLog: $script:LogPath" -Icon 48
        $exitCode = 3
    } else {
        Log "=== VBA import OK. Restart Outlook to load AES."
        Show-Result -Text "AES modules are updated and saved.`n`nClose and reopen Outlook to load them." -Icon 64
        $exitCode = 0
    }
} catch {
    Log "VBA import FAILED: $($_.Exception.Message)"
    if ($guardOn) {
        # Modules are incomplete, so sends stay held. Say why from now on.
        try { Set-ThisOutlookSessionCode -Vbe $vbe -Code (Get-GuardCode -Failed); Log "Sends stay held until Outlook is restarted" }
        catch { Log "Could not update the send guard: $($_.Exception.Message)" }
    }
    Show-Result -Text "AES VBA update failed. Nothing was saved, so the last saved project is still on disk.`n`nClose and reopen Outlook to go back to it. If Outlook asks to save the VBA project, choose No.`n`n$($_.Exception.Message)`n`nLog: $script:LogPath" -Icon 16
    $exitCode = 1
} finally {
    if ($tempDir -and (Test-Path $tempDir)) { Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
    $mutex.ReleaseMutex()
    Show-PendingPopup
}
exit $exitCode
