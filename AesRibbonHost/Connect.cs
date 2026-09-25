using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Extensibility;
using Microsoft.Office.Core;
using stdole;

namespace Aliniant.AesRibbonHost
{
    /// <summary>
    /// Outlook COM add-in: hosts AES Ribbon XML on Home (Mail) with large icons
    /// and live ON/OFF. Actions Execute the existing VBA AES CommandBar buttons.
    /// </summary>
    // AutoDual: Office resolves ribbon callbacks (OnRibbonLoad, GetServiceLabel, ...)
    // by IDispatch name lookup on the class — they must be on the class interface.
    [ComVisible(true)]
    [Guid("A1E50001-AE51-4B0B-9C11-AE5B1BB00001")]
    [ProgId("Aliniant.AesRibbonHost")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        public const string ProgId = "Aliniant.AesRibbonHost";

        private object _application;
        private IRibbonUI _ribbon;
        private FileSystemWatcher _stateWatcher;
        private Timer _statePollTimer;
        private string _lastStateSnapshot;
        private long _lastInvalidateTicks;
        private readonly object _stateSync = new object();

        static Connect()
        {
            HostLog.Write("CLR loaded AesRibbonHost assembly");
        }

        public Connect()
        {
            HostLog.Write("Connect ctor");
        }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            HostLog.Write("OnConnection mode=" + connectMode);
            try
            {
                _application = application;
                // Expose this instance as COMAddIn.Object for VBA InvalidateAesRibbon.
                // Must not throw — a fault here shows as "runtime error" and unload.
                try
                {
                    if (addInInst != null)
                    {
                        dynamic addin = addInInst;
                        addin.Object = this;
                        HostLog.Write("COMAddIn.Object assigned");
                    }
                }
                catch (Exception ex)
                {
                    HostLog.Write("COMAddIn.Object assign skipped: " + ex.Message);
                }

                StartStateWatch();
            }
            catch (Exception ex)
            {
                HostLog.Write("OnConnection FATAL: " + ex);
                throw;
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            HostLog.Write("OnDisconnection mode=" + removeMode);
            StopStateWatch();
            _ribbon = null;
            _application = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom)
        {
            HostLog.Write("OnStartupComplete");
            StartStateWatch();
            InvalidateAesRibbon();
        }

        public void OnBeginShutdown(ref Array custom)
        {
            HostLog.Write("OnBeginShutdown");
            StopStateWatch();
            _ribbon = null;
        }

        public string GetCustomUI(string ribbonID)
        {
            HostLog.Write("GetCustomUI ribbonID=" + ribbonID);
            // Map Outlook windows → built-in tab that hosts the AES group.
            string tabIdMso = null;
            if (string.Equals(ribbonID, "Microsoft.Outlook.Explorer", StringComparison.OrdinalIgnoreCase))
                tabIdMso = "TabMail";
            else if (string.Equals(ribbonID, "Microsoft.Outlook.Mail.Compose", StringComparison.OrdinalIgnoreCase))
                tabIdMso = "TabNewMailMessage";
            else if (string.Equals(ribbonID, "Microsoft.Outlook.Mail.Read", StringComparison.OrdinalIgnoreCase))
                tabIdMso = "TabReadMessage";

            if (tabIdMso == null)
            {
                HostLog.Write("GetCustomUI skipped (no AES tab for this ribbon)");
                return string.Empty;
            }

            var xml = LoadRibbonXml(tabIdMso);
            HostLog.Write("GetCustomUI tab=" + tabIdMso + " length=" + (xml == null ? 0 : xml.Length));
            return xml;
        }

        public void OnRibbonLoad(IRibbonUI ribbonUi)
        {
            HostLog.Write("OnRibbonLoad");
            _ribbon = ribbonUi;
        }

        /// <summary>Called from VBA after ToggleService so Home icons refresh.</summary>
        public void InvalidateAesRibbon()
        {
            try
            {
                // Only the service control changes state frequently; full Invalidate
                // reloads every AES icon from disk and freezes Outlook briefly.
                if (_ribbon == null) return;
                _ribbon.InvalidateControl("AESActivateDeactivate");
                Interlocked.Exchange(ref _lastInvalidateTicks, Environment.TickCount);
            }
            catch
            {
                // Ribbon may not be ready yet.
            }
        }

        public string GetServiceLabel(IRibbonControl control)
        {
            if (IsServiceBusy()) return "AES PROC";
            if (!IsServiceOn()) return "AES OFF";
            return WatcherCount() == 0 ? "AES ON (0 inboxes!)" : "AES ON";
        }

        public string GetServiceScreentip(IRibbonControl control)
        {
            if (IsServiceBusy())
            {
                var n = PendingJobCount();
                return n > 0
                    ? ("AES PROC — processing " + n + " scan(s)")
                    : "AES PROC — processing scan(s)";
            }
            if (!IsServiceOn()) return "AES scanning is OFF";
            return WatcherCount() == 0
                ? "AES is ON but NO inboxes are being watched"
                : "AES scanning is ON";
        }

        public string GetServiceSupertip(IRibbonControl control)
        {
            if (IsServiceBusy())
                return "AES is analyzing header(s) / building footer or deep-scan report. Label shows AES PROC; returns to green ON / red OFF when finished.";
            if (!IsServiceOn())
                return "Automatic inbox scanning is stopped. Click to turn ON.";
            return WatcherCount() == 0
                ? "The AES service is ON, but all email accounts are disabled in AES Settings, so nothing is actually being scanned. Enable at least one account in AES Settings."
                : "Automatic inbox scanning is active. Click to turn OFF.";
        }

        public object GetServiceImage(IRibbonControl control)
        {
            if (IsServiceBusy())
                return LoadIconPicture("aes_service_busy.bmp") ?? LoadIconPicture("aes_service_on.bmp");
            if (IsServiceOn() && WatcherCount() == 0)
            {
                // ON but nothing watched: amber warning icon instead of green
                return LoadIconPicture("aes_service_busy.bmp") ?? LoadIconPicture("aes_service_on.bmp");
            }
            var file = IsServiceOn() ? "aes_service_on.bmp" : "aes_service_off.bmp";
            return LoadIconPicture(file);
        }

        public object GetShortScanImage(IRibbonControl control)
        {
            return LoadIconPicture("aes_short_scan.bmp") ?? LoadIconPicture("aes_short_scan_drawn.bmp");
        }

        public object GetFullScanImage(IRibbonControl control)
        {
            return LoadIconPicture("aes_full_scan.bmp");
        }

        public object GetDeepScanImage(IRibbonControl control)
        {
            return LoadIconPicture("aes_deep_scan.bmp") ?? LoadIconPicture("aes_full_scan.bmp");
        }

        public object GetGuriImage(IRibbonControl control)
        {
            return LoadIconPicture("aes_guri.bmp")
                ?? LoadIconPicture("aes_guri.png")
                ?? LoadIconPicture("aes_diagnostics.png");
        }

        public object GetAuraImage(IRibbonControl control)
        {
            return LoadIconPicture("aes_aura.bmp")
                ?? LoadIconPicture("aes_aura.png")
                ?? LoadIconPicture("aes_guri.bmp")
                ?? LoadIconPicture("aes_diagnostics.png");
        }

        public object GetDiagnosticsImage(IRibbonControl control)
        {
            return LoadIconPicture("aes_diagnostics.png");
        }

        public object GetViewLogsImage(IRibbonControl control)
        {
            return LoadIconPicture("aes_view_logs.png");
        }

        public void OnToggleService(IRibbonControl control)
        {
            ExecuteVbaButton("AES_SERVICE_TOGGLE", "ToggleService");
            // VBA updates state file; refresh after a moment via Invalidate
            InvalidateAesRibbon();
        }

        public void OnShortScan(IRibbonControl control)
        {
            ExecuteVbaButton("AES_SHORTSCAN", "ShortScanEmail");
        }

        public void OnFullScan(IRibbonControl control)
        {
            ExecuteVbaButton("AES_FULLSCAN", "FullScanEmail");
        }

        public void OnDeepScan(IRibbonControl control)
        {
            ExecuteVbaButton("AES_DEEPSCAN", "DeepScanEmail");
        }

        public void OnGuri(IRibbonControl control)
        {
            // Always direct-launch (or --raise) so GURI starts when it is not running.
            // VBA toolbar path is best-effort only; do not return early on it alone.
            TryExecuteVbaButton("AES_GURI", "ShowGuriGui");
            RaiseOrLaunchGuri("--raise", "GURI");
        }

        public void OnAura(IRibbonControl control)
        {
            TryExecuteVbaButton("AES_AURA", "ShowAuraGui");
            RaiseOrLaunchGuri("--raise --aura", "Aura");
        }

        /// <summary>Shared raise/launch path for GURI and Aura ribbon buttons.</summary>
        private void RaiseOrLaunchGuri(string scriptArgs, string productLabel)
        {
            bool hadWindow = BringGuriWindowToFront();
            try { NativeMethods.AllowSetForegroundWindow(-1); } catch { }
            bool launched = LaunchGuriGuiDirect(scriptArgs);
            if (!launched && !hadWindow)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Could not open " + productLabel + ".\n\nSet Install root in GURI (Database tab) and ensure Python is installed." +
                    "\nOptionally re-import MSCANToolbar so the Add-ins AES bar includes " + productLabel + ".",
                    "Aliniant AES",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }

            if (!hadWindow)
            {
                for (int i = 0; i < 20; i++)
                {
                    Thread.Sleep(100);
                    if (BringGuriWindowToFront())
                        break;
                }
            }
            else
            {
                Thread.Sleep(150);
                BringGuriWindowToFront();
            }
        }

        public void OnDiagnostics(IRibbonControl control)
        {
            ExecuteVbaButton("AES_DIAGNOSTICS", "RunDiagnostics");
        }

        public void OnSettings(IRibbonControl control)
        {
            if (TryExecuteVbaButton("AES_SETTINGS", "ShowSettings"))
                return;
            // Outlook has no Application.Run, and the AES CommandBar button is
            // often missing until VBA is re-imported. Open the Python dialog directly.
            if (LaunchSettingsDialogDirect())
                return;
            System.Windows.Forms.MessageBox.Show(
                "Could not open AES Settings.\n\nSet Install root in GURI (Database tab) and ensure Python is installed.",
                "Aliniant AES",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Warning);
        }

        public void OnViewLogs(IRibbonControl control)
        {
            ExecuteVbaButton("AES_LOGS", "ViewLogs");
        }

        private static string LoadRibbonXml(string tabIdMso)
        {
            var asm = Assembly.GetExecutingAssembly();
            const string name = "Aliniant.AesRibbonHost.Ribbon.xml";
            string raw = null;
            using (var stream = asm.GetManifestResourceStream(name))
            {
                if (stream != null)
                {
                    using (var reader = new StreamReader(stream))
                        raw = reader.ReadToEnd();
                }
            }

            if (string.IsNullOrEmpty(raw))
            {
                var root = InstallRoot();
                var paths = new System.Collections.Generic.List<string>
                {
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ribbon.xml"),
                };
                if (!string.IsNullOrEmpty(root))
                {
                    paths.Add(Path.Combine(root, "AesRibbonHost", "Ribbon.xml"));
                    paths.Add(Path.Combine(root, "VBA", "MSCAN_Ribbon.xml"));
                    paths.Add(Path.Combine(root, "assets", "outlook", "MSCAN_Ribbon.xml"));
                }
                foreach (var path in paths)
                {
                    if (File.Exists(path))
                    {
                        raw = File.ReadAllText(path);
                        break;
                    }
                }
            }

            if (string.IsNullOrEmpty(raw))
                return string.Empty;

            // Template uses __TAB_IDMSO__; older copies may still hard-code TabMail.
            if (raw.Contains("__TAB_IDMSO__"))
                return raw.Replace("__TAB_IDMSO__", tabIdMso);
            return raw.Replace("idMso=\"TabMail\"", "idMso=\"" + tabIdMso + "\"");
        }

        private bool IsServiceOn()
        {
            try
            {
                var path = StateFilePath();
                if (File.Exists(path))
                {
                    var json = File.ReadAllText(path).ToLowerInvariant();
                    if (json.Contains("\"enabled\":true") || json.Contains("\"enabled\": true"))
                        return true;
                    if (json.Contains("\"enabled\":false") || json.Contains("\"enabled\": false"))
                        return false;
                }
            }
            catch { /* fall through */ }

            // Fallback: toolbar caption
            try
            {
                var caption = FindToolbarButtonCaption("AES_SERVICE_TOGGLE");
                if (!string.IsNullOrEmpty(caption))
                    return caption.IndexOf("ON", StringComparison.OrdinalIgnoreCase) >= 0
                           && caption.IndexOf("OFF", StringComparison.OrdinalIgnoreCase) < 0;
            }
            catch { }

            return true;
        }

        private bool IsServiceBusy()
        {
            try
            {
                var path = StateFilePath();
                if (!File.Exists(path)) return false;
                var json = File.ReadAllText(path).ToLowerInvariant();
                return json.Contains("\"busy\":true") || json.Contains("\"busy\": true");
            }
            catch
            {
                return false;
            }
        }

        private int PendingJobCount()
        {
            return ReadStateInt("pending", 0);
        }

        private int WatcherCount()
        {
            // Default -1 = unknown (old state file without "watchers"); callers
            // only treat an explicit 0 as the "ON but nothing watched" state.
            return ReadStateInt("watchers", -1);
        }

        private int ReadStateInt(string keyName, int fallback)
        {
            try
            {
                var path = StateFilePath();
                if (!File.Exists(path)) return fallback;
                var json = File.ReadAllText(path);
                var key = "\"" + keyName + "\"";
                var idx = json.IndexOf(key, StringComparison.OrdinalIgnoreCase);
                if (idx < 0) return fallback;
                var colon = json.IndexOf(':', idx + key.Length);
                if (colon < 0) return fallback;
                var end = colon + 1;
                while (end < json.Length && (char.IsWhiteSpace(json[end]) || json[end] == '"')) end++;
                var start = end;
                while (end < json.Length && char.IsDigit(json[end])) end++;
                if (end > start && int.TryParse(json.Substring(start, end - start), out var n))
                    return n;
            }
            catch { }
            return fallback;
        }

        private static string StateFilePath()
        {
            var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(local, "GeoFooter", "aes_service_state.json");
        }

        /// <summary>
        /// Keep the ribbon in sync when VBA writes busy/ON/OFF without a successful
        /// COM Invalidate (common after async geo jobs finish).
        /// </summary>
        private void StartStateWatch()
        {
            lock (_stateSync)
            {
                try
                {
                    var path = StateFilePath();
                    var dir = Path.GetDirectoryName(path);
                    if (string.IsNullOrEmpty(dir)) return;
                    Directory.CreateDirectory(dir);
                    _lastStateSnapshot = ReadStateSnapshot();

                    if (_stateWatcher == null)
                    {
                        _stateWatcher = new FileSystemWatcher(dir)
                        {
                            Filter = Path.GetFileName(path),
                            NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName,
                            IncludeSubdirectories = false,
                            EnableRaisingEvents = true,
                        };
                        _stateWatcher.Changed += OnStateFileChanged;
                        _stateWatcher.Created += OnStateFileChanged;
                        _stateWatcher.Renamed += OnStateFileRenamed;
                        HostLog.Write("State watcher attached: " + path);
                    }

                    if (_statePollTimer == null)
                    {
                        // Backup poll — some editors replace the file in ways FSW misses.
                        _statePollTimer = new Timer(_ =>
                        {
                            try { RefreshRibbonIfStateChanged(); }
                            catch { /* never throw on timer */ }
                        }, null, 1500, 1500);
                    }
                }
                catch (Exception ex)
                {
                    HostLog.Write("StartStateWatch failed: " + ex.Message);
                }
            }
        }

        private void StopStateWatch()
        {
            lock (_stateSync)
            {
                try
                {
                    if (_statePollTimer != null)
                    {
                        _statePollTimer.Dispose();
                        _statePollTimer = null;
                    }
                    if (_stateWatcher != null)
                    {
                        _stateWatcher.EnableRaisingEvents = false;
                        _stateWatcher.Changed -= OnStateFileChanged;
                        _stateWatcher.Created -= OnStateFileChanged;
                        _stateWatcher.Renamed -= OnStateFileRenamed;
                        _stateWatcher.Dispose();
                        _stateWatcher = null;
                    }
                }
                catch { }
            }
        }

        private void OnStateFileRenamed(object sender, RenamedEventArgs e)
        {
            OnStateFileChanged(sender, e);
        }

        private void OnStateFileChanged(object sender, FileSystemEventArgs e)
        {
            try { RefreshRibbonIfStateChanged(); }
            catch { }
        }

        private void RefreshRibbonIfStateChanged()
        {
            var snap = ReadStateSnapshot();
            if (snap == null) return;
            string prev;
            lock (_stateSync)
            {
                prev = _lastStateSnapshot;
                if (string.Equals(prev, snap, StringComparison.Ordinal))
                    return;
                _lastStateSnapshot = snap;
            }

            // Debounce rapid multi-event writes from Notepad-style replace.
            var now = Environment.TickCount;
            var last = Interlocked.Read(ref _lastInvalidateTicks);
            if (unchecked((int)(now - last)) >= 0 && unchecked((int)(now - last)) < 200)
                return;

            HostLog.Write("State changed → invalidate ribbon (" + snap + ")");
            InvalidateAesRibbon();
        }

        private static string ReadStateSnapshot()
        {
            try
            {
                var path = StateFilePath();
                if (!File.Exists(path)) return string.Empty;
                // Share ReadWrite so we don't block VBA mid-write.
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var sr = new StreamReader(fs))
                    return (sr.ReadToEnd() ?? string.Empty).Trim();
            }
            catch
            {
                return null;
            }
        }

        private static readonly System.Collections.Generic.Dictionary<string, object> IconCache =
            new System.Collections.Generic.Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        private static bool _iconCacheLogged;

        private static object LoadIconPicture(string fileName)
        {
            try
            {
                if (string.IsNullOrEmpty(fileName)) return null;
                object cached;
                lock (IconCache)
                {
                    if (IconCache.TryGetValue(fileName, out cached) && cached != null)
                        return cached;
                }

                // Prefer the .png variant: real alpha channel, so hollow shapes
                // and glows render cleanly on any ribbon theme (BMPs are the
                // legacy VBA CommandBar fallback and draw as opaque squares).
                var pngName = Path.ChangeExtension(fileName, ".png");
                foreach (var dir in IconDirs())
                {
                    var path = Path.Combine(dir, pngName);
                    if (!File.Exists(path)) path = Path.Combine(dir, fileName);
                    if (!File.Exists(path)) continue;
                    using (var img = Image.FromFile(path))
                    {
                        // Keep the bitmap alive independently of the file stream
                        var clone = new Bitmap(img);
                        var pic = PictureConverter.ImageToPictureDisp(clone);
                        lock (IconCache)
                        {
                            IconCache[fileName] = pic;
                        }
                        if (!_iconCacheLogged)
                        {
                            _iconCacheLogged = true;
                            HostLog.Write("LoadIconPicture cache primed: " + path);
                        }
                        return pic;
                    }
                }
                HostLog.Write("LoadIconPicture MISSING: " + fileName);
            }
            catch (Exception ex)
            {
                HostLog.Write("LoadIconPicture FAIL " + fileName + ": " + ex.Message);
            }
            return null;
        }

        private static string[] IconDirs()
        {
            var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var root = InstallRoot();
            var list = new System.Collections.Generic.List<string>();
            if (!string.IsNullOrEmpty(local))
                list.Add(Path.Combine(local, "GeoFooter", "icons"));
            if (!string.IsNullOrEmpty(root))
            {
                list.Add(Path.Combine(root, "assets", "icons"));
            }
            list.Add(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "icons"));
            return list.ToArray();
        }

        /// <summary>
        /// Suite install root from %LOCALAPPDATA%\GeoFooter\install_root.txt
        /// (written by GURI Database tab), then GEOFOOTER_ROOT, then legacy C:\GeoFooter.
        /// </summary>
        private static string InstallRoot()
        {
            try
            {
                var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                var pointer = Path.Combine(local, "GeoFooter", "install_root.txt");
                if (File.Exists(pointer))
                {
                    var line = (File.ReadAllText(pointer).Split('\n')[0] ?? "").Trim().Trim('"');
                    if (!string.IsNullOrEmpty(line) && Directory.Exists(line))
                        return line;
                }
            }
            catch { }

            try
            {
                var env = (Environment.GetEnvironmentVariable("GEOFOOTER_ROOT") ?? "").Trim().Trim('"');
                if (!string.IsNullOrEmpty(env) && Directory.Exists(env))
                    return env;
            }
            catch { }

            // Last-resort legacy folder if still present
            if (Directory.Exists(@"C:\GeoFooter") &&
                (File.Exists(@"C:\GeoFooter\VERSION") ||
                 Directory.Exists(@"C:\GeoFooter\aes") ||
                 Directory.Exists(@"C:\GeoFooter\VBA")))
                return @"C:\GeoFooter";

            return null;
        }

        private void ExecuteVbaButton(string tag, string onActionHint)
        {
            if (TryExecuteVbaButton(tag, onActionHint))
                return;
            if (TryRunVbaMacro(onActionHint))
                return;

            HostLog.Write("ExecuteVbaButton: no matching button/macro for tag=" + tag + " hint=" + onActionHint);
            System.Windows.Forms.MessageBox.Show(
                "AES button '" + onActionHint + "' was not found on the VBA toolbar." +
                "\n\nAfter a VBA re-import: Alt+F11 → Debug → Compile VBAProject," +
                "\nthen Alt+F8 → run RecoverAesUi (or restart Outlook)." +
                "\n\nIf Compile fails, fix errors first (remove duplicate ThisOutlookSession under Modules).",
                "Aliniant AES",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Warning);
        }

        /// <summary>Find AES CommandBar control by tag/OnAction and Execute it.</summary>
        private bool TryExecuteVbaButton(string tag, string onActionHint)
        {
            if (_application == null) return false;

            try
            {
                // Outlook's Application object has no CommandBars — they live on
                // each Explorer window.
                dynamic app = _application;
                try
                {
                    dynamic activeExp = app.ActiveExplorer();
                    if (activeExp != null && TryExecuteOnBars(activeExp.CommandBars, tag, onActionHint))
                    {
                        InvalidateAesRibbon();
                        return true;
                    }
                }
                catch { }

                foreach (dynamic exp in app.Explorers)
                {
                    try
                    {
                        if (TryExecuteOnBars(exp.CommandBars, tag, onActionHint))
                        {
                            InvalidateAesRibbon();
                            return true;
                        }
                    }
                    catch { }
                }

                HostLog.Write("TryExecuteVbaButton: no matching button for tag=" + tag);
                return false;
            }
            catch (Exception ex)
            {
                HostLog.Write("TryExecuteVbaButton FAIL tag=" + tag + ": " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Call a Public VBA Sub via Application.Run when the CommandBar button is missing
        /// (stale toolbar / GURI not yet on the bar after an older CreateToolbar).
        /// </summary>
        private bool TryRunVbaMacro(string procName)
        {
            if (_application == null || string.IsNullOrWhiteSpace(procName)) return false;

            string bare = procName.Trim();
            int dot = bare.LastIndexOf('.');
            if (dot >= 0 && dot < bare.Length - 1)
                bare = bare.Substring(dot + 1);

            string project = "Project1";
            try
            {
                dynamic app = _application;
                try
                {
                    string name = app.VBE.ActiveVBProject.Name as string;
                    if (!string.IsNullOrEmpty(name)) project = name;
                }
                catch { }

                foreach (var candidate in new[] {
                    project + ".MSCANToolbar." + bare,
                    "MSCANToolbar." + bare,
                    project + "." + bare,
                    bare })
                {
                    try
                    {
                        app.Run(candidate);
                        HostLog.Write("TryRunVbaMacro OK: " + candidate);
                        InvalidateAesRibbon();
                        return true;
                    }
                    catch (Exception ex)
                    {
                        HostLog.Write("TryRunVbaMacro miss " + candidate + ": " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                HostLog.Write("TryRunVbaMacro FAIL: " + ex.Message);
            }
            return false;
        }

        /// <summary>
        /// Raise or start guri\gui.py without VBA (same paths as MSCANToolbar.ShowGuriGui).
        /// </summary>
        private static bool LaunchGuriGuiDirect(string extraArgs = "--raise")
        {
            try
            {
                string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                string root = InstallRoot() ?? "";

                var pyList = new System.Collections.Generic.List<string>();
                if (!string.IsNullOrEmpty(root))
                {
                    pyList.Add(Path.Combine(root, @".venv\Scripts\pythonw.exe"));
                    pyList.Add(Path.Combine(root, @".venv\Scripts\python.exe"));
                }
                pyList.Add(Path.Combine(userProfile, @"AppData\Local\Programs\Python\Python313\pythonw.exe"));
                pyList.Add(Path.Combine(userProfile, @"AppData\Local\Programs\Python\Python312\pythonw.exe"));
                pyList.Add(Path.Combine(userProfile, @"AppData\Local\Programs\Python\Python311\pythonw.exe"));
                pyList.Add(Path.Combine(local, @"Programs\Python\Python313\pythonw.exe"));
                pyList.Add(Path.Combine(local, @"Programs\Python\Python312\pythonw.exe"));
                pyList.Add(@"C:\Python313\pythonw.exe");
                pyList.Add(@"C:\Python312\pythonw.exe");
                string[] pyCandidates = pyList.ToArray();

                var scriptList = new System.Collections.Generic.List<string>();
                if (!string.IsNullOrEmpty(root))
                {
                    scriptList.Add(Path.Combine(root, "guri", "gui.py"));
                    scriptList.Add(Path.Combine(root, "guri_gui.py")); // legacy filename
                }
                scriptList.Add(@"C:\GeoFooter\guri\gui.py");
                string[] scriptCandidates = scriptList.ToArray();

                string py = null;
                foreach (var p in pyCandidates)
                {
                    if (File.Exists(p)) { py = p; break; }
                }

                string script = null;
                foreach (var s in scriptCandidates)
                {
                    if (File.Exists(s)) { script = s; break; }
                }

                if (py == null || script == null)
                {
                    HostLog.Write("LaunchGuriGuiDirect: missing py=" + (py ?? "") +
                        " script=" + (script ?? "") + " root=" + root);
                    return false;
                }

                if (string.IsNullOrEmpty(root))
                {
                    try { root = Path.GetDirectoryName(Path.GetDirectoryName(script)) ?? ""; }
                    catch { root = @"C:\GeoFooter"; }
                }

                string args = "\"" + script + "\"";
                if (!string.IsNullOrWhiteSpace(extraArgs))
                    args += " " + extraArgs.Trim();

                var psi = new ProcessStartInfo
                {
                    FileName = py,
                    Arguments = args,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = string.IsNullOrEmpty(root) ? Environment.CurrentDirectory : root,
                };
                Process.Start(psi);
                HostLog.Write("LaunchGuriGuiDirect: " + py + " " + args + " cwd=" + psi.WorkingDirectory);
                return true;
            }
            catch (Exception ex)
            {
                HostLog.Write("LaunchGuriGuiDirect FAIL: " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Open aes/settings_dialog.py without VBA. The ribbon cannot call
        /// Outlook.Application.Run, and the Settings toolbar button is often absent.
        /// </summary>
        private bool LaunchSettingsDialogDirect()
        {
            try
            {
                string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                string root = InstallRoot() ?? "";
                string geo = Path.Combine(local, "GeoFooter");
                Directory.CreateDirectory(geo);

                string py = FindPythonw(root, userProfile, local);
                string script = null;
                if (!string.IsNullOrEmpty(root))
                {
                    string candidate = Path.Combine(root, "aes", "settings_dialog.py");
                    if (File.Exists(candidate)) script = candidate;
                }
                if (script == null && File.Exists(@"C:\GeoFooter\aes\settings_dialog.py"))
                    script = @"C:\GeoFooter\aes\settings_dialog.py";
                if (py == null || script == null)
                {
                    HostLog.Write("LaunchSettingsDialogDirect: missing py=" + (py ?? "") +
                        " script=" + (script ?? ""));
                    return false;
                }

                string accountsPath = Path.Combine(geo, "aes_settings_accounts.json");
                string outPath = Path.Combine(geo, "aes_settings_result.json");
                string routePath = Path.Combine(geo, "aes_risk_route.json");
                string loggingPath = Path.Combine(geo, "aes_logging.json");
                File.WriteAllText(accountsPath, BuildAccountsJson(), new UTF8Encoding(false));
                try { if (File.Exists(outPath)) File.Delete(outPath); } catch { }

                string args = "\"" + script + "\"" +
                    " --accounts \"" + accountsPath + "\"" +
                    " --out \"" + outPath + "\"" +
                    " --route \"" + routePath + "\"" +
                    " --logging \"" + loggingPath + "\"";
                var psi = new ProcessStartInfo
                {
                    FileName = py,
                    Arguments = args,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = string.IsNullOrEmpty(root) ? Environment.CurrentDirectory : root,
                };
                Process.Start(psi);
                HostLog.Write("LaunchSettingsDialogDirect: " + py + " " + args);
                return true;
            }
            catch (Exception ex)
            {
                HostLog.Write("LaunchSettingsDialogDirect FAIL: " + ex.Message);
                return false;
            }
        }

        private static string FindPythonw(string root, string userProfile, string local)
        {
            var candidates = new System.Collections.Generic.List<string>();
            if (!string.IsNullOrEmpty(root))
            {
                candidates.Add(Path.Combine(root, @".venv\Scripts\pythonw.exe"));
                candidates.Add(Path.Combine(root, @".venv\Scripts\python.exe"));
            }
            candidates.Add(Path.Combine(userProfile, @"AppData\Local\Programs\Python\Python313\pythonw.exe"));
            candidates.Add(Path.Combine(local, @"Programs\Python\Python313\pythonw.exe"));
            candidates.Add(Path.Combine(userProfile, @"AppData\Local\Programs\Python\Python312\pythonw.exe"));
            foreach (string path in candidates)
            {
                if (File.Exists(path)) return path;
            }
            return null;
        }

        private string BuildAccountsJson()
        {
            var flags = ReadScanAccountFlags();
            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine("  \"accounts\": [");
            bool first = true;
            try
            {
                dynamic app = _application;
                dynamic accounts = app.Session.Accounts;
                int count = (int)accounts.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic acc = null;
                    try { acc = accounts.Item(i); } catch { continue; }
                    if (acc == null) continue;
                    string storeId = "";
                    try
                    {
                        dynamic store = acc.DeliveryStore;
                        if (store != null) storeId = (store.StoreID as string) ?? "";
                    }
                    catch { }
                    if (string.IsNullOrEmpty(storeId))
                    {
                        try { storeId = (acc.StoreID as string) ?? ""; } catch { }
                    }
                    if (string.IsNullOrEmpty(storeId)) continue;
                    string display = "";
                    string smtp = "";
                    try { display = (acc.DisplayName as string) ?? ""; } catch { }
                    try { smtp = (acc.SmtpAddress as string) ?? ""; } catch { }
                    if (string.IsNullOrEmpty(smtp)) smtp = display;
                    bool enabled = true, responses = true, inCc = true;
                    if (flags.TryGetValue(storeId, out var flag))
                    {
                        enabled = flag.Item1;
                        responses = flag.Item2;
                        inCc = flag.Item3;
                    }
                    if (!first) sb.AppendLine(",");
                    first = false;
                    sb.Append("    {\"store_id\":\"").Append(JsonEscape(storeId))
                        .Append("\",\"display\":\"").Append(JsonEscape(display))
                        .Append("\",\"smtp\":\"").Append(JsonEscape(smtp))
                        .Append("\",\"enabled\":").Append(enabled ? "true" : "false")
                        .Append(",\"responses\":").Append(responses ? "true" : "false")
                        .Append(",\"in_cc\":").Append(inCc ? "true" : "false")
                        .Append("}");
                }
            }
            catch (Exception ex)
            {
                HostLog.Write("BuildAccountsJson: " + ex.Message);
            }
            sb.AppendLine();
            sb.AppendLine("  ]");
            sb.AppendLine("}");
            return sb.ToString();
        }

        private static System.Collections.Generic.Dictionary<string, Tuple<bool, bool, bool>> ReadScanAccountFlags()
        {
            var map = new System.Collections.Generic.Dictionary<string, Tuple<bool, bool, bool>>(
                StringComparer.OrdinalIgnoreCase);
            try
            {
                string path = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "GeoFooter", "aes_scan_accounts.ini");
                if (!File.Exists(path)) return map;
                foreach (string raw in File.ReadAllLines(path))
                {
                    string line = (raw ?? "").Trim();
                    if (line.Length == 0 || line[0] == ';' || line[0] == '#') continue;
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    string key = line.Substring(0, eq).Trim();
                    int pipe = key.IndexOf('|');
                    if (pipe > 0) key = key.Substring(0, pipe).Trim();
                    if (key.Length == 0) continue;
                    string[] parts = line.Substring(eq + 1).Trim().Split(',');
                    bool on0 = FlagOn(parts, 0, true);
                    bool on1 = FlagOn(parts, 1, true);
                    bool on2 = FlagOn(parts, 2, true);
                    map[key] = Tuple.Create(on0, on1, on2);
                }
            }
            catch { }
            return map;
        }

        private static bool FlagOn(string[] parts, int index, bool fallback)
        {
            if (parts == null || index >= parts.Length) return fallback;
            string text = (parts[index] ?? "").Trim();
            if (text.Length == 0) return fallback;
            return text == "1" || text.Equals("true", StringComparison.OrdinalIgnoreCase)
                || text.Equals("on", StringComparison.OrdinalIgnoreCase);
        }

        private static string JsonEscape(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            return value.Replace("\\", "\\\\").Replace("\"", "\\\"")
                .Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
        }

        /// <summary>
        /// Find the GURI main window and pull it forward. Must run from the Outlook
        /// ribbon callback (user input) or Windows will ignore SetForegroundWindow.
        /// </summary>
        private static bool BringGuriWindowToFront()
        {
            try
            {
                IntPtr found = IntPtr.Zero;
                NativeMethods.EnumWindows((hwnd, lParam) =>
                {
                    var sb = new StringBuilder(512);
                    NativeMethods.GetWindowText(hwnd, sb, sb.Capacity);
                    string title = sb.ToString();
                    if (string.IsNullOrEmpty(title)) return true;
                    // Match main window; ignore tray balloons / helper dialogs.
                    if (title.IndexOf("GURI Database Viewer", StringComparison.OrdinalIgnoreCase) < 0)
                        return true;
                    found = hwnd;
                    return false;
                }, IntPtr.Zero);

                if (found == IntPtr.Zero)
                {
                    HostLog.Write("BringGuriWindowToFront: no GURI window found");
                    return false;
                }

                uint pid;
                NativeMethods.GetWindowThreadProcessId(found, out pid);
                if (pid != 0)
                    NativeMethods.AllowSetForegroundWindow(unchecked((int)pid));

                if (NativeMethods.IsIconic(found))
                    NativeMethods.ShowWindow(found, NativeMethods.SW_RESTORE);
                else
                    NativeMethods.ShowWindow(found, NativeMethods.SW_SHOW);

                // Attach to the current foreground thread so Windows allows the
                // focus change (SetForegroundWindow alone is often ignored).
                IntPtr fg = NativeMethods.GetForegroundWindow();
                uint fgTid = NativeMethods.GetWindowThreadProcessId(fg, IntPtr.Zero);
                uint thisTid = NativeMethods.GetCurrentThreadId();
                bool attached = false;
                try
                {
                    if (fg != IntPtr.Zero && fgTid != 0 && fgTid != thisTid)
                        attached = NativeMethods.AttachThreadInput(thisTid, fgTid, true);

                    NativeMethods.BringWindowToTop(found);
                    NativeMethods.SetWindowPos(
                        found, NativeMethods.HWND_TOPMOST, 0, 0, 0, 0,
                        NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_SHOWWINDOW);
                    NativeMethods.SetWindowPos(
                        found, NativeMethods.HWND_NOTOPMOST, 0, 0, 0, 0,
                        NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_SHOWWINDOW);
                    bool ok = NativeMethods.SetForegroundWindow(found);
                    HostLog.Write("BringGuriWindowToFront hwnd=" + found.ToInt64() + " pid=" + pid + " ok=" + ok);
                }
                finally
                {
                    if (attached)
                        NativeMethods.AttachThreadInput(thisTid, fgTid, false);
                }
                return true;
            }
            catch (Exception ex)
            {
                HostLog.Write("BringGuriWindowToFront FAIL: " + ex.Message);
                return false;
            }
        }

        private static class NativeMethods
        {
            public const int SW_RESTORE = 9;
            public const int SW_SHOW = 5;
            public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
            public static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
            public const uint SWP_NOSIZE = 0x0001;
            public const uint SWP_NOMOVE = 0x0002;
            public const uint SWP_SHOWWINDOW = 0x0040;

            public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

            [DllImport("user32.dll")]
            public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

            [DllImport("user32.dll", CharSet = CharSet.Unicode)]
            public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

            [DllImport("user32.dll")]
            public static extern bool IsIconic(IntPtr hWnd);

            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);

            [DllImport("user32.dll")]
            public static extern bool BringWindowToTop(IntPtr hWnd);

            [DllImport("user32.dll")]
            public static extern IntPtr GetForegroundWindow();

            [DllImport("user32.dll")]
            public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

            [DllImport("user32.dll")]
            public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

            [DllImport("user32.dll")]
            public static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr processId);

            [DllImport("user32.dll")]
            public static extern bool AllowSetForegroundWindow(int dwProcessId);

            [DllImport("user32.dll")]
            public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

            [DllImport("kernel32.dll")]
            public static extern uint GetCurrentThreadId();
        }

        private static bool TryExecuteOnBars(dynamic commandBars, string tag, string onActionHint)
        {
            if (commandBars == null) return false;
            int count;
            try { count = (int)commandBars.Count; }
            catch { return false; }

            for (int i = 1; i <= count; i++)
            {
                dynamic bar = null;
                try
                {
                    bar = commandBars[i];
                    string name = bar.Name as string ?? "";
                    if (!name.Equals("AES", StringComparison.OrdinalIgnoreCase)
                        && name.IndexOf("AES", StringComparison.OrdinalIgnoreCase) < 0
                        && !name.Equals("Aliniant AES", StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (TryExecuteOnBar(bar, tag, onActionHint))
                        return true;
                }
                catch { }
            }

            // Also try by name directly
            foreach (var barName in new[] { "AES", "Aliniant AES" })
            {
                try
                {
                    dynamic bar = commandBars[barName];
                    if (bar != null && TryExecuteOnBar(bar, tag, onActionHint))
                        return true;
                }
                catch { }
            }
            return false;
        }

        private static bool TryExecuteOnBar(dynamic bar, string tag, string onActionHint)
        {
            int n;
            try { n = (int)bar.Controls.Count; }
            catch { return false; }

            for (int i = 1; i <= n; i++)
            {
                try
                {
                    dynamic ctl = bar.Controls[i];
                    string ctlTag = "";
                    string onAction = "";
                    try { ctlTag = ctl.Tag as string ?? ""; } catch { }
                    try { onAction = ctl.OnAction as string ?? ""; } catch { }

                    bool match = (!string.IsNullOrEmpty(tag) && ctlTag.Equals(tag, StringComparison.OrdinalIgnoreCase))
                                 || (!string.IsNullOrEmpty(onActionHint)
                                     && onAction.IndexOf(onActionHint, StringComparison.OrdinalIgnoreCase) >= 0);
                    if (!match) continue;

                    ctl.Execute();
                    return true;
                }
                catch { }
            }
            return false;
        }

        private string FindToolbarButtonCaption(string tag)
        {
            if (_application == null) return null;
            try
            {
                dynamic app = _application;
                try
                {
                    dynamic activeExp = app.ActiveExplorer();
                    if (activeExp != null)
                    {
                        var activeCap = FindCaptionOnBars(activeExp.CommandBars, tag);
                        if (!string.IsNullOrEmpty(activeCap)) return activeCap;
                    }
                }
                catch { }
                foreach (dynamic exp in app.Explorers)
                {
                    try
                    {
                        var cap = FindCaptionOnBars(exp.CommandBars, tag);
                        if (!string.IsNullOrEmpty(cap)) return cap;
                    }
                    catch { }
                }
            }
            catch { }
            return null;
        }

        private static string FindCaptionOnBars(dynamic commandBars, string tag)
        {
            if (commandBars == null) return null;
            foreach (var barName in new[] { "AES", "Aliniant AES" })
            {
                try
                {
                    dynamic bar = commandBars[barName];
                    if (bar == null) continue;
                    int n = (int)bar.Controls.Count;
                    for (int i = 1; i <= n; i++)
                    {
                        dynamic ctl = bar.Controls[i];
                        string ctlTag = "";
                        try { ctlTag = ctl.Tag as string ?? ""; } catch { }
                        if (ctlTag.Equals(tag, StringComparison.OrdinalIgnoreCase))
                            return ctl.Caption as string;
                    }
                }
                catch { }
            }
            return null;
        }
    }
}
