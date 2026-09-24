# Kill stuck GeoFooter / AES / GURI Python processes, then restart guri\gui.py.
# Invoked by Restart_guri_gui.bat. -Root is the suite install root (parent of scripts\).

param(
    [string]$Root = (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
)

$ErrorActionPreference = "SilentlyContinue"
# Bat/PowerShell quoting can leave a trailing " when %~dp0 ends with \.
$Root = ($Root -replace '"', '').Trim().TrimEnd("\", "/")
if (-not $Root) {
    $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
}

Write-Host "========================================"
Write-Host "  Restart GURI GUI"
Write-Host "========================================"
Write-Host ""
Write-Host "Stopping stuck GeoFooter Python processes..."
Write-Host "(only processes whose command line references GeoFooter / AES / GURI)"
Write-Host ""

$patterns = @(
    [regex]::Escape(($Root -replace "\\", "/")),
    [regex]::Escape($Root),
    "geofooter",
    "geolocate_headers\.py",
    "guri[\\/](gui|gui_service|outlook_scraper|file_scraper|learning)\.py",
    "aes[\\/](settings_dialog|action_handler|classify_dialog|diagnostics_dialog)\.py",
    "guri_gui\.py",
    "aes_\w+_dialog\.py",
    "aes_action_handler\.py"
)

$killed = 0
Get-CimInstance Win32_Process -Filter "Name='python.exe' OR Name='pythonw.exe'" | ForEach-Object {
    $cmd = [string]$_.CommandLine
    if ([string]::IsNullOrWhiteSpace($cmd)) { return }
    $hit = $false
    foreach ($p in $patterns) {
        if ($cmd -imatch $p) { $hit = $true; break }
    }
    if (-not $hit) { return }
    Write-Host ("  Killing PID {0}  {1}" -f $_.ProcessId, $cmd)
    Stop-Process -Id $_.ProcessId -Force
    $script:killed++
}
Write-Host ("Stopped {0} process(es)." -f $killed)
Write-Host ""

Start-Sleep -Seconds 2

$gui = Join-Path $Root "guri\gui.py"
if (-not (Test-Path $gui)) {
    Write-Host "ERROR: guri\gui.py not found under $Root"
    exit 1
}

# Guard against truncated / corrupt files that exit silently under pythonw.
$tail = Get-Content $gui -Tail 30 -ErrorAction SilentlyContinue | Out-String
if ($tail -notmatch 'if __name__') {
    Write-Host "ERROR: guri\gui.py looks truncated/corrupt (no __main__ entry point)."
    Write-Host "Restore from Cursor local history or a backup before restarting."
    exit 1
}

$pyCandidates = @(
    (Join-Path $Root ".venv\Scripts\pythonw.exe"),
    "C:\Python313\pythonw.exe",
    "C:\Python313\python.exe",
    (Get-Command pythonw -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1),
    (Get-Command python -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1)
) | Where-Object { $_ -and (Test-Path $_) }

$py = $pyCandidates | Select-Object -First 1
if (-not $py) {
    Write-Host "ERROR: Python not found. Install Python or fix PATH."
    exit 1
}

Write-Host "Starting GURI GUI..."
Write-Host "  $py"
Write-Host "  $gui"
Start-Process -FilePath $py -ArgumentList "`"$gui`"" -WorkingDirectory $Root
Write-Host ""
Write-Host "GURI GUI relaunch requested."
exit 0
