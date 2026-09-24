# Kill stuck GeoFooter / AES / GURI Python processes, then restart guri_gui.
# Invoked by Restart_guri_gui.bat

param(
    [string]$Root = (Split-Path -Parent $MyInvocation.MyCommand.Path)
)

$ErrorActionPreference = "SilentlyContinue"
# Bat/PowerShell quoting can leave a trailing " when %~dp0 ends with \.
$Root = ($Root -replace '"', '').Trim().TrimEnd("\", "/")
if (-not $Root) {
    $Root = Split-Path -Parent $MyInvocation.MyCommand.Path
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
    "guri_gui\.py",
    "guri_gui_service\.py",
    "guri_outlook_scraper\.py",
    "guri_file_scraper\.py",
    "guri_learning\.py",
    "aes_settings_dialog\.py",
    "aes_action_handler\.py",
    "aes_classify_dialog\.py",
    "aes_diagnostics_dialog\.py"
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

$gui = Join-Path $Root "guri_gui.py"
if (-not (Test-Path $gui)) {
    Write-Host "ERROR: guri_gui.py not found in $Root"
    exit 1
}

# Guard against truncated / corrupt files that exit silently under pythonw.
$tail = Get-Content $gui -Tail 30 -ErrorAction SilentlyContinue | Out-String
if ($tail -notmatch 'if __name__') {
    Write-Host "ERROR: guri_gui.py looks truncated/corrupt (no __main__ entry point)."
    Write-Host "Restore from Cursor local history or a backup before restarting."
    exit 1
}

$pyCandidates = @(
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
