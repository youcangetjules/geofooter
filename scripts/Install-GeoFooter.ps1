# Master installer for GeoFooter / AES / GURI / Aura on Windows.
# PostgreSQL is required for GURI. This script does not install PostgreSQL.
param(
    [string]$Root = "",
    [switch]$Yes
)

$ErrorActionPreference = "Stop"

if (-not $Root) {
    $Root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}
$Root = (Resolve-Path $Root).Path

function Write-Banner([string]$Text) {
    Write-Host ""
    Write-Host ("=" * 72)
    Write-Host $Text
    Write-Host ("=" * 72)
}

function Confirm-Step([string]$Prompt) {
    if ($Yes) { return $true }
    $answer = Read-Host "$Prompt [Y/n]"
    return ($answer -eq "" -or $answer -match '^(y|yes)$')
}

function Find-PythonLauncher {
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) { return @{ Exe = $py.Source; Prefix = @("-3") } }
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) { return @{ Exe = $python.Source; Prefix = @() } }
    return $null
}

function Test-Postgres {
    $service = $false
    try {
        $hits = @(Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'postgres' })
        $service = $hits.Count -gt 0
    } catch {
        $service = $false
    }
    $psql = [bool](Get-Command psql -ErrorAction SilentlyContinue)
    $port = $false
    $client = $null
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $wait = $client.BeginConnect("127.0.0.1", 5432, $null, $null)
        if ($wait.AsyncWaitHandle.WaitOne(700) -and $client.Connected) {
            $port = $true
        }
    } catch {
        $port = $false
    } finally {
        if ($client) { $client.Close() }
    }
    return @{ Service = $service; Psql = $psql; Port = $port }
}

Write-Banner "GeoFooter installer"
Write-Host "This installs both:"
Write-Host "  1. AES email tool  - Outlook scanner, aes:// actions, VBA, ribbon"
Write-Host "  2. GURI GUI        - desktop app (guri\gui.py), including the Aura tab"
Write-Host "Install root: $Root"
Write-Host ""
Write-Host "PostgreSQL is required."
Write-Host "GURI stores records in PostgreSQL. SQLite and MySQL are deprecated."
Write-Host "Install PostgreSQL for Windows, create a database named guri_db,"
Write-Host "and put the password in guri_postgres_config.json at the install root."
Write-Host "Download: https://www.postgresql.org/download/windows/"
Write-Host "Example file uses localhost, user postgres, database guri_db, port 5432."

$pg = Test-Postgres
if ($pg.Service -or $pg.Psql -or $pg.Port) {
    $bits = @()
    if ($pg.Service) { $bits += "Windows service" }
    if ($pg.Psql) { $bits += "psql on PATH" }
    if ($pg.Port) { $bits += "localhost:5432 accepting connections" }
    Write-Host ""
    Write-Host ("PostgreSQL looks present (" + ($bits -join ", ") + ").")
    Write-Host "You still need the guri_db database and a password in guri_postgres_config.json."
} else {
    Write-Host ""
    Write-Host "PostgreSQL was not detected on this PC (no service, no psql, port 5432 closed)."
    Write-Host "Install it before you open GURI. Setup of Python and Outlook can continue now."
}

$example = Join-Path $Root "guri_postgres_config.example.json"
$config = Join-Path $Root "guri_postgres_config.json"
if (-not (Test-Path $config)) {
    if (-not (Test-Path $example)) {
        throw "Missing $example"
    }
    Copy-Item $example $config
    Write-Host ""
    Write-Host "Copied guri_postgres_config.example.json to guri_postgres_config.json."
    Write-Host "Edit that file and set the PostgreSQL password. Do not commit it."
} else {
    Write-Host ""
    Write-Host "guri_postgres_config.json already exists; left it unchanged."
}

Write-Banner "Python packages"
$venvPy = Join-Path $Root ".venv\Scripts\python.exe"
if (-not (Test-Path $venvPy)) {
    if (-not (Confirm-Step "Create .venv and install requirements.txt?")) {
        throw "Python environment is required. Re-run and accept the venv step."
    }
    $launcher = Find-PythonLauncher
    if (-not $launcher) {
        throw "Python 3.10+ was not found. Install it from https://www.python.org/downloads/ and re-run."
    }
    Write-Host "Creating .venv ..."
    $venvArgs = @()
    $venvArgs += $launcher.Prefix
    $venvArgs += @("-m", "venv", (Join-Path $Root ".venv"))
    & $launcher.Exe @venvArgs
    if ($LASTEXITCODE -ne 0) { throw "venv creation failed ($LASTEXITCODE)" }
}
if (-not (Test-Path $venvPy)) { throw "Expected $venvPy" }

Write-Host "Installing requirements.txt ..."
& $venvPy -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { throw "pip upgrade failed ($LASTEXITCODE)" }
& $venvPy -m pip install -r (Join-Path $Root "requirements.txt")
if ($LASTEXITCODE -ne 0) { throw "pip install failed ($LASTEXITCODE)" }

Write-Banner "Install root"
$env:GEOFOOTER_INSTALL_TARGET = $Root
Push-Location $Root
try {
    & $venvPy -c "import os; from geofooter.paths import set_install_root; print(set_install_root(os.environ['GEOFOOTER_INSTALL_TARGET']))"
    if ($LASTEXITCODE -ne 0) { throw "set_install_root failed ($LASTEXITCODE)" }
} finally {
    Pop-Location
}
Write-Host "Wrote %LOCALAPPDATA%\GeoFooter\install_root.txt"

Write-Host "Checking that the email tool and GURI GUI both import ..."
& $venvPy -c "import aes, guri.gui, aura, geofooter"
if ($LASTEXITCODE -ne 0) { throw "Package import check failed ($LASTEXITCODE)" }
Write-Host "aes, guri, aura, and geofooter import correctly."

Write-Banner "1. AES email tool"
Write-Host "AES runs inside classic Outlook. These steps register footer actions,"
Write-Host "import the VBA scanner, and install the ribbon."

Write-Banner "aes:// actions"
if (Confirm-Step "Register the aes:// protocol for this user?") {
    & $venvPy (Join-Path $Root "aes\action_handler.py") "--register"
    if ($LASTEXITCODE -ne 0) { throw "aes:// registration failed ($LASTEXITCODE)" }
}

Write-Banner "Outlook VBA"
Write-Host "Outlook should be open. First-time import needs Trust access to the VBA project object model."
if (Confirm-Step "Import VBA modules into Outlook now?") {
    $bat = Join-Path $PSScriptRoot "Import_VBA_to_Outlook.bat"
    & cmd.exe /c $bat
    if ($LASTEXITCODE -ne 0) { throw "VBA import failed ($LASTEXITCODE)" }
}

Write-Banner "Outlook ribbon"
Write-Host "Close Outlook completely (including the tray icon) before this step."
Write-Host "Building the ribbon needs the .NET Framework 4.8 SDK (dotnet)."
if (Confirm-Step "Build and register AesRibbonHost now?") {
    $installer = Join-Path $Root "AesRibbonHost\install.ps1"
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $installer
    if ($LASTEXITCODE -ne 0) { throw "Ribbon install failed ($LASTEXITCODE)" }
}

Write-Banner "2. GURI GUI"
Write-Host "GURI is the desktop app. The same .venv has PySide6 and the GURI packages."
Write-Host "PostgreSQL must be running, with database guri_db and the password in"
Write-Host "guri_postgres_config.json, or the GUI will open and stay disconnected."

$desktop = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktop "GURI.lnk"
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $venvPy
$shortcut.Arguments = '"' + (Join-Path $Root "guri\gui.py") + '"'
$shortcut.WorkingDirectory = $Root
$shortcut.Description = "GURI GUI"
$shortcut.Save()
Write-Host "Desktop shortcut: $shortcutPath"
Write-Host "Launcher: scripts\Launch_GURI_GUI.bat"

Write-Banner "Done"
Write-Host "AES email tool: reopen Outlook after a ribbon install. New mail is scanned from there."
Write-Host "GURI GUI: double-click the desktop shortcut, or run scripts\Launch_GURI_GUI.bat."
Write-Host "PostgreSQL is still required: database guri_db, password in guri_postgres_config.json."
