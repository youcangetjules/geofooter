# Backup or restore Outlook's VbaProject.OTM (where AES VBA actually lives).
# Outlook "losing modules" almost always means this file was never saved
# or was deleted — Project Explorer is only the in-memory copy.
#
# Usage:
#   .\Backup_VbaProject_OTM.ps1              # copy current OTM to backups\
#   .\Backup_VbaProject_OTM.ps1 -Status      # show OTM + backups
#   .\Backup_VbaProject_OTM.ps1 -Restore     # restore newest backup (Outlook must be CLOSED)
#   .\Backup_VbaProject_OTM.ps1 -Restore -BackupFile "....otm"

[CmdletBinding()]
param(
    [switch]$Status,
    [switch]$Restore,
    [string]$BackupFile = ""
)

$ErrorActionPreference = "Stop"
$otm = Join-Path $env:APPDATA "Microsoft\Outlook\VbaProject.OTM"
$bakDir = Join-Path $env:LOCALAPPDATA "GeoFooter\vba_otm_backups"
New-Item -ItemType Directory -Force -Path $bakDir | Out-Null

function Show-Status {
    Write-Host "OTM: $otm"
    if (Test-Path $otm) {
        $i = Get-Item $otm
        Write-Host ("  EXISTS  size={0:N0} bytes  written={1}" -f $i.Length, $i.LastWriteTime)
        if ($i.Length -lt 50000) {
            Write-Host "  WARNING: small OTM often means empty/near-empty project."
        }
    } else {
        Write-Host "  MISSING — Outlook will open an empty Project1 until you import + File > Save."
    }
    Write-Host "Backups: $bakDir"
    Get-ChildItem $bakDir -Filter "*.otm" -EA SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 10 |
        ForEach-Object { Write-Host ("  {0:yyyy-MM-dd HH:mm}  {1,10:N0}  {2}" -f $_.LastWriteTime, $_.Length, $_.Name) }
}

if ($Status) {
    Show-Status
    exit 0
}

if ($Restore) {
    if (Get-Process OUTLOOK -EA SilentlyContinue) {
        throw "Close Outlook completely (tray too) before restoring OTM."
    }
    $src = $BackupFile
    if (-not $src) {
        $src = Get-ChildItem $bakDir -Filter "*.otm" -EA SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1 -ExpandProperty FullName
    }
    if (-not $src -or -not (Test-Path $src)) {
        throw "No backup OTM found under $bakDir"
    }
    $destDir = Split-Path $otm -Parent
    New-Item -ItemType Directory -Force -Path $destDir | Out-Null
    Copy-Item -LiteralPath $src -Destination $otm -Force
    Write-Host "Restored:`n  $src`n-> $otm"
    Show-Status
    Write-Host "Start Outlook, Alt+F11 — Modules should be back."
    exit 0
}

# default: backup
if (-not (Test-Path $otm)) {
    Show-Status
    throw "Nothing to back up — VbaProject.OTM is missing. Import modules, File > Save in VBE, quit Outlook, then re-run backup."
}
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$len = (Get-Item $otm).Length
$dest = Join-Path $bakDir "VbaProject_$stamp`_$len.otm"
Copy-Item -LiteralPath $otm -Destination $dest -Force
Write-Host "Backed up to $dest"
Show-Status
exit 0
