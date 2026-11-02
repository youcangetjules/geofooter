# If Outlook's VbaProject.OTM is missing (or tiny) and a GeoFooter backup exists,
# restore it. Outlook must be closed. Pair with Start_Outlook_AES.bat.
#
# Usage:
#   .\Ensure_VbaProject_OTM.ps1
#   .\Ensure_VbaProject_OTM.ps1 -MinBytes 50000

[CmdletBinding()]
param([int]$MinBytes = 50000)

$ErrorActionPreference = "Stop"
$otm = Join-Path $env:APPDATA "Microsoft\Outlook\VbaProject.OTM"
$bakDir = Join-Path $env:LOCALAPPDATA "GeoFooter\vba_otm_backups"

if (Get-Process OUTLOOK -EA SilentlyContinue) {
    throw "Close Outlook completely (tray too) before Ensure_VbaProject_OTM."
}

$needRestore = $false
if (-not (Test-Path $otm)) {
    Write-Host "OTM missing."
    $needRestore = $true
} else {
    $len = (Get-Item $otm).Length
    Write-Host "OTM present size=$len"
    if ($len -lt $MinBytes) {
        Write-Host "OTM smaller than $MinBytes — treating as empty."
        $needRestore = $true
    }
}

if (-not $needRestore) {
    Write-Host "OTM OK — nothing to restore."
    exit 0
}

$src = Get-ChildItem $bakDir -Filter "*.otm" -EA SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if (-not $src) {
    throw "No backup under $bakDir. Import VBA once, File > Save, quit Outlook, run Backup_VbaProject_OTM.bat."
}

New-Item -ItemType Directory -Force -Path (Split-Path $otm) | Out-Null
Copy-Item -LiteralPath $src.FullName -Destination $otm -Force
Write-Host "Restored $($src.Name) -> $otm ($((Get-Item $otm).Length) bytes)"
exit 0
