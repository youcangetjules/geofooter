$ErrorActionPreference = "Continue"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$dll = Join-Path $root "bin\Release\AesRibbonHost.dll"
$regasm = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
if (-not (Test-Path $regasm)) {
    $regasm = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
}

# Remove per-user COM bits
$clsid = "{A1E50001-AE51-4B0B-9C11-AE5B1BB00001}"
Remove-Item "HKCU:\Software\Classes\Aliniant.AesRibbonHost" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKCU:\Software\Classes\CLSID\$clsid" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKCU:\Software\Microsoft\Office\Outlook\Addins\Aliniant.AesRibbonHost" -Recurse -Force -ErrorAction SilentlyContinue

# Best-effort machine unregister if previously installed elevated
if ((Test-Path $dll) -and (Test-Path $regasm)) {
    & $regasm $dll /unregister /silent 2>$null | Out-Null
}

Write-Host "Unregistered Aliniant.AesRibbonHost. Restart Outlook."
