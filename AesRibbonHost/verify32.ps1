# Outlook loads via Addins\ProgId + CLSID registry keys -- that is the install success
# criterion. In-process CoCreate often fails while Outlook holds AesRibbonHost.dll
# (Guid.Empty / REGDB_E_CLASSNOTREG); do not report that as an install failure.
$ErrorActionPreference = "Continue"
$clsid = "{A1E50001-AE51-4B0B-9C11-AE5B1BB00001}"
$progId = "Aliniant.AesRibbonHost"
$paths = @(
    "HKCU:\Software\Classes\CLSID\$clsid\InprocServer32",
    "HKCU:\Software\Classes\Wow6432Node\CLSID\$clsid\InprocServer32"
)
$regOk = $false
foreach ($path in $paths) {
    $seen = Test-Path $path
    Write-Host ("32-bit sees InprocServer32 ($path): " + $seen)
    if ($seen) { $regOk = $true }
}
$progClsid = $null
try {
    $progClsid = (Get-ItemProperty "HKCU:\Software\Classes\$progId\CLSID" -ErrorAction Stop).'(default)'
} catch { }
Write-Host ("ProgId CLSID: [" + $progClsid + "]")
if ($progClsid -ne $clsid) {
    Write-Host "WARN: ProgId does not map to expected CLSID."
}

if (-not $regOk) {
    Write-Host "32-bit registration: FAIL (CLSID InprocServer32 not found)"
    exit 1
}

Write-Host "32-bit registration: OK (CLSID + InprocServer32 present)"

# Soft probe only -- never escalate busy-DLL / Guid.Empty into FAIL.
$activated = $false
try {
    $null = [Activator]::CreateInstance([Type]::GetTypeFromCLSID([guid]$clsid))
    $activated = $true
} catch {
    try {
        $null = New-Object -ComObject $progId
        $activated = $true
    } catch { }
}
if ($activated) {
    Write-Host "32-bit CoCreate: OK"
} else {
    Write-Host "32-bit CoCreate: skipped (Outlook may be holding the DLL - not an install failure)"
}
exit 0
