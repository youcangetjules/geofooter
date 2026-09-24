$clsid = "{A1E50001-AE51-4B0B-9C11-AE5B1BB00001}"
$paths = @(
    "HKCU:\Software\Classes\CLSID\$clsid\InprocServer32",
    "HKCU:\Software\Classes\Wow6432Node\CLSID\$clsid\InprocServer32"
)
$ok = $false
foreach ($path in $paths) {
    Write-Host ("32-bit sees InprocServer32 ($path): " + (Test-Path $path))
    if (Test-Path $path) { $ok = $true }
}
if ($ok) {
    try {
        $null = New-Object -ComObject Aliniant.AesRibbonHost
        Write-Host "32-bit CoCreate: OK"
    } catch {
        Write-Host ("32-bit CoCreate FAIL: " + $_.Exception.Message)
    }
} else {
    Write-Host "32-bit CoCreate SKIP: CLSID InprocServer32 not found"
}
