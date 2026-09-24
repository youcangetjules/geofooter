# Build and register Aliniant AES Ribbon Host for classic Outlook (COM add-in).
# 32-bit Outlook requires COM keys written via 32-bit reg.exe / registry view.
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$clsid = "{A1E50001-AE51-4B0B-9C11-AE5B1BB00001}"
$progId = "Aliniant.AesRibbonHost"
$dll = Join-Path $root "bin\Release\AesRibbonHost.dll"

Write-Host "Building AesRibbonHost (x86)..."
dotnet build "$root\AesRibbonHost.csproj" -c Release
if (-not (Test-Path $dll)) { throw "Build failed - DLL not found: $dll" }

$regasm = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
if (-not (Test-Path $regasm)) { throw "32-bit RegAsm.exe not found: $regasm" }
Write-Host "Using RegAsm: $regasm"

$regfile = Join-Path $root "bin\Release\AesRibbonHost.reg"
$hkcuReg = Join-Path $root "bin\Release\AesRibbonHost-HKCU.reg"

Write-Host "Generating registration script..."
& $regasm $dll /codebase /regfile:"$regfile"
if (-not (Test-Path $regfile)) { throw "RegAsm /regfile failed" }

$raw = Get-Content $regfile -Raw
$raw = $raw -replace 'HKEY_CLASSES_ROOT', 'HKEY_CURRENT_USER\Software\Classes'
# Also register under Wow6432Node\Classes so 32-bit Outlook always finds the CLSID
# when 32-bit reg.exe redirected the primary write (or 64-bit hosts look there).
$wow = $raw -replace '\\Software\\Classes\\CLSID\\', '\Software\Classes\Wow6432Node\CLSID\'
$wow = $wow -replace '\\Software\\Classes\\Aliniant\.AesRibbonHost', '\Software\Classes\Wow6432Node\Aliniant.AesRibbonHost'
$wow = $wow -replace '\\Software\\Classes\\Record\\', '\Software\Classes\Wow6432Node\Record\'
# Keep ProgId in both places: merge primary + Wow lines (avoid dropping Class keys).
$merged = $raw.TrimEnd() + "`r`n`r`n" + ($wow -replace 'REGEDIT4\s*', '')
Set-Content -Path $hkcuReg -Value $merged -Encoding ASCII

$reg32 = Join-Path $env:WINDIR "SysWOW64\reg.exe"
if (-not (Test-Path $reg32)) { $reg32 = Join-Path $env:WINDIR "System32\reg.exe" }

Write-Host "Importing COM registration with: $reg32"
& $reg32 import $hkcuReg
if ($LASTEXITCODE -ne 0) { throw "reg import failed: $LASTEXITCODE" }

# Locate CLSID in whichever view 32-bit reg.exe actually wrote.
$clsidCandidates = @(
    "HKCU:\Software\Classes\CLSID\$clsid",
    "HKCU:\Software\Classes\Wow6432Node\CLSID\$clsid",
    "HKCU:\Software\Wow6432Node\Classes\CLSID\$clsid"
)
$found = $null
foreach ($p in $clsidCandidates) {
    if (Test-Path $p) {
        $found = $p
        Write-Host "CLSID present at: $p"
        break
    }
}

if (-not $found) {
    Write-Host "CLSID missing after import - writing keys explicitly (32-bit view)..."
    $codeBase = "file:///" + ($dll -replace '\\', '/')
    $ps32 = Join-Path $env:WINDIR "SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    $writeScript = Join-Path $root "bin\Release\write_clsid32.ps1"
    @"
`$ErrorActionPreference = 'Stop'
`$clsid = '$clsid'
`$progId = '$progId'
`$codeBase = '$codeBase'
`$root = "HKCU:\Software\Classes\CLSID\`$clsid"
New-Item -Path `$root -Force | Out-Null
Set-ItemProperty -Path `$root -Name '(default)' -Value 'Aliniant.AesRibbonHost.Connect'
`$inproc = Join-Path `$root 'InprocServer32'
New-Item -Path `$inproc -Force | Out-Null
Set-ItemProperty -Path `$inproc -Name '(default)' -Value 'mscoree.dll'
Set-ItemProperty -Path `$inproc -Name 'ThreadingModel' -Value 'Both'
Set-ItemProperty -Path `$inproc -Name 'Class' -Value 'Aliniant.AesRibbonHost.Connect'
Set-ItemProperty -Path `$inproc -Name 'Assembly' -Value 'AesRibbonHost, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'
Set-ItemProperty -Path `$inproc -Name 'RuntimeVersion' -Value 'v4.0.30319'
Set-ItemProperty -Path `$inproc -Name 'CodeBase' -Value `$codeBase
`$ver = Join-Path `$inproc '1.0.0.0'
New-Item -Path `$ver -Force | Out-Null
Set-ItemProperty -Path `$ver -Name 'Class' -Value 'Aliniant.AesRibbonHost.Connect'
Set-ItemProperty -Path `$ver -Name 'Assembly' -Value 'AesRibbonHost, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'
Set-ItemProperty -Path `$ver -Name 'RuntimeVersion' -Value 'v4.0.30319'
Set-ItemProperty -Path `$ver -Name 'CodeBase' -Value `$codeBase
`$prog = "HKCU:\Software\Classes\`$progId"
New-Item -Path `$prog -Force | Out-Null
Set-ItemProperty -Path `$prog -Name '(default)' -Value 'Aliniant.AesRibbonHost.Connect'
New-Item -Path (Join-Path `$prog 'CLSID') -Force | Out-Null
Set-ItemProperty -Path (Join-Path `$prog 'CLSID') -Name '(default)' -Value `$clsid
Write-Host 'Explicit CLSID write OK'
"@ | Set-Content -Path $writeScript -Encoding ASCII
    if (Test-Path $ps32) {
        & $ps32 -NoProfile -ExecutionPolicy Bypass -File $writeScript
    } else {
        & powershell -NoProfile -ExecutionPolicy Bypass -File $writeScript
    }
    foreach ($p in $clsidCandidates) {
        if (Test-Path $p) { $found = $p; Write-Host "CLSID present at: $p"; break }
    }
}

# Soft mirror into Wow6432Node when the primary Classes\CLSID view exists (never abort install).
try {
    $src = "HKCU:\Software\Classes\CLSID\$clsid"
    $dstRoot = "HKCU:\Software\Classes\Wow6432Node\CLSID"
    $dst = "HKCU:\Software\Classes\Wow6432Node\CLSID\$clsid"
    if ((Test-Path $src) -and -not (Test-Path $dst)) {
        Write-Host "Mirroring CLSID into Wow6432Node..."
        New-Item -Path $dstRoot -Force | Out-Null
        Copy-Item $src $dst -Recurse -Force -ErrorAction Stop
    }
} catch {
    Write-Host ("Wow6432Node mirror skipped: " + $_.Exception.Message)
}

$keyPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$progId"
Write-Host "Writing Outlook Add-in key: $keyPath"
New-Item -Path $keyPath -Force | Out-Null
Set-ItemProperty -Path $keyPath -Name "FriendlyName" -Value "Aliniant AES Ribbon"
Set-ItemProperty -Path $keyPath -Name "Description" -Value "Hosts AES Ribbon XML on Home (large icons, live ON/OFF)"
New-ItemProperty -Path $keyPath -Name "LoadBehavior" -PropertyType DWord -Value 3 -Force | Out-Null

$ps32 = Join-Path $env:WINDIR "SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
$verify = Join-Path $root "verify32.ps1"
@'
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
'@ | Set-Content $verify -Encoding ASCII

if (Test-Path $ps32) {
    Write-Host "Verifying 32-bit CLSID visibility..."
    & $ps32 -NoProfile -ExecutionPolicy Bypass -File $verify
}

$logDir = Join-Path $env:LOCALAPPDATA "GeoFooter\Logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
Write-Host ""
if ($found) {
    Write-Host "OK. Fully quit Outlook (tray too), then reopen."
} else {
    Write-Host "WARN: CLSID still not visible - check AesRibbonHost.log after Outlook restart."
}
Write-Host ("Add-in log: " + (Join-Path $logDir "AesRibbonHost.log"))
