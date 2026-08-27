# Builds CalcSrv.exe (the out-of-proc server) and CalcSrvClient.exe.
# Run from a Developer PowerShell for VS (x64).
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

if (-not (Get-Command cl.exe -ErrorAction SilentlyContinue)) {
    Write-Error "cl.exe not found. Open a 'Developer PowerShell for VS' and try again."
}

$out = Join-Path $PSScriptRoot 'x64'
New-Item -ItemType Directory -Force -Path $out | Out-Null
Push-Location $out
try {
    # /SUBSYSTEM:WINDOWS because the server has a message loop and no console.
    cl /nologo /W4 /EHsc /std:c++17 /Zi `
       ..\CalcSrv.cpp `
       /Fe:CalcSrv.exe `
       /link /DEBUG /SUBSYSTEM:WINDOWS ole32.lib advapi32.lib user32.lib
    if ($LASTEXITCODE -ne 0) { Write-Error "Server build failed." }

    cl /nologo /W4 /EHsc /std:c++17 /Zi `
       ..\CalcSrvClient.cpp `
       /Fe:CalcSrvClient.exe `
       /link /DEBUG ole32.lib
    if ($LASTEXITCODE -ne 0) { Write-Error "Client build failed." }
}
finally { Pop-Location }

Write-Host "`nBuilt x64\CalcSrv.exe and x64\CalcSrvClient.exe" -ForegroundColor Green
Write-Host "Next, from an ELEVATED prompt:" -ForegroundColor Yellow
Write-Host "    & `"$out\CalcSrv.exe`" -RegServer"
Write-Host "And make sure Stage 3's CalcPS.dll is registered, or activation fails with E_NOINTERFACE."
