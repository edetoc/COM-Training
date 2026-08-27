# Builds Calc.dll and CalcClient.exe. Run from a Developer PowerShell for VS.
#
#   .\build.ps1            -> x64 (default), output in .\x64
#   .\build.ps1 -Arch x86  -> 32-bit, output in .\x86   (needed for Lab 2.2)
param(
    [ValidateSet('x64', 'x86')]
    [string]$Arch = 'x64'
)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

if (-not (Get-Command cl.exe -ErrorAction SilentlyContinue)) {
    Write-Error "cl.exe not found. Open a 'Developer PowerShell for VS' and try again."
}

# Confirm the compiler matches the requested architecture. Building x86 needs an
# x86 developer prompt (or `Enter-VsDevShell -Arch x86`); a mismatch here is the
# single most common reason Lab 2.2 behaves confusingly.
$clTarget = (& cl.exe 2>&1 | Select-String -Pattern 'for (x86|x64)' | Select-Object -First 1).Matches.Groups[1].Value
if ($clTarget -ne $Arch) {
    Write-Error "This shell's cl.exe targets $clTarget but -Arch is $Arch. Open the matching Developer PowerShell."
}

$out = Join-Path $PSScriptRoot $Arch
New-Item -ItemType Directory -Force -Path $out | Out-Null
Push-Location $out
try {
    cl /nologo /W4 /EHsc /std:c++17 /Zi /LD `
       ..\Calc.cpp `
       /Fe:Calc.dll `
       /link /DEBUG /DEF:..\Calc.def ole32.lib advapi32.lib
    if ($LASTEXITCODE -ne 0) { Write-Error "Server build failed." }

    cl /nologo /W4 /EHsc /std:c++17 /Zi `
       ..\CalcClient.cpp `
       /Fe:CalcClient.exe `
       /link /DEBUG ole32.lib
    if ($LASTEXITCODE -ne 0) { Write-Error "Client build failed." }
}
finally { Pop-Location }

Write-Host "`nBuilt $Arch\Calc.dll and $Arch\CalcClient.exe" -ForegroundColor Green
Write-Host "Next: register the DLL from an ELEVATED prompt:" -ForegroundColor Yellow
Write-Host "    regsvr32 `"$out\Calc.dll`""
