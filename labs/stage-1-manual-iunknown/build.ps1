# Builds the Stage 1 console app. Run from a Developer PowerShell for VS (x64).
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

if (-not (Get-Command cl.exe -ErrorAction SilentlyContinue)) {
    Write-Error "cl.exe not found. Open a 'Developer PowerShell for VS' (x64) and try again."
}

# /Zi + /DEBUG so the debugger and WinDbg have symbols - several labs need them.
cl /nologo /W4 /EHsc /std:c++17 /Zi `
   main.cpp Calculator.cpp `
   /Fe:Lab01.exe `
   /link /DEBUG ole32.lib

if ($LASTEXITCODE -ne 0) { Write-Error "Build failed." }
Write-Host "`nBuilt Lab01.exe - run it with .\Lab01.exe" -ForegroundColor Green
