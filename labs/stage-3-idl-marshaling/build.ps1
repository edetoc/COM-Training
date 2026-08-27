# Runs MIDL on Calculator.idl, then builds the proxy/stub DLL.
# Run from a Developer PowerShell for VS.
#
#   .\build.ps1            -> x64 (default)
#   .\build.ps1 -Arch x86  -> 32-bit (Lab 4.2 needs both)
param(
    [ValidateSet('x64', 'x86')]
    [string]$Arch = 'x64'
)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

foreach ($tool in 'midl.exe', 'cl.exe', 'link.exe') {
    if (-not (Get-Command $tool -ErrorAction SilentlyContinue)) {
        Write-Error "$tool not found. Open a 'Developer PowerShell for VS' and try again."
    }
}

$out = Join-Path $PSScriptRoot $Arch
New-Item -ItemType Directory -Force -Path $out | Out-Null

# --- Step 1: MIDL ---------------------------------------------------------
# Generates: Calculator.h      the C/C++ interface declarations
#            Calculator_i.c    the IID/CLSID definitions
#            Calculator_p.c    the proxy/stub code (NDR format strings)
#            dlldata.c         the proxy DLL's entry points
#            Calculator.tlb    the type library
Write-Host "MIDL..." -ForegroundColor Cyan
midl /nologo /W1 /char signed /env $Arch `
     /out $out `
     /h Calculator.h `
     /iid Calculator_i.c `
     /proxy Calculator_p.c `
     /dlldata dlldata.c `
     /tlb Calculator.tlb `
     Calculator.idl
if ($LASTEXITCODE -ne 0) { Write-Error "MIDL failed." }

# --- Step 2: compile and link the proxy/stub DLL --------------------------
# REGISTER_PROXY_DLL is what makes dlldata.c emit DllRegisterServer and friends.
Write-Host "Building CalcPS.dll..." -ForegroundColor Cyan
Push-Location $out
try {
    cl /nologo /c /W3 /Zi /DWIN32 /D_WIN32_WINNT=0x0A00 /DREGISTER_PROXY_DLL `
       dlldata.c Calculator_p.c Calculator_i.c
    if ($LASTEXITCODE -ne 0) { Write-Error "Compile failed." }

    link /nologo /DLL /DEBUG /DEF:..\CalcPS.def /OUT:CalcPS.dll `
         dlldata.obj Calculator_p.obj Calculator_i.obj `
         rpcrt4.lib ole32.lib oleaut32.lib kernel32.lib advapi32.lib
    if ($LASTEXITCODE -ne 0) { Write-Error "Link failed." }
}
finally { Pop-Location }

Write-Host "`nBuilt $Arch\CalcPS.dll" -ForegroundColor Green
Write-Host "Next, from an ELEVATED prompt:" -ForegroundColor Yellow
Write-Host "    regsvr32 `"$out\CalcPS.dll`""
Write-Host "Read the generated files in $Arch\ - that is the point of Lab 4.1."
