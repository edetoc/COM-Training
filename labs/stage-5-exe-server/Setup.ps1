[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateSet('Register', 'Unregister')]
    [string] $Action = 'Register',
    [ValidateSet('x64', 'x86')]
    [string] $Architecture = 'x64'
)

$ErrorActionPreference = 'Stop'
if (-not [Environment]::Is64BitOperatingSystem -or -not [Environment]::Is64BitProcess) {
    throw 'Run this script from 64-bit PowerShell on 64-bit Windows.'
}

$output = Join-Path $PSScriptRoot $Architecture
$regsvr32 = if ($Architecture -eq 'x64') {
    Join-Path $env:WINDIR 'System32\regsvr32.exe'
} else {
    Join-Path $env:WINDIR 'SysWOW64\regsvr32.exe'
}
$proxy = Join-Path $output 'Stage5PS.dll'
$dll = Join-Path $output 'CalcDll.dll'
$server = Join-Path $output 'CalcSrv.exe'
$required = @($proxy, $dll)
if ($Architecture -eq 'x64') { $required += $server }
foreach ($file in $required) {
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
        throw "Missing $file. Build Stage5.sln for $Architecture first."
    }
}

if (-not $PSCmdlet.ShouldProcess("Stage 5 $Architecture training components", $Action)) { return }
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
try {
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'Registration requires an elevated PowerShell window. The script does not elevate itself.'
    }
} finally {
    $identity.Dispose()
}

function Invoke-RegistrationProcess([string] $Executable, [string] $Arguments) {
    $process = Start-Process -FilePath $Executable -ArgumentList $Arguments -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        throw "$Executable failed (exit $($process.ExitCode)). Stop clients and check elevation and binary paths before retrying."
    }
}

if ($Action -eq 'Register') {
    Invoke-RegistrationProcess $regsvr32 "/s `"$proxy`""
    Invoke-RegistrationProcess $regsvr32 "/s `"$dll`""
    if ($Architecture -eq 'x64') { Invoke-RegistrationProcess $server '-RegServer' }
} else {
    if ($Architecture -eq 'x64') { Invoke-RegistrationProcess $server '-UnregServer' }
    Invoke-RegistrationProcess $regsvr32 "/s /u `"$dll`""
    Invoke-RegistrationProcess $regsvr32 "/s /u `"$proxy`""
}
Write-Output "$Action completed for Stage 5 $Architecture."