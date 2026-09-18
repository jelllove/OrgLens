[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$SourcePath,
    [ValidateSet('Auto', 'x64', 'x86')][string]$OutlookBitness = 'Auto'
)
$ErrorActionPreference = 'Stop'
$id = '{D7E2D48A-9466-4D83-856F-AC197BC23A98}'
$progId = 'OrgLens.Connect'
if (-not $SourcePath) {
    $SourcePath = if (Test-Path (Join-Path $PSScriptRoot 'OrgLens.Outlook.dll')) { $PSScriptRoot }
        else {
            $root = Split-Path $PSScriptRoot -Parent
            $version = ([xml](Get-Content (Join-Path $root 'Directory.Build.props') -Raw)).Project.PropertyGroup.Version
            Join-Path $root "artifacts\OrgLens-$version"
        }
}
$SourcePath = (Resolve-Path $SourcePath).Path
$requiredFiles = @('OrgLens.Outlook.dll', 'OrgLens.Core.dll', 'OrgLens.Desktop.dll', 'Newtonsoft.Json.dll', 'Newtonsoft.Json.LICENSE.md',
    'OrgLens.Preview.exe', 'OrgLens.Preview.exe.config', 'Install.ps1', 'Uninstall.ps1', 'README.md',
    'LICENSE', 'CHANGELOG.md', 'OrgLens.Icon.txt', 'docs\images\orglens-settings.png', 'docs\images\orglens-icon.png')
foreach ($file in $requiredFiles) {
    if (-not (Test-Path (Join-Path $SourcePath $file))) { throw "Incomplete package: $file is missing. Run scripts\Build.ps1 first." }
}
if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) {
    if (-not $WhatIfPreference) {
        throw 'Close classic Outlook before installing. OrgLens will never terminate Outlook automatically.'
    }
    Write-Warning 'Outlook is running. A real installation requires closing it first.'
}
if ($OutlookBitness -eq 'Auto') {
    $paths = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\OUTLOOK.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\OUTLOOK.EXE",
        "$env:ProgramFiles\Microsoft Office\Office16\OUTLOOK.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OUTLOOK.EXE"
    )
    $outlook = $paths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $outlook) { throw 'Classic Outlook was not detected. Install classic Outlook or supply -OutlookBitness x64/x86 for a nonstandard installation.' }
    $stream = [IO.File]::OpenRead($outlook)
    try {
        $reader = [IO.BinaryReader]::new($stream)
        $stream.Position = 0x3c
        $offset = $reader.ReadInt32()
        $stream.Position = $offset + 4
        $machine = $reader.ReadUInt16()
        switch ($machine) {
            0x8664 { $OutlookBitness = 'x64' }
            0x014c { $OutlookBitness = 'x86' }
            default { throw "Unsupported Outlook executable architecture: $machine" }
        }
    }
    finally { $stream.Dispose() }
}
$installPath = Join-Path $env:LOCALAPPDATA 'OrgLens\Addin'
$assemblyName = [Reflection.AssemblyName]::GetAssemblyName((Join-Path $SourcePath 'OrgLens.Outlook.dll'))
$codeBase = ([Uri](Join-Path $installPath 'OrgLens.Outlook.dll')).AbsoluteUri
$registryView = if ($OutlookBitness -eq 'x64') { [Microsoft.Win32.RegistryView]::Registry64 }
    else { [Microsoft.Win32.RegistryView]::Registry32 }

if (-not $PSCmdlet.ShouldProcess("Current user, Outlook $OutlookBitness, $installPath", 'Install and register OrgLens')) { return }
$null = New-Item $installPath -ItemType Directory -Force
foreach ($file in $requiredFiles) {
    $source = Join-Path $SourcePath $file
    $destination = Join-Path $installPath $file
    if (-not [string]::Equals($source, $destination, [StringComparison]::OrdinalIgnoreCase)) {
        $null = New-Item (Split-Path $destination -Parent) -ItemType Directory -Force
        Copy-Item $source $destination -Force
    }
}
$base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $registryView)
function Set-RegistryValues([string]$Path, [hashtable]$Values) {
    $key = $base.CreateSubKey($Path)
    try {
        foreach ($entry in $Values.GetEnumerator()) {
            $kind = if ($entry.Value -is [int]) { [Microsoft.Win32.RegistryValueKind]::DWord }
                else { [Microsoft.Win32.RegistryValueKind]::String }
            $key.SetValue($entry.Key, $entry.Value, $kind)
        }
    }
    finally { $key.Dispose() }
}
try {
    Set-RegistryValues "Software\Classes\$progId" @{ '' = 'OrgLens classic Outlook add-in' }
    Set-RegistryValues "Software\Classes\$progId\CLSID" @{ '' = $id }
    Set-RegistryValues "Software\Classes\CLSID\$id" @{ '' = 'OrgLens.Outlook.Connect' }
    Set-RegistryValues "Software\Classes\CLSID\$id\ProgId" @{ '' = $progId }
    $managed = @{
        'Class' = 'OrgLens.Outlook.Connect'
        'Assembly' = $assemblyName.FullName
        'RuntimeVersion' = 'v4.0.30319'
        'CodeBase' = $codeBase
    }
    Set-RegistryValues "Software\Classes\CLSID\$id\InprocServer32" ($managed + @{
        '' = 'mscoree.dll'; 'ThreadingModel' = 'Both'
    })
    Set-RegistryValues "Software\Classes\CLSID\$id\InprocServer32\$($assemblyName.Version)" $managed
    Set-RegistryValues "Software\Classes\CLSID\$id\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}" @{}
    Set-RegistryValues "Software\Microsoft\Office\Outlook\Addins\$progId" @{
        FriendlyName = 'OrgLens'
        Description = 'Highlight emails from your management chain.'
        LoadBehavior = 3
        CommandLineSafe = 0
    }
}
finally { $base.Dispose() }
Write-Host "OrgLens installed for the current user ($OutlookBitness)."
Write-Host 'Open classic Outlook, select the OrgLens tab, then Formatting rules.'
Write-Host 'No mailbox rules have been changed. Apply them explicitly in the settings window.'
Write-Host "Uninstall with: powershell.exe -NoProfile -File `"$installPath\Uninstall.ps1`""
