[CmdletBinding()]
param([string]$InnoCompilerPath)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$version = ([xml](Get-Content (Join-Path $root 'Directory.Build.props') -Raw)).Project.PropertyGroup.Version
$package = Join-Path $root "artifacts\OrgLens-$version"
$fixtureExe = Join-Path $root 'tests\OrgLens.Tests\bin\Release\net48\OrgLens.Tests.exe'
if (-not (Test-Path $fixtureExe)) { throw 'Build OrgLens before running installer tests.' }
$identity = [Reflection.AssemblyName]::GetAssemblyName((Join-Path $package 'OrgLens.Outlook.dll'))
$clsid = '{D7E2D48A-9466-4D83-856F-AC197BC23A98}'
$script:checks = 0

function Assert([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
    $script:checks++
    Write-Host "PASS $Message"
}

function Run-Setup([string]$Executable, [string[]]$Options, [string]$LogPath) {
    $arguments = @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/SP-', '/NORESTART', "/LOG=`"$LogPath`"") + $Options
    # -Wait includes the uninstaller's self-deleting child, not just its launcher.
    $process = Start-Process -FilePath $Executable -ArgumentList $arguments -PassThru -Wait
    return $process.ExitCode
}

function Test-Architecture([string]$Architecture, [string]$OutlookPath, [Microsoft.Win32.RegistryView]$View) {
    $id = [Guid]::NewGuid().ToString('D')
    $temp = Join-Path "$env:LOCALAPPDATA\Temp" "OrgLens-SetupTests-$id"
    $installDir = Join-Path $temp ("OrgLens # % " + [char]0x6D4B + [char]0x8BD5 + '\Addin')
    $prefix = "Software\OrgLens\SetupTests\$id"
    $classKey = "$prefix\Software\Classes\CLSID\$clsid\InprocServer32"
    $addinKey = "$prefix\Software\Microsoft\Office\Outlook\Addins\OrgLens.Connect"
    $arpKey = "Software\Microsoft\Windows\CurrentVersion\Uninstall\OrgLens.SetupTest.${id}_is1"
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $View)
    $setup = $null
    $fixture = $null
    $succeeded = $false
    $uninstalled = $false
    $null = New-Item $temp -ItemType Directory
    try {
        $snapshot = Join-Path $temp 'Package'
        Copy-Item -LiteralPath $package -Destination $snapshot -Recurse
        $setup = & (Join-Path $PSScriptRoot 'Build-Installer.ps1') -InnoCompilerPath $InnoCompilerPath -PackagePath $snapshot `
            -TestRunId $id -TestOutlookPath $OutlookPath -TestProcessName 'OrgLens.Tests.exe'
        Assert ($setup.Path -like "*-Test-$id.exe") "$Architecture test binary is visibly isolated"

        $fixture = Start-Process $fixtureExe -ArgumentList '--installer-process-fixture' -WindowStyle Hidden -PassThru
        Start-Sleep -Milliseconds 400
        Assert (-not $fixture.HasExited) 'The harmless process-detection fixture is alive'
        $exit = Run-Setup $setup.Path @("/DIR=`"$installDir`"") (Join-Path $temp 'blocked-process.log')
        Assert ($exit -ne 0 -and -not (Test-Path (Join-Path $installDir 'OrgLens.Outlook.dll'))) `
            "$Architecture setup refuses to install while the detected process is running"
        Stop-Process -Id $fixture.Id
        $fixture.WaitForExit()
        $fixture = $null

        $key = $base.CreateSubKey($classKey)
        try { $key.SetValue('Class', 'Foreign.Component') } finally { $key.Dispose() }
        $exit = Run-Setup $setup.Path @("/DIR=`"$installDir`"") (Join-Path $temp 'blocked-owner.log')
        $key = $base.OpenSubKey($classKey)
        try {
            Assert ($exit -ne 0 -and $key.GetValue('Class') -eq 'Foreign.Component') `
                "$Architecture setup preserves a foreign COM registration"
        } finally { $key.Dispose() }
        $base.DeleteSubKeyTree($prefix)

        $null = New-Item $installDir -ItemType Directory -Force
        $personalFile = Join-Path $installDir 'personal.orglens.json'
        [IO.File]::WriteAllText($personalFile, 'personal settings fixture')
        $profileDir = Join-Path (Split-Path $installDir -Parent) 'Profiles'
        $null = New-Item $profileDir -ItemType Directory
        $profileFile = Join-Path $profileDir 'fixture.json'
        [IO.File]::WriteAllText($profileFile, 'account cache fixture')
        [IO.File]::WriteAllText((Join-Path $installDir 'OrgLens.Outlook.dll'), 'legacy binary fixture')
        $seed = @{
            "$prefix\Software\Classes\CLSID\$clsid" = @{ '' = 'OrgLens.Outlook.Connect' }
            "$prefix\Software\Classes\CLSID\$clsid\ProgId" = @{ '' = 'OrgLens.Connect' }
            "$prefix\Software\Classes\OrgLens.Connect" = @{ '' = 'OrgLens classic Outlook add-in' }
            "$prefix\Software\Classes\OrgLens.Connect\CLSID" = @{ '' = $clsid }
            $classKey = @{
                '' = 'mscoree.dll'; 'ThreadingModel' = 'Both'; 'Class' = 'OrgLens.Outlook.Connect'
                'Assembly' = 'OrgLens.Outlook, Version=0.2.2.0, Culture=neutral, PublicKeyToken=null'
                'RuntimeVersion' = 'v4.0.30319'
                'CodeBase' = ([Uri](Join-Path $installDir 'OrgLens.Outlook.dll')).AbsoluteUri
            }
        }
        $seed["$classKey\0.2.2.0"] = $seed[$classKey]
        foreach ($path in $seed.Keys) {
            $key = $base.CreateSubKey($path)
            try {
                foreach ($name in $seed[$path].Keys) { $key.SetValue($name, $seed[$path][$name]) }
            } finally { $key.Dispose() }
        }

        $exit = Run-Setup $setup.Path @("/DIR=`"$installDir`"") (Join-Path $temp 'install.log')
        Assert ($exit -eq 0) "$Architecture legacy-install upgrade succeeds without elevation"
        Assert ((Get-Content (Join-Path $temp 'install.log') -Raw).Contains("Detected classic Outlook ${Architecture}:")) `
            "$Architecture PE detection is not redirected to a different executable"
        $key = $base.OpenSubKey($classKey)
        Assert ($null -ne $key) "$Architecture managed COM registration exists in the expected registry view"
        try {
            Assert ($key.GetValue('Assembly') -eq $identity.FullName) "$Architecture COM assembly identity matches the installed DLL"
            $uri = [Uri]$key.GetValue('CodeBase')
            Assert ($uri.IsFile -and [string]::Equals($uri.LocalPath, (Join-Path $installDir 'OrgLens.Outlook.dll'),
                [StringComparison]::OrdinalIgnoreCase)) "$Architecture CodeBase preserves spaces, Unicode, percent and hash characters"
        } finally { $key.Dispose() }
        $key = $base.OpenSubKey("$classKey\$($identity.Version)")
        Assert ($null -ne $key) "$Architecture current assembly-version registration exists"
        $key.Dispose()
        $key = $base.OpenSubKey($addinKey)
        Assert ($null -ne $key) "$Architecture Outlook add-in registration exists"
        try { Assert ($key.GetValue('LoadBehavior') -eq 3) "$Architecture add-in uses normal load behavior" }
        finally { $key.Dispose() }
        $key = $base.OpenSubKey($arpKey)
        Assert ($null -ne $key) "$Architecture Windows Installed apps entry exists"
        try { Assert ($key.GetValue('DisplayVersion') -eq $version) "$Architecture Installed apps reports the current version" }
        finally { $key.Dispose() }
        foreach ($file in @('OrgLens.Outlook.dll','OrgLens.Core.dll','OrgLens.Desktop.dll','Newtonsoft.Json.dll',
            'OrgLens.Preview.exe','LICENSE','Newtonsoft.Json.LICENSE.md','README.md','docs\images\orglens-settings.png')) {
            Assert ((Get-FileHash (Join-Path $installDir $file)).Hash -eq (Get-FileHash (Join-Path $snapshot $file)).Hash) `
                "$Architecture installed payload matches: $file"
        }

        $exit = Run-Setup $setup.Path @("/DIR=`"$installDir`"") (Join-Path $temp 'reinstall.log')
        Assert ($exit -eq 0) "$Architecture reinstall succeeds"
        Assert ((Get-Content $personalFile -Raw) -eq 'personal settings fixture' `
            -and (Get-Content $profileFile -Raw) -eq 'account cache fixture') "$Architecture upgrade preserves personal settings"

        $uninstaller = Join-Path $installDir 'unins000.exe'
        Assert (Test-Path $uninstaller) "$Architecture native uninstaller exists"
        $fixture = Start-Process $fixtureExe -ArgumentList '--installer-process-fixture' -WindowStyle Hidden -PassThru
        Start-Sleep -Milliseconds 400
        $exit = Run-Setup $uninstaller @() (Join-Path $temp 'blocked-uninstall.log')
        Assert ($exit -ne 0 -and (Test-Path (Join-Path $installDir 'OrgLens.Outlook.dll'))) `
            "$Architecture uninstall refuses to modify a running add-in"
        Stop-Process -Id $fixture.Id
        $fixture.WaitForExit()
        $fixture = $null
        $exit = Run-Setup $uninstaller @() (Join-Path $temp 'uninstall.log')
        Assert ($exit -eq 0) "$Architecture native uninstall succeeds"
        $uninstalled = $true
        for ($attempt = 0; $attempt -lt 30 -and (Test-Path (Join-Path $installDir 'OrgLens.Outlook.dll')); $attempt++) {
            Start-Sleep -Milliseconds 100
        }
        Assert (-not (Test-Path (Join-Path $installDir 'OrgLens.Outlook.dll'))) "$Architecture uninstall removes installed binaries"
        foreach ($path in @($classKey, $addinKey, $arpKey)) {
            $key = $base.OpenSubKey($path)
            $exists = $null -ne $key
            if ($key) { $key.Dispose() }
            Assert (-not $exists) "$Architecture uninstall removes its registration: $path"
        }
        Assert ((Get-Content $personalFile -Raw) -eq 'personal settings fixture' `
            -and (Get-Content $profileFile -Raw) -eq 'account cache fixture') "$Architecture uninstall preserves personal settings"
        $succeeded = $true
    } finally {
        if ($fixture -and -not $fixture.HasExited) { Stop-Process -Id $fixture.Id }
        $uninstaller = Join-Path $installDir 'unins000.exe'
        if (-not $uninstalled -and (Test-Path $uninstaller)) {
            $cleanupExit = Run-Setup $uninstaller @() (Join-Path $temp 'cleanup.log')
            if ($cleanupExit -ne 0) { Write-Warning "Test uninstaller cleanup failed; logs retained at $temp" }
        }
        for ($attempt = 0; $attempt -lt 50 -and (Test-Path $uninstaller); $attempt++) {
            Start-Sleep -Milliseconds 100
        }
        $base.DeleteSubKeyTree($prefix, $false)
        $base.DeleteSubKeyTree($arpKey, $false)
        $base.Dispose()
        if ($setup) {
            Remove-Item -LiteralPath $setup.Path, $setup.ChecksumPath -Force
        }
        if ($succeeded -and -not (Test-Path $uninstaller)) {
            Remove-Item -LiteralPath $temp -Recurse -Force
            $rollback = Join-Path $env:LOCALAPPDATA "OrgLens\SetupTests\$id"
            if (Test-Path $rollback) { Remove-Item -LiteralPath $rollback -Recurse -Force }
        } else {
            Write-Warning "Installer test logs retained at $temp"
        }
    }
}

if ([Environment]::Is64BitOperatingSystem) {
    Test-Architecture 'x64' "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" ([Microsoft.Win32.RegistryView]::Registry64)
    Test-Architecture 'x86' "$env:WINDIR\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" ([Microsoft.Win32.RegistryView]::Registry32)
} else {
    Test-Architecture 'x86' "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" ([Microsoft.Win32.RegistryView]::Registry32)
}
Write-Host "PASS: $script:checks isolated installer checks. Real Outlook registration was not changed."
