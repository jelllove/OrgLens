[CmdletBinding()]
param(
    [string]$PackagePath,
    [string]$InnoCompilerPath,
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')][string]$TestRunId,
    [string]$TestOutlookPath,
    [ValidatePattern('^[A-Za-z0-9_.-]+\.exe$')][string]$TestProcessName
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (($TestOutlookPath -or $TestProcessName) -and -not $TestRunId) {
    throw 'Test overrides require a compile-time isolated TestRunId.'
}
if (-not $PackagePath) {
    $version = ([xml](Get-Content (Join-Path $root 'Directory.Build.props') -Raw)).Project.PropertyGroup.Version
    $PackagePath = Join-Path $root "artifacts\OrgLens-$version"
}
$PackagePath = (Resolve-Path -LiteralPath $PackagePath).Path
$assembly = [Reflection.AssemblyName]::GetAssemblyName((Join-Path $PackagePath 'OrgLens.Outlook.dll'))
$version = $assembly.Version.ToString(3)

if (-not $InnoCompilerPath) {
    $command = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($command) {
        $InnoCompilerPath = $command.Source
    } else {
        $candidates = foreach ($base in @("$env:LOCALAPPDATA\Programs", ${env:ProgramFiles(x86)}, $env:ProgramFiles)) {
            if ($base) { Join-Path $base 'Inno Setup 6\ISCC.exe' }
        }
        $InnoCompilerPath = $candidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
    }
}
if (-not $InnoCompilerPath -or -not (Test-Path -LiteralPath $InnoCompilerPath -PathType Leaf)) {
    throw 'Inno Setup 6.7 or newer is required. Install JRSoftware.InnoSetup with winget, or supply -InnoCompilerPath. Use Build.ps1 -SkipInstaller only when a developer ZIP is sufficient.'
}
$output = Join-Path $root 'artifacts'
$arguments = @(
    "/DAppVersion=$version",
    "/DAssemblyVersion=$($assembly.Version)",
    "/DAssemblyIdentity=$($assembly.FullName)",
    "/DPackageDir=$PackagePath",
    "/DOutputDir=$output"
)
$name = "OrgLens-$version-Setup"
if ($TestRunId) {
    $arguments += "/DTestRunId=$TestRunId"
    $name += "-Test-$TestRunId"
    if ($TestOutlookPath) { $arguments += "/DTestOutlookPath=$TestOutlookPath" }
    if ($TestProcessName) { $arguments += "/DTestProcessName=$TestProcessName" }
}
$arguments += Join-Path $root 'installer\OrgLens.iss'
& $InnoCompilerPath @arguments | Out-Host
if ($LASTEXITCODE -ne 0) { throw 'OrgLens installer compilation failed.' }
$path = Join-Path $output "$name.exe"
if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Installer output is missing: $path" }
$hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
[IO.File]::WriteAllText("$path.sha256", "$hash  $name.exe`n", [Text.Encoding]::ASCII)
[pscustomobject]@{ Path = $path; ChecksumPath = "$path.sha256"; Version = $version }
