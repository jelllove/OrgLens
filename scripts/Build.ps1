[CmdletBinding()]
param([switch]$SkipTests)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
Push-Location $root
try {
    & dotnet build .\OrgLens.sln --configuration Release --nologo
    if ($LASTEXITCODE -ne 0) { throw 'OrgLens build failed.' }
    if (-not $SkipTests) {
        & .\tests\OrgLens.Tests\bin\Release\net48\OrgLens.Tests.exe
        if ($LASTEXITCODE -ne 0) { throw 'OrgLens checks failed.' }
        $ui = Start-Process .\src\OrgLens.Preview\bin\Release\net48\OrgLens.Preview.exe -ArgumentList '--smoke-test' -PassThru -Wait
        if ($ui.ExitCode -ne 0) { throw 'OrgLens UI smoke test failed.' }
    }
    $version = ([xml](Get-Content (Join-Path $root 'Directory.Build.props') -Raw)).Project.PropertyGroup.Version
    $packageName = "OrgLens-$version"
    $package = Join-Path $root "artifacts\$packageName"
    $null = New-Item $package -ItemType Directory -Force
    foreach ($project in @('OrgLens.Outlook', 'OrgLens.Preview')) {
        Get-ChildItem (Join-Path $root "src\$project\bin\Release\net48") -File |
            Where-Object { $_.Extension -in '.dll', '.exe', '.config' } |
            Copy-Item -Destination $package -Force
    }
    Copy-Item (Join-Path $PSScriptRoot 'Install.ps1'), (Join-Path $PSScriptRoot 'Uninstall.ps1') -Destination $package -Force
    foreach ($file in @('README.md', 'LICENSE', 'CHANGELOG.md')) {
        Copy-Item (Join-Path $root $file) -Destination $package -Force
    }
    $images = Join-Path $package 'docs\images'
    $null = New-Item $images -ItemType Directory -Force
    Copy-Item (Join-Path $root 'docs\images\orglens-settings.png') -Destination $images -Force
    $assets = Get-Content (Join-Path $root 'src\OrgLens.Core\obj\project.assets.json') -Raw | ConvertFrom-Json
    $dependency = $assets.libraries.PSObject.Properties | Where-Object { $_.Name -like 'Newtonsoft.Json/*' }
    $relativeLicense = Join-Path ($dependency.Value.path.Replace('/', '\')) 'LICENSE.md'
    $license = $assets.packageFolders.PSObject.Properties | ForEach-Object {
        Join-Path $_.Name $relativeLicense
    } | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $license) { throw 'The Newtonsoft.Json distribution license could not be found.' }
    Copy-Item $license (Join-Path $package 'Newtonsoft.Json.LICENSE.md') -Force
    $archive = Join-Path $root "artifacts\$packageName.zip"
    Compress-Archive -Path $package -DestinationPath $archive -Force
    $hash = (Get-FileHash $archive -Algorithm SHA256).Hash.ToLowerInvariant()
    [IO.File]::WriteAllText("$archive.sha256", "$hash  $packageName.zip`n", [Text.Encoding]::ASCII)
    Write-Host "Package ready: $package"
    Write-Host 'Run OrgLens.Preview.exe for the sample-only demo; Install.ps1 registers the real Outlook add-in.'
}
finally { Pop-Location }
