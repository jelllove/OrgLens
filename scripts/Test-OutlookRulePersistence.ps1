[CmdletBinding()]
param(
    [string]$AssemblyDirectory,
    [switch]$DemonstrateViewSaveBug,
    [switch]$DemonstrateViewApplyBug
)
$ErrorActionPreference = 'Stop'
if (-not $AssemblyDirectory) {
    $AssemblyDirectory = Join-Path (Split-Path $PSScriptRoot -Parent) 'src\OrgLens.Outlook\bin\Release\net48'
}
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Run this test with Windows PowerShell: powershell.exe -NoProfile -STA -File scripts\Test-OutlookRulePersistence.ps1'
}
$null = [Reflection.Assembly]::LoadFrom((Join-Path $AssemblyDirectory 'OrgLens.Core.dll'))
$assembly = [Reflection.Assembly]::LoadFrom((Join-Path $AssemblyDirectory 'OrgLens.Outlook.dll'))
$application = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
$explorer = $application.ActiveExplorer()
if (-not $explorer) { throw 'Open a mail folder in classic Outlook before running this explicit integration test.' }
$folder = $explorer.CurrentFolder
$views = $folder.Views
$originalViewName = $explorer.CurrentView.Name
$originalCount = $views.Count
$name = 'OrgLens persistence test ' + [Guid]::NewGuid().ToString('N')
$temporary = $null
$copySource = $null
function Assert-PlannedRules($NativeRules, $ExpectedRules) {
    foreach ($expected in $ExpectedRules) {
        $actual = $NativeRules.Item($expected.Name)
        if (-not [string]::Equals($actual.Filter, $expected.Filter, [StringComparison]::Ordinal) -or
            $actual.Enabled -ne $expected.Style.Enabled -or
            $actual.Font.Size -ne $expected.Style.FontSize -or
            $actual.Font.Color -ne [int]$expected.Style.Color -or
            $actual.Font.Bold -ne $expected.Style.Bold -or
            $actual.Font.Italic -ne $expected.Style.Italic) {
            throw ('Native condition/style persistence failed for ' + $expected.Name)
        }
    }
    $owned = 0
    for ($i = 1; $i -le $NativeRules.Count; $i++) {
        $rule = $NativeRules.Item($i)
        if (-not $rule.Standard -and [OrgLens.Core.RuleDefinition]::IsOwned($rule.Name)) { $owned++ }
    }
    if ($owned -ne $ExpectedRules.Count) { throw 'The persisted owned-rule count differs from the plan.' }
}
function Get-UnownedRuleState($NativeRules) {
    $state = @(
        for ($i = 1; $i -le $NativeRules.Count; $i++) {
            $rule = $NativeRules.Item($i)
            if (-not $rule.Standard -and [OrgLens.Core.RuleDefinition]::IsOwned($rule.Name)) { continue }
            $font = $rule.Font
            [pscustomobject]@{
                Name = $rule.Name; Standard = $rule.Standard; Filter = $rule.Filter; Enabled = $rule.Enabled
                FontName = $font.Name; Size = $font.Size; Color = $font.Color; ExtendedColor = $font.ExtendedColor
                Bold = $font.Bold; Italic = $font.Italic; Underline = $font.Underline; Strikethrough = $font.Strikethrough
            }
        }
    )
    return ConvertTo-Json -InputObject $state -Compress
}
try {
    $temporary = $views.Add($name, 0, 1)
    $collection = $temporary.AutoFormatRules
    $unrelated = $collection.Add('Unrelated fixture rule')
    $unrelated.Filter = '"http://schemas.microsoft.com/mapi/proptag/0x0037001f" CI_STARTSWITH ''ORGLENS_FIXTURE'''
    $unrelated.Enabled = $true
    $collection.Save()
    $unrelatedFilter = $unrelated.Filter
    $unownedState = Get-UnownedRuleState $views.Item($name).AutoFormatRules
    $adapterType = $assembly.GetType('OrgLens.Outlook.NativeFormattingRules', $true)
    $arguments = New-Object object[] 1
    $arguments[0] = $collection
    $adapter = $adapterType.GetConstructors()[0].Invoke($arguments)
    $configuration = [OrgLens.Core.OrgLensConfiguration]::new()
    $configuration.CustomAddresses.Add('custom@example.com')
    $people = [OrgLens.Core.ManagerPerson[]]@(
        [OrgLens.Core.ManagerPerson]::new('fixture-manager', 'Fixture manager', 'manager@example.com', '/o=Fixture/cn=manager', 1),
        [OrgLens.Core.ManagerPerson]::new('fixture-leader', 'Fixture leader', 'leader@example.com', '/o=Fixture/cn=leader', 2)
    )
    $planned = [OrgLens.Core.RuleDefinition]::Build($configuration, $people)
    [OrgLens.Core.RuleReconciler]::Replace($adapter, $planned)
    if ($DemonstrateViewSaveBug) { $temporary.Save() }
    Write-Host ('Before activation: ' + $views.Item($name).AutoFormatRules.Item('OrgLens.v2:BossUnread').Filter.Length)
    if ($DemonstrateViewApplyBug) { $temporary.Apply() }
    else { $explorer.CurrentView = $name }
    Write-Host ('After activation: ' + $views.Item($name).AutoFormatRules.Item('OrgLens.v2:BossUnread').Filter.Length)
    $explorer.CurrentView = $originalViewName
    $explorer.CurrentView = $name
    $saved = $views.Item($name)
    $savedRules = $saved.AutoFormatRules
    foreach ($expected in $planned) {
        $actual = $savedRules.Item($expected.Name)
        Write-Host ($expected.Name + ': filter characters after Save = ' + $actual.Filter.Length)
        if (-not [string]::Equals($actual.Filter, $expected.Filter, [StringComparison]::Ordinal)) {
            throw ('Native persistence failed for ' + $expected.Name + ': its saved condition differs from the generated condition.')
        }
        if ($actual.Enabled -ne $expected.Style.Enabled -or $actual.Font.Size -ne $expected.Style.FontSize) {
            throw ('Native style persistence failed for ' + $expected.Name)
        }
    }
    if ($savedRules.Item('Unrelated fixture rule').Filter -ne $unrelatedFilter) {
        throw 'An unrelated formatting condition was changed.'
    }
    Write-Host 'PASS: all four real Outlook conditions and font sizes survived saving, activation and view switching.'

    $arguments[0] = $views.Item($name).AutoFormatRules
    $adapter = $adapterType.GetConstructors()[0].Invoke($arguments)
    $configuration.BossUnread.FontSize = 17
    $configuration.BossRead.Enabled = $false
    $planned = [OrgLens.Core.RuleDefinition]::Build($configuration, $people)
    [OrgLens.Core.RuleReconciler]::Replace($adapter, $planned)
    $explorer.CurrentView = $name
    $arguments[0] = $explorer.CurrentView.AutoFormatRules
    $adapter = $adapterType.GetConstructors()[0].Invoke($arguments)
    $adapter.Verify($planned)
    Write-Host 'PASS: updating an already active view retains its complete conditions.'

    foreach ($missing in @($planned.Name) + @('ALL')) {
        $manual = $explorer.CurrentView.AutoFormatRules
        for ($i = $manual.Count; $i -ge 1; $i--) {
            $rule = $manual.Item($i)
            if (-not $rule.Standard -and [OrgLens.Core.RuleDefinition]::IsOwned($rule.Name) -and
                ($missing -eq 'ALL' -or $rule.Name -eq $missing)) {
                $manual.Remove($i)
            }
        }
        $manual.Save()
        $arguments[0] = $views.Item($name).AutoFormatRules
        $adapter = $adapterType.GetConstructors()[0].Invoke($arguments)
        $survivors = $adapter.CaptureOwned()
        $expectedCount = if ($missing -eq 'ALL') { 0 } else { $planned.Count - 1 }
        if ($survivors.Count -ne $expectedCount) { throw 'Manual deletion was not persisted before Apply.' }
        [OrgLens.Core.RuleReconciler]::Replace($adapter, $planned)
        $explorer.CurrentView = $name
        $arguments[0] = $explorer.CurrentView.AutoFormatRules
        $adapter = $adapterType.GetConstructors()[0].Invoke($arguments)
        $adapter.Verify($planned)
        Assert-PlannedRules $views.Item($name).AutoFormatRules $planned
        if ((Get-UnownedRuleState $views.Item($name).AutoFormatRules) -ne $unownedState) {
            throw 'Recreating a manually deleted rule changed an unrelated or built-in rule.'
        }
        Write-Host ('PASS: Apply recreates manually deleted ' + $missing + ' with complete conditions and styles.')
    }

    $collection = $views.Item($name).AutoFormatRules
    for ($i = $collection.Count; $i -ge 1; $i--) {
        if ($collection.Item($i).Name -eq $planned[0].Name) { $collection.Remove($i) }
    }
    $collection.Save()
    $arguments[0] = $views.Item($name).AutoFormatRules
    $adapter = $adapterType.GetConstructors()[0].Invoke($arguments)
    $rejected = $false
    try { $adapter.Verify($planned) }
    catch [System.Management.Automation.MethodInvocationException] {
        if ($_.Exception.InnerException -isnot [OrgLens.Core.OrgLensException]) { throw }
        $rejected = $true
    }
    if (-not $rejected) { throw 'The saved-condition verifier accepted a missing native rule.' }
    foreach ($rule in $adapter.CaptureOwned()) {
        if ($rule.Enabled) { throw 'A missing native rule left surviving OrgLens rules enabled.' }
    }
    [OrgLens.Core.RuleReconciler]::Replace($adapter, $planned)
    $explorer.CurrentView = $name
    Write-Host 'PASS: verification fails closed when a native rule is missing.'

    $rollback = $adapter.CaptureOwned()
    [OrgLens.Core.RuleReconciler]::RemoveOwned($adapter)
    $adapter.RestoreOwned($rollback)
    $restored = $views.Item($name).AutoFormatRules
    foreach ($expected in $planned) {
        $actual = $restored.Item($expected.Name)
        if ($actual.Filter -ne $expected.Filter -or $actual.Enabled -ne $expected.Style.Enabled -or
            $actual.Font.Size -ne $expected.Style.FontSize -or
            $actual.Font.Color -ne [int]$expected.Style.Color) {
            throw ('Native rollback failed for ' + $expected.Name)
        }
    }
    Write-Host 'PASS: native rule snapshots restore conditions and styles independently of View.XML.'

    $collection = $views.Item($name).AutoFormatRules
    $unsafe = $collection.Item('OrgLens.v2:BossUnread')
    $unsafe.Filter = ''; $unsafe.Enabled = $true; $collection.Save()
    $arguments[0] = $views.Item($name).AutoFormatRules
    $adapter = $adapterType.GetConstructors()[0].Invoke($arguments)
    $unsafeSnapshot = $adapter.CaptureOwned()
    $rejected = $false
    try { $adapter.Verify($planned) }
    catch [System.Management.Automation.MethodInvocationException] {
        if ($_.Exception.InnerException -isnot [OrgLens.Core.OrgLensException]) { throw }
        $rejected = $true
    }
    if (-not $rejected) { throw 'The saved-condition verifier accepted a missing sender condition.' }
    $disabled = $views.Item($name).AutoFormatRules
    foreach ($expected in $planned) {
        if ($disabled.Item($expected.Name).Enabled) { throw 'A condition verification failure left an OrgLens rule enabled.' }
    }
    Write-Host 'PASS: missing conditions are rejected and unsafe rules are disabled.'
    $adapter.RestoreOwned($unsafeSnapshot)
    if ($views.Item($name).AutoFormatRules.Item('OrgLens.v2:BossUnread').Enabled) {
        throw 'Rollback reactivated an unsafe old empty-condition rule.'
    }
    Write-Host 'PASS: rollback never reactivates a legacy match-all condition.'

    [OrgLens.Core.RuleReconciler]::RemoveOwned($adapter)
    $remaining = $views.Item($name).AutoFormatRules
    for ($i=1; $i -le $remaining.Count; $i++) {
        if ([OrgLens.Core.RuleDefinition]::IsOwned($remaining.Item($i).Name)) {
            throw 'An OrgLens rule remained after persisted removal.'
        }
    }
    if ($remaining.Item('Unrelated fixture rule').Filter -ne $unrelatedFilter) {
        throw 'Removal changed an unrelated condition.'
    }
    Write-Host 'PASS: removal is persisted and preserves unrelated rules.'

    $copySource = $views.Add(('OrgLens recreation source ' + [Guid]::NewGuid().ToString('N')), 0, 1)
    $explorer.CurrentView = $copySource.Name
    $folder.Views.Item($name).Delete()
    $temporary = $null
    $views = $folder.Views
    $sourceRules = $views.Item($copySource.Name).AutoFormatRules
    $sourceUnrelated = $sourceRules.Add('Unrelated fixture rule')
    $sourceUnrelated.Filter = $unrelatedFilter
    $sourceUnrelated.Enabled = $true
    $sourceUnrelated.Font.Name = 'Arial'
    $sourceUnrelated.Font.Size = 13
    $sourceUnrelated.Font.Bold = $true
    $sourceUnrelated.Font.Italic = $true
    $sourceUnrelated.Font.Underline = $true
    $sourceUnrelated.Font.Strikethrough = $true
    $sourceUnrelated.Font.ExtendedColor = 5
    $interleaved = $sourceRules.Insert('Unrelated interleaved fixture', $sourceRules.Count)
    $interleaved.Filter = '"urn:schemas:httpmail:read" = 0'
    $interleaved.Enabled = $false
    $matchAll = $sourceRules.Add('Unrelated intentional match-all fixture')
    $matchAll.Filter = ''
    $matchAll.Enabled = $true
    $standard = $sourceRules.Item(1)
    $standard.Enabled = $false
    $standard.Font.Size = 15
    $standard.Font.Italic = $true
    $standard.Font.Color = [int][OrgLens.Core.MailColor]::Blue
    $sourceRules.Save()
    $explorer.CurrentView = $copySource.Name
    $sourceState = Get-UnownedRuleState $views.Item($copySource.Name).AutoFormatRules
    $currentSource = $folder.CurrentView
    $factory = [OrgLens.Outlook.OutlookOrgLensService].GetMethod('CreatePrivateView',
        [Reflection.BindingFlags]'NonPublic,Static')
    $temporary = $factory.Invoke($null, [object[]]@($views, $currentSource, $name))
    if ($explorer.CurrentView.Name -ne $copySource.Name) {
        throw 'Creating the private view unexpectedly changed the current view selection.'
    }
    if ((Get-UnownedRuleState $views.Item($copySource.Name).AutoFormatRules) -ne $sourceState) {
        throw 'Creating the private view changed its source view rules.'
    }
    if ((Get-UnownedRuleState $views.Item($name).AutoFormatRules) -ne $sourceState) {
        throw 'Creating the private view did not preserve all unrelated and built-in rule attributes.'
    }
    $arguments[0] = $temporary.AutoFormatRules
    $adapter = $adapterType.GetConstructors()[0].Invoke($arguments)
    if ($adapter.CaptureOwned().Count -ne 0) { throw 'The recreated view unexpectedly contains owned rules.' }
    [OrgLens.Core.RuleReconciler]::Replace($adapter, $planned)
    $explorer.CurrentView = $name
    $explorer.CurrentView = $copySource.Name
    $explorer.CurrentView = $name
    $arguments[0] = $explorer.CurrentView.AutoFormatRules
    $adapter = $adapterType.GetConstructors()[0].Invoke($arguments)
    $adapter.Verify($planned)
    Assert-PlannedRules $views.Item($name).AutoFormatRules $planned
    if ((Get-UnownedRuleState $views.Item($name).AutoFormatRules) -ne $sourceState -or
        (Get-UnownedRuleState $views.Item($copySource.Name).AutoFormatRules) -ne $sourceState) {
        throw 'Recreating the private view changed an unrelated or built-in rule.'
    }
    Write-Host 'PASS: Apply recreates a deleted private view and preserves its source conditions.'
}
finally {
    $explorer.CurrentView = $originalViewName
    if ($temporary) { $temporary.Delete() }
    if ($copySource) { $copySource.Delete() }
    if ($views.Count -ne $originalCount -or $explorer.CurrentView.Name -ne $originalViewName) {
        throw 'The original Outlook view state did not match after test cleanup. Inspect Outlook before continuing.'
    }
    Write-Host 'Temporary private view removed. Original view selection restored; its rules were not edited.'
}
