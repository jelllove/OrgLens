[CmdletBinding(SupportsShouldProcess)]
param()
$ErrorActionPreference = 'Stop'
if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) {
    if (-not $WhatIfPreference) {
        throw 'Close classic Outlook before uninstalling. OrgLens will never terminate Outlook automatically.'
    }
    Write-Warning 'Outlook is running. A real uninstall requires closing it first.'
}
if (-not $PSCmdlet.ShouldProcess('Current user OrgLens COM registration', 'Unregister OrgLens')) { return }
$id = '{D7E2D48A-9466-4D83-856F-AC197BC23A98}'
foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
    try {
        $key = $base.OpenSubKey("Software\Classes\CLSID\$id\InprocServer32")
        if ($key) {
            try {
                if ($key.GetValue('Class') -ne 'OrgLens.Outlook.Connect') {
                    throw 'COM class ownership did not match OrgLens. No further keys will be removed.'
                }
            }
            finally { $key.Dispose() }
        }
        $base.DeleteSubKeyTree('Software\Microsoft\Office\Outlook\Addins\OrgLens.Connect', $false)
        $base.DeleteSubKeyTree('Software\Classes\OrgLens.Connect', $false)
        $base.DeleteSubKeyTree("Software\Classes\CLSID\$id", $false)
    }
    finally { $base.Dispose() }
}
Write-Host 'OrgLens is unregistered. Installed files are retained in %LOCALAPPDATA%\OrgLens\Addin and may be deleted.'
Write-Host 'Outlook views were NOT changed. To remove formatting, use Remove in OrgLens before uninstalling,'
Write-Host 'or select a different view using View > Change View in Outlook.'
