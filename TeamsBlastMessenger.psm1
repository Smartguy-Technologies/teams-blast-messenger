#Requires -Version 5.1
<#
    Teams Blast Messenger root module.
    Loads the shared .NET types once per AppDomain, then dot-sources Private and Public functions.
    The module is imported both by the UI runspace and by every background worker runspace.
#>
Set-StrictMode -Version Latest

$script:TbmModuleRoot   = $PSScriptRoot
$script:TbmManifestPath = [System.IO.Path]::Combine($PSScriptRoot, 'TeamsBlastMessenger.psd1')

foreach ($assemblyName in @('System.Net.Http', 'System.IO.Compression', 'System.IO.Compression.FileSystem', 'Microsoft.VisualBasic', 'System.Security')) {
    try {
        Add-Type -AssemblyName $assemblyName -ErrorAction Stop
    }
    catch {
        # Not every assembly exists on every platform (e.g. System.Security DPAPI on non-Windows). Callers check before use.
        Write-Verbose -Message ("Assembly '{0}' could not be loaded: {1}" -f $assemblyName, $_.Exception.Message)
    }
}

if ($null -eq ('TeamsBlastMessenger.TbmLogger' -as [type])) {
    $typeSourcePath = [System.IO.Path]::Combine($PSScriptRoot, 'Types', 'TbmTypes.cs')
    $typeSource     = [System.IO.File]::ReadAllText($typeSourcePath)
    $addTypeParameters = @{ TypeDefinition = $typeSource; Language = 'CSharp'; ErrorAction = 'Stop' }
    if ($PSVersionTable.PSEdition -ne 'Core') {
        # Windows PowerShell compiles against System.dll only; name the other assemblies the types use.
        # (PowerShell 7 already references the whole framework, and naming them again would conflict.)
        $addTypeParameters['ReferencedAssemblies'] = @('System.Xml', [System.Management.Automation.PSObject].Assembly.Location)
    }
    Add-Type @addTypeParameters
}

foreach ($folderName in @('Private', 'Public')) {
    $folderPath = [System.IO.Path]::Combine($PSScriptRoot, $folderName)
    if (-not [System.IO.Directory]::Exists($folderPath)) { continue }
    $scriptFiles = [System.IO.Directory]::GetFiles($folderPath, '*.ps1')
    [System.Array]::Sort($scriptFiles, [System.StringComparer]::OrdinalIgnoreCase)
    foreach ($scriptFile in $scriptFiles) {
        . $scriptFile
    }
}

Export-ModuleMember -Function @('Show-TbmMainWindow', 'Invoke-TbmSendWorker', 'Invoke-TbmInteractionWorker')
