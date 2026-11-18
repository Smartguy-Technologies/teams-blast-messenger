#region Paths

function Get-TbmAppDataPath {
    <#
    .SYNOPSIS
        Returns a well-known Teams Blast Messenger path under %APPDATA%\TeamsBlastMessenger.
    .DESCRIPTION
        The root can be redirected with the TBM_APPDATA environment variable (used by tests).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter()]
        [ValidateSet('Root', 'Logs', 'History', 'Cache', 'Exports', 'Settings', 'TokenCache')]
        [string]$Name = 'Root'
    )

    $root = [System.Environment]::GetEnvironmentVariable('TBM_APPDATA')
    if ([string]::IsNullOrWhiteSpace($root)) {
        $appData = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ApplicationData)
        $root = [System.IO.Path]::Combine($appData, 'TeamsBlastMessenger')
    }

    switch ($Name) {
        'Root'       { return $root }
        'Logs'       { return [System.IO.Path]::Combine($root, 'Logs') }
        'History'    { return [System.IO.Path]::Combine($root, 'History') }
        'Cache'      { return [System.IO.Path]::Combine($root, 'Cache') }
        'Exports'    { return [System.IO.Path]::Combine($root, 'Exports') }
        'Settings'   { return [System.IO.Path]::Combine($root, 'settings.json') }
        'TokenCache' { return [System.IO.Path]::Combine($root, 'Cache', 'token.bin') }
    }
}

function Initialize-TbmEnvironment {
    <#
    .SYNOPSIS
        Creates the app-data folders, loads settings and initialises the shared logger.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()

    foreach ($name in @('Root', 'Logs', 'History', 'Cache', 'Exports')) {
        $null = [System.IO.Directory]::CreateDirectory((Get-TbmAppDataPath -Name $name))
    }

    # Start the log first so policy and settings warnings reach the log file.
    [TeamsBlastMessenger.TbmLogger]::Initialize((Get-TbmAppDataPath -Name 'Logs'), [TeamsBlastMessenger.TbmLogLevel]::Info, 0)
    Write-TbmLog -Level Info -Source 'App' -Message ('Teams Blast Messenger starting. PowerShell {0}, {1}.' -f $PSVersionTable.PSVersion, [System.Environment]::OSVersion.VersionString)

    $settings = Get-TbmSettings
    Write-TbmPolicySummary
    $level = [TeamsBlastMessenger.TbmLogLevel]::Info
    $null = [System.Enum]::TryParse([string]$settings.LogLevel, $true, [ref]$level)
    [TeamsBlastMessenger.TbmLogger]::MinimumLevel = $level
    [TeamsBlastMessenger.TbmLogger]::Purge([int]$settings.LogRetentionDays)
    return $settings
}

#endregion

#region Logging

function Write-TbmLog {
    <#
    .SYNOPSIS
        Writes an entry to the shared, thread-safe log (file + UI queue).
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [TeamsBlastMessenger.TbmLogLevel]$Level = [TeamsBlastMessenger.TbmLogLevel]::Info,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Source,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Message,

        [Parameter()]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    if ($null -ne $ErrorRecord) {
        $position = ''
        if ($null -ne $ErrorRecord.InvocationInfo) { $position = $ErrorRecord.InvocationInfo.PositionMessage }
        $detail = '{0}{1}{2}{1}{3}' -f $Message, [System.Environment]::NewLine, $ErrorRecord.Exception.ToString(), $position
        [TeamsBlastMessenger.TbmLogger]::Write([TeamsBlastMessenger.TbmLogLevel]::Error, $Source, $detail)
        return
    }
    [TeamsBlastMessenger.TbmLogger]::Write($Level, $Source, $Message)
}

#endregion

#region Settings and organization policy
<#
    Effective settings are layered, lowest to highest precedence:
        1. built-in defaults (Get-TbmSettingSchema)
        2. HKCU\SOFTWARE\Policies\TeamsBlastMessenger\Defaults
        3. HKLM\SOFTWARE\Policies\TeamsBlastMessenger\Defaults
        4. the user's %APPDATA%\TeamsBlastMessenger\settings.json
        5. HKCU\SOFTWARE\Policies\TeamsBlastMessenger\Enforced
        6. HKLM\SOFTWARE\Policies\TeamsBlastMessenger\Enforced
    "Defaults" are organization starting values users may change; "Enforced" values are locked.
    settings.json stores only what the user changed, so later changes to policy Defaults still reach them.
#>

$script:TbmPolicyRegistryPath = 'SOFTWARE\Policies\TeamsBlastMessenger'
$script:TbmPolicy = $null

function Get-TbmSettingSchema {
    <#
    .SYNOPSIS
        Describes every setting: default, type, allowed values or range, label and whether policy may set it.
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Specialized.OrderedDictionary])]
    param()

    return [ordered]@{
        AuthMode                 = @{ Default = 'Auto'; Type = 'String'; Allowed = @('Auto', 'GraphSdk', 'GraphRest'); Label = 'Sign-in method'; Policy = $true }
        ClientId                 = @{ Default = ''; Type = 'String'; Format = 'Guid'; Label = 'Client ID'; Policy = $true }
        TenantId                 = @{ Default = 'organizations'; Type = 'String'; Format = 'Tenant'; Label = 'Tenant ID'; Policy = $true }
        Transport                = @{ Default = 'Auto'; Type = 'String'; Allowed = @('Auto', 'Graph', 'DeepLinkAuto', 'DeepLinkManual'); Label = 'Send via'; Policy = $true }
        MessageFormat            = @{ Default = 'Text'; Type = 'String'; Allowed = @('Text', 'Html'); Label = 'Message format'; Policy = $true }
        DeepLinkTarget           = @{ Default = 'Desktop'; Type = 'String'; Allowed = @('Desktop', 'Web'); Label = 'Deep links open in'; Policy = $true }
        DeepLinkSendDelaySeconds = @{ Default = 6; Type = 'Int'; Min = 1; Max = 120; Label = 'Auto-send wait (seconds)'; Policy = $true }
        DeepLinkMaxUrlLength     = @{ Default = 2000; Type = 'Int'; Min = 500; Max = 32000; Label = 'Max deep-link length'; Policy = $true }
        PacingDelayMs            = @{ Default = 1000; Type = 'Int'; Min = 0; Max = 60000; Label = 'Delay between messages (ms)'; Policy = $true }
        TestRecipient            = @{ Default = ''; Type = 'String'; Format = 'Email'; Label = 'Test recipient'; Policy = $true }
        LogLevel                 = @{ Default = 'Info'; Type = 'String'; Allowed = @('Debug', 'Info', 'Warn', 'Error'); Label = 'Log level'; Policy = $true }
        LogRetentionDays         = @{ Default = 30; Type = 'Int'; Min = 1; Max = 3650; Label = 'Keep logs (days)'; Policy = $true }
        InteractionMaxChatPages  = @{ Default = 40; Type = 'Int'; Min = 1; Max = 400; Label = 'Chat pages to search'; Policy = $true }
        LastFolder               = @{ Default = ''; Type = 'String'; Label = 'Last folder'; Policy = $false }
    }
}

function Get-TbmDefaultSettings {
    <#
    .SYNOPSIS
        Returns the built-in defaults (before any policy or user value).
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()

    $defaults = @{}
    $schema = Get-TbmSettingSchema
    foreach ($name in $schema.Keys) { $defaults[$name] = $schema[$name].Default }
    return $defaults
}

function Resolve-TbmSettingName {
    <#
    .SYNOPSIS
        Returns the canonical setting name for a case-insensitive match, or $null if unknown.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Name
    )
    foreach ($candidate in (Get-TbmSettingSchema).Keys) {
        if ([string]::Equals($candidate, $Name.Trim(), [System.StringComparison]::OrdinalIgnoreCase)) { return $candidate }
    }
    return $null
}

function Test-TbmSettingValue {
    <#
    .SYNOPSIS
        Converts and validates a value for a setting (used for registry policy, settings.json and the Settings dialog).
    .OUTPUTS
        PSCustomObject: Valid (bool), Value (normalised value), Error (message).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter()]
        [AllowNull()]
        [object]$Value
    )

    $schema = Get-TbmSettingSchema
    if (-not $schema.Contains($Name)) { return [pscustomobject]@{ Valid = $false; Value = $null; Error = ('Unknown setting "{0}".' -f $Name) } }
    $spec = $schema[$Name]
    $label = $spec.Label

    if ($spec.Type -eq 'Int') {
        $number = [long]0
        $ok = $false
        if ($Value -is [int] -or $Value -is [long] -or $Value -is [uint32]) {
            $number = [long]$Value
            $ok = $true
        }
        elseif ($null -ne $Value) {
            $ok = [long]::TryParse(([string]$Value).Trim(), [System.Globalization.NumberStyles]::Integer, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)
        }
        if (-not $ok -or $number -lt $spec.Min -or $number -gt $spec.Max) {
            return [pscustomobject]@{ Valid = $false; Value = $null; Error = ('{0} must be a whole number from {1} to {2}.' -f $label, $spec.Min, $spec.Max) }
        }
        return [pscustomobject]@{ Valid = $true; Value = [int]$number; Error = $null }
    }

    $text = ''
    if ($null -ne $Value) { $text = ([string]$Value).Trim() }

    if ($spec.ContainsKey('Allowed')) {
        foreach ($allowed in $spec.Allowed) {
            if ([string]::Equals($allowed, $text, [System.StringComparison]::OrdinalIgnoreCase)) { return [pscustomobject]@{ Valid = $true; Value = $allowed; Error = $null } }
        }
        return [pscustomobject]@{ Valid = $false; Value = $null; Error = ('{0} must be one of: {1}.' -f $label, [string]::Join(', ', $spec.Allowed)) }
    }

    switch ([string]$spec['Format']) {
        'Guid' {
            $guid = [guid]::Empty
            if ($text.Length -gt 0 -and -not [guid]::TryParse($text, [ref]$guid)) {
                return [pscustomobject]@{ Valid = $false; Value = $null; Error = ('{0} must be a GUID (Application ID from the app registration).' -f $label) }
            }
        }
        'Email' {
            if ($text.Length -gt 0 -and -not (Test-TbmEmailAddress -Value $text)) {
                return [pscustomobject]@{ Valid = $false; Value = $null; Error = ('{0} must be an email address.' -f $label) }
            }
        }
        'Tenant' {
            if ($text.Length -eq 0) { $text = 'organizations' }
            if ($text.IndexOfAny([char[]]@(' ', '/', '\', '?', '#')) -ge 0) {
                return [pscustomobject]@{ Valid = $false; Value = $null; Error = ('{0} must be a tenant GUID or domain (e.g. contoso.onmicrosoft.com).' -f $label) }
            }
        }
    }
    return [pscustomobject]@{ Valid = $true; Value = $text; Error = $null }
}

function Read-TbmRegistryPolicy {
    <#
    .SYNOPSIS
        Reads the raw values under HKLM/HKCU\SOFTWARE\Policies\TeamsBlastMessenger\<Section>.
        Returns $null off Windows or when the key does not exist.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('LocalMachine', 'CurrentUser')]
        [string]$Hive,

        [Parameter(Mandatory)]
        [ValidateSet('Defaults', 'Enforced')]
        [string]$Section
    )

    if (-not (Test-TbmIsWindows)) { return $null }
    $registryHive = [Microsoft.Win32.RegistryHive]::LocalMachine
    if ($Hive -eq 'CurrentUser') { $registryHive = [Microsoft.Win32.RegistryHive]::CurrentUser }
    # 64-bit view so 32-bit and 64-bit PowerShell read the same policy.
    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($registryHive, [Microsoft.Win32.RegistryView]::Registry64)
    try {
        $key = $baseKey.OpenSubKey(($script:TbmPolicyRegistryPath + '\' + $Section), $false)
        if ($null -eq $key) { return $null }
        try {
            $values = @{}
            foreach ($valueName in $key.GetValueNames()) {
                if ([string]::IsNullOrEmpty($valueName)) { continue }
                $values[$valueName] = $key.GetValue($valueName)
            }
            return $values
        }
        finally {
            $key.Dispose()
        }
    }
    finally {
        $baseKey.Dispose()
    }
}

function Get-TbmPolicy {
    <#
    .SYNOPSIS
        Loads and validates organization policy from the registry into $script:TbmPolicy.
    .OUTPUTS
        Hashtable: Defaults, Enforced (setting -> value), Sources (setting -> registry key that supplied it).
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()

    $policy = @{ Defaults = @{}; Enforced = @{}; Sources = @{} }
    $hiveLabels = [ordered]@{ CurrentUser = 'HKCU'; LocalMachine = 'HKLM' }   # HKLM read last so it wins.
    foreach ($section in @('Defaults', 'Enforced')) {
        foreach ($hive in $hiveLabels.Keys) {
            $source = '{0}\{1}\{2}' -f $hiveLabels[$hive], $script:TbmPolicyRegistryPath, $section
            $raw = $null
            try {
                $raw = Read-TbmRegistryPolicy -Hive $hive -Section $section
            }
            catch {
                Write-TbmLog -Level Warn -Source 'Policy' -Message ('Could not read {0}: {1}' -f $source, $_.Exception.Message)
            }
            if ($null -eq $raw) { continue }
            foreach ($valueName in $raw.Keys) {
                $name = Resolve-TbmSettingName -Name ([string]$valueName)
                if ($null -eq $name -or -not (Get-TbmSettingSchema)[$name].Policy) {
                    Write-TbmLog -Level Warn -Source 'Policy' -Message ('Ignoring unknown policy value "{0}" in {1}.' -f $valueName, $source)
                    continue
                }
                $check = Test-TbmSettingValue -Name $name -Value $raw[$valueName]
                if (-not $check.Valid) {
                    Write-TbmLog -Level Warn -Source 'Policy' -Message ('Ignoring invalid policy value {0} in {1}: {2}' -f $name, $source, $check.Error)
                    continue
                }
                $policy[$section][$name] = $check.Value
                if ($section -eq 'Enforced' -or -not $policy.Enforced.ContainsKey($name)) { $policy.Sources[$name] = $source }
            }
        }
    }
    $script:TbmPolicy = $policy
    return $policy
}

function Get-TbmPolicyState {
    <#
    .SYNOPSIS
        Returns the loaded policy, loading it on first use.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()
    if ($null -eq $script:TbmPolicy) { return Get-TbmPolicy }
    return $script:TbmPolicy
}

function Test-TbmSettingEnforced {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )
    return (Get-TbmPolicyState).Enforced.ContainsKey($Name)
}

function Get-TbmPolicySource {
    <#
    .SYNOPSIS
        The registry key that supplied a setting's policy value, or $null.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )
    $sources = (Get-TbmPolicyState).Sources
    if ($sources.ContainsKey($Name)) { return [string]$sources[$Name] }
    return $null
}

function Get-TbmEffectiveDefaults {
    <#
    .SYNOPSIS
        Built-in defaults overlaid with policy Defaults: what a user gets before changing anything.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()
    $defaults = Get-TbmDefaultSettings
    $policy = Get-TbmPolicyState
    foreach ($name in $policy.Defaults.Keys) { $defaults[$name] = $policy.Defaults[$name] }
    return $defaults
}

function Merge-TbmSettings {
    <#
    .SYNOPSIS
        Applies the precedence: built-in < policy Defaults < user < policy Enforced.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [hashtable]$PolicyDefaults,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [hashtable]$User,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [hashtable]$PolicyEnforced
    )

    $settings = Get-TbmDefaultSettings
    foreach ($layer in @($PolicyDefaults, $User, $PolicyEnforced)) {
        foreach ($name in $layer.Keys) { $settings[$name] = $layer[$name] }
    }
    return $settings
}

function Read-TbmUserSettings {
    <#
    .SYNOPSIS
        Reads the user's settings.json; unknown keys and invalid values are ignored.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()

    $user = @{}
    $path = Get-TbmAppDataPath -Name 'Settings'
    if (-not [System.IO.File]::Exists($path)) { return $user }
    try {
        $json = [System.IO.File]::ReadAllText($path)
        if ([string]::IsNullOrWhiteSpace($json)) { return $user }
        $stored = ConvertFrom-Json -InputObject $json
        foreach ($property in $stored.PSObject.Properties) {
            $name = Resolve-TbmSettingName -Name $property.Name
            if ($null -eq $name) { continue }
            $check = Test-TbmSettingValue -Name $name -Value $property.Value
            if ($check.Valid) { $user[$name] = $check.Value }
            else { Write-TbmLog -Level Warn -Source 'Settings' -Message ('Ignoring invalid saved value for {0}: {1}' -f $name, $check.Error) }
        }
    }
    catch {
        Write-TbmLog -Source 'Settings' -Message 'Could not read settings.json; using defaults.' -ErrorRecord $_
    }
    return $user
}

function Get-TbmSettings {
    <#
    .SYNOPSIS
        Returns the effective settings (built-in, policy Defaults, user, policy Enforced). Reloads policy.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()

    $policy = Get-TbmPolicy
    return Merge-TbmSettings -PolicyDefaults $policy.Defaults -User (Read-TbmUserSettings) -PolicyEnforced $policy.Enforced
}

function Save-TbmSettings {
    <#
    .SYNOPSIS
        Saves only the settings the user changed from the effective defaults; enforced settings are never saved.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Settings
    )

    $defaults = Get-TbmEffectiveDefaults
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $changed = [ordered]@{}
    foreach ($name in (Get-TbmSettingSchema).Keys) {
        if (-not $Settings.ContainsKey($name) -or (Test-TbmSettingEnforced -Name $name)) { continue }
        $value = $Settings[$name]
        $valueText = [System.Convert]::ToString($value, $culture)
        $defaultText = [System.Convert]::ToString($defaults[$name], $culture)
        if ([string]::Equals($valueText, $defaultText, [System.StringComparison]::Ordinal)) { continue }
        $changed[$name] = $value
    }

    $json = ConvertTo-Json -InputObject $changed -Depth 4
    Write-TbmTextFileAtomic -Path (Get-TbmAppDataPath -Name 'Settings') -Content $json
    Write-TbmLog -Level Info -Source 'Settings' -Message ('Settings saved ({0} personal override(s)).' -f $changed.Count)
}

function Write-TbmPolicySummary {
    <#
    .SYNOPSIS
        Logs which settings come from organization policy, and from where.
    #>
    [CmdletBinding()]
    param()

    $policy = Get-TbmPolicyState
    if ($policy.Defaults.Count -eq 0 -and $policy.Enforced.Count -eq 0) {
        Write-TbmLog -Level Info -Source 'Policy' -Message 'No organization policy found.'
        return
    }
    foreach ($section in @('Defaults', 'Enforced')) {
        $entries = [System.Collections.Generic.List[string]]::new()
        foreach ($name in $policy[$section].Keys) {
            $source = $policy.Sources[$name]
            if ($section -eq 'Defaults' -and $policy.Enforced.ContainsKey($name)) { $source = 'overridden by Enforced' }
            $entries.Add(('{0}={1} [{2}]' -f $name, $policy[$section][$name], $source))
        }
        $entries.Sort([System.StringComparer]::OrdinalIgnoreCase)
        if ($entries.Count -gt 0) { Write-TbmLog -Level Info -Source 'Policy' -Message ('Policy {0}: {1}' -f $section, [string]::Join('; ', $entries)) }
    }
}

#endregion

#region File and data helpers

function Write-TbmTextFileAtomic {
    <#
    .SYNOPSIS
        Writes UTF-8 (no BOM) text to a temp file and swaps it into place so readers never see a partial file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Content
    )

    $directory = [System.IO.Path]::GetDirectoryName($Path)
    $null = [System.IO.Directory]::CreateDirectory($directory)
    $tempPath = $Path + '.tmp'
    [System.IO.File]::WriteAllText($tempPath, $Content, [System.Text.UTF8Encoding]::new($false))
    if ([System.IO.File]::Exists($Path)) {
        # [NullString]::Value: PowerShell would otherwise pass '' (not null) for the backup path.
        [System.IO.File]::Replace($tempPath, $Path, [NullString]::Value)
    }
    else {
        [System.IO.File]::Move($tempPath, $Path)
    }
}

function Write-TbmCsvFile {
    <#
    .SYNOPSIS
        Writes an RFC 4180 CSV file (UTF-8 with BOM so Excel detects the encoding).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Header,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string[]]]$Rows
    )

    $builder = [System.Text.StringBuilder]::new()
    $fields = [System.Collections.Generic.List[string]]::new()
    foreach ($name in $Header) { $fields.Add([TeamsBlastMessenger.TbmText]::CsvField($name)) }
    $null = $builder.AppendLine([string]::Join(',', $fields))
    foreach ($row in $Rows) {
        $fields.Clear()
        foreach ($value in $row) { $fields.Add([TeamsBlastMessenger.TbmText]::CsvField($value)) }
        $null = $builder.AppendLine([string]::Join(',', $fields))
    }
    [System.IO.File]::WriteAllText($Path, $builder.ToString(), [System.Text.UTF8Encoding]::new($true))
    Write-TbmLog -Level Info -Source 'Export' -Message ('Exported {0} row(s) to {1}' -f $Rows.Count, $Path)
}

function ConvertTo-TbmUtcDateTime {
    <#
    .SYNOPSIS
        Converts a DateTime or ISO-8601 string (as returned by Graph or stored in history) to a UTC DateTime.
        Returns $null for empty input.
    #>
    [CmdletBinding()]
    [OutputType([System.Nullable[datetime]])]
    param(
        [Parameter()]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [System.DateTimeKind]::Unspecified) { return [datetime]::SpecifyKind($Value, [System.DateTimeKind]::Utc) }
        return $Value.ToUniversalTime()
    }
    if ($Value -is [System.DateTimeOffset]) { return $Value.UtcDateTime }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $styles = [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal
    return [System.DateTimeOffset]::Parse($text, [System.Globalization.CultureInfo]::InvariantCulture, $styles).UtcDateTime
}

function ConvertTo-TbmIsoString {
    <#
    .SYNOPSIS
        Formats a UTC DateTime as ISO-8601 with milliseconds (Graph filter / history format).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter()]
        [AllowNull()]
        [System.Nullable[datetime]]$Value
    )

    if ($null -eq $Value) { return $null }
    return ([datetime]$Value).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-TbmPropertyValue {
    <#
    .SYNOPSIS
        Safely reads a (possibly nested, dot-separated) property from a PSCustomObject/hashtable under StrictMode.
    #>
    [CmdletBinding()]
    [OutputType([object])]
    param(
        [Parameter()]
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    $current = $InputObject
    foreach ($part in $Name.Split('.')) {
        if ($null -eq $current) { return $null }
        if ($current -is [System.Collections.IDictionary]) {
            if (-not $current.Contains($part)) { return $null }
            $current = $current[$part]
            continue
        }
        $property = $current.PSObject.Properties[$part]
        if ($null -eq $property) { return $null }
        $current = $property.Value
    }
    return $current
}

function Test-TbmIsWindows {
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    return [System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT
}

#endregion
