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
        [void][System.IO.Directory]::CreateDirectory((Get-TbmAppDataPath -Name $name))
    }

    $settings = Get-TbmSettings
    $level = [TeamsBlastMessenger.TbmLogLevel]::Info
    [void][System.Enum]::TryParse([string]$settings.LogLevel, $true, [ref]$level)
    [TeamsBlastMessenger.TbmLogger]::Initialize((Get-TbmAppDataPath -Name 'Logs'), $level, [int]$settings.LogRetentionDays)
    Write-TbmLog -Level Info -Source 'App' -Message ('Teams Blast Messenger starting. PowerShell {0}, {1}.' -f $PSVersionTable.PSVersion, [System.Environment]::OSVersion.VersionString)
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

#region Settings

function Get-TbmDefaultSettings {
    <#
    .SYNOPSIS
        Returns the default settings. The type of each default defines the stored type.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()

    return @{
        AuthMode                 = 'Auto'          # Auto | GraphSdk | GraphRest
        ClientId                 = ''
        TenantId                 = 'organizations'
        Transport                = 'Auto'          # Auto | Graph | DeepLinkAuto | DeepLinkManual
        DeepLinkTarget           = 'Desktop'       # Desktop (msteams:) | Web (https:)
        DeepLinkSendDelaySeconds = 6
        DeepLinkMaxUrlLength     = 2000
        PacingDelayMs            = 1000
        TestRecipient            = ''
        MessageFormat            = 'Text'          # Text | Html
        LogLevel                 = 'Info'
        LogRetentionDays         = 30
        InteractionMaxChatPages  = 40
        LastFolder               = ''
    }
}

function Get-TbmSettings {
    <#
    .SYNOPSIS
        Loads settings.json merged over the defaults. Unknown keys are ignored; bad values fall back to defaults.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()

    $settings = Get-TbmDefaultSettings
    $path = Get-TbmAppDataPath -Name 'Settings'
    if (-not [System.IO.File]::Exists($path)) { return $settings }

    try {
        $json = [System.IO.File]::ReadAllText($path)
        if ([string]::IsNullOrWhiteSpace($json)) { return $settings }
        $stored = ConvertFrom-Json -InputObject $json
        foreach ($property in $stored.PSObject.Properties) {
            if (-not $settings.ContainsKey($property.Name)) { continue }
            $defaultValue = $settings[$property.Name]
            try {
                $settings[$property.Name] = [System.Management.Automation.LanguagePrimitives]::ConvertTo($property.Value, $defaultValue.GetType())
            }
            catch {
                Write-TbmLog -Level Warn -Source 'Settings' -Message ("Ignoring invalid value for '{0}'." -f $property.Name)
            }
        }
    }
    catch {
        Write-TbmLog -Source 'Settings' -Message 'Could not read settings.json; using defaults.' -ErrorRecord $_
    }
    return $settings
}

function Save-TbmSettings {
    <#
    .SYNOPSIS
        Persists settings to %APPDATA%\TeamsBlastMessenger\settings.json (atomic write).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Settings
    )

    $ordered = [ordered]@{}
    $keys = [System.Collections.Generic.List[string]]::new()
    foreach ($key in $Settings.Keys) { $keys.Add([string]$key) }
    $keys.Sort([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($key in $keys) { $ordered[$key] = $Settings[$key] }

    $json = ConvertTo-Json -InputObject $ordered -Depth 4
    Write-TbmTextFileAtomic -Path (Get-TbmAppDataPath -Name 'Settings') -Content $json
    Write-TbmLog -Level Info -Source 'Settings' -Message 'Settings saved.'
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
    [void][System.IO.Directory]::CreateDirectory($directory)
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
    [void]$builder.AppendLine([string]::Join(',', $fields))
    foreach ($row in $Rows) {
        $fields.Clear()
        foreach ($value in $row) { $fields.Add([TeamsBlastMessenger.TbmText]::CsvField($value)) }
        [void]$builder.AppendLine([string]::Join(',', $fields))
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
