#region Token template engine
<#
    Syntax
      {Column Name}          value from the row; matching ignores case, spaces and punctuation
                             ({first name}, {First_Name} and {FirstName} all match "First Name")
      {Column Name|default}  default used when the cell is empty
      {{ and }}              literal braces
      Built-ins: {Today} {Now} {SenderName} {SenderFirstName} {SenderEmail}
    Spreadsheet columns take precedence over built-ins with the same name.
#>

$script:TbmTokenRegex = [System.Text.RegularExpressions.Regex]::new('\{\{|\}\}|\{([^{}|]+)(?:\|([^{}]*))?\}', [System.Text.RegularExpressions.RegexOptions]::Compiled)
$script:TbmNonAlphanumericRegex = [System.Text.RegularExpressions.Regex]::new('[^\p{L}\p{N}]+', [System.Text.RegularExpressions.RegexOptions]::Compiled)
$script:TbmBuiltInTokenNames = @('Today', 'Now', 'SenderName', 'SenderFirstName', 'SenderEmail')

function ConvertTo-TbmTokenKey {
    <#
    .SYNOPSIS
        Normalises a token or column name for matching: letters and digits only, lower case.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Name
    )
    return $script:TbmNonAlphanumericRegex.Replace($Name, '').ToLowerInvariant()
}

function Get-TbmBuiltInTokenValues {
    <#
    .SYNOPSIS
        Returns built-in token values for the signed-in sender (empty sender values when not connected).
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter()]
        [hashtable]$AuthState
    )

    $culture = [System.Globalization.CultureInfo]::CurrentCulture
    $now = [datetime]::Now
    $senderName = ''
    $senderEmail = ''
    if ($null -ne $AuthState) {
        $senderName = [string]$AuthState.MeName
        $senderEmail = [string]$AuthState.MeUpn
    }
    $firstName = ''
    if (-not [string]::IsNullOrWhiteSpace($senderName)) {
        $firstName = $senderName.Trim().Split([char[]]@(' '), [System.StringSplitOptions]::RemoveEmptyEntries)[0]
    }
    return @{
        Today           = $now.ToString('d', $culture)
        Now             = $now.ToString('g', $culture)
        SenderName      = $senderName
        SenderFirstName = $firstName
        SenderEmail     = $senderEmail
    }
}

function New-TbmTokenContext {
    <#
    .SYNOPSIS
        Builds the lookup used to resolve tokens: normalised key -> column name, and normalised key -> built-in value.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [string[]]$ColumnNames,

        [Parameter()]
        [hashtable]$BuiltIns
    )

    $columnMap = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::Ordinal)
    $collisions = [System.Collections.Generic.List[string]]::new()
    foreach ($column in $ColumnNames) {
        $key = ConvertTo-TbmTokenKey -Name $column
        if ([string]::IsNullOrEmpty($key)) { continue }
        if ($columnMap.ContainsKey($key)) {
            $collisions.Add(('"{0}" and "{1}"' -f $columnMap[$key], $column))
            continue
        }
        $columnMap[$key] = $column
    }

    $builtInMap = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::Ordinal)
    if ($null -ne $BuiltIns) {
        foreach ($name in $BuiltIns.Keys) { $builtInMap[(ConvertTo-TbmTokenKey -Name ([string]$name))] = [string]$BuiltIns[$name] }
    }

    return @{
        Columns    = $columnMap
        BuiltIns   = $builtInMap
        Collisions = $collisions
    }
}

function Get-TbmTemplateTokens {
    <#
    .SYNOPSIS
        Lists the tokens used in a template (escaped braces excluded).
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[pscustomobject]])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Template
    )

    $tokens = [System.Collections.Generic.List[pscustomobject]]::new()
    foreach ($match in $script:TbmTokenRegex.Matches($Template)) {
        if (-not $match.Groups[1].Success) { continue }
        $name = $match.Groups[1].Value.Trim()
        $tokens.Add([pscustomobject]@{
                Raw        = $match.Value
                Name       = $name
                Key        = ConvertTo-TbmTokenKey -Name $name
                HasDefault = $match.Groups[2].Success
                Default    = $(if ($match.Groups[2].Success) { $match.Groups[2].Value } else { $null })
            })
    }
    return , $tokens
}

function Test-TbmTemplate {
    <#
    .SYNOPSIS
        Validates a template against the token context.
    .OUTPUTS
        PSCustomObject: UnknownTokens, ColumnTokens (distinct names), UsesRowData (bool), TokenCount.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Template,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Context
    )

    $unknown = [System.Collections.Generic.List[string]]::new()
    $columnTokens = [System.Collections.Generic.List[string]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    $tokens = Get-TbmTemplateTokens -Template $Template
    foreach ($token in $tokens) {
        if (-not $seen.Add($token.Key)) { continue }
        if ($Context.Columns.ContainsKey($token.Key)) { $columnTokens.Add($token.Name) }
        elseif (-not $Context.BuiltIns.ContainsKey($token.Key)) { $unknown.Add($token.Name) }
    }
    return [pscustomobject]@{
        UnknownTokens = $unknown
        ColumnTokens  = $columnTokens
        UsesRowData   = $columnTokens.Count -gt 0
        TokenCount    = $tokens.Count
    }
}

function ConvertTo-TbmRowDictionary {
    <#
    .SYNOPSIS
        Snapshots a DataRow's spreadsheet values into a case-insensitive dictionary (safe to hand to other threads).
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.Dictionary[string, string]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Data.DataRow]$Row,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [string[]]$ColumnNames
    )

    $values = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($column in $ColumnNames) {
        $value = $Row[$column]
        if ($value -is [System.DBNull]) { $values[$column] = '' } else { $values[$column] = [string]$value }
    }
    return , $values
}

function Expand-TbmTemplate {
    <#
    .SYNOPSIS
        Replaces tokens with row / built-in values.
    .DESCRIPTION
        Unknown tokens are left as typed so they are visible in previews. Empty values without a default
        are reported in MissingTokens. Use -HtmlEncodeValues when the template itself is HTML.
    .OUTPUTS
        PSCustomObject: Text, MissingTokens (List[string]), UnknownTokens (List[string]).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Template,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Context,

        [Parameter()]
        [System.Collections.Generic.Dictionary[string, string]]$Row,

        [Parameter()]
        [switch]$HtmlEncodeValues
    )

    $builder = [System.Text.StringBuilder]::new($Template.Length + 64)
    $missing = [System.Collections.Generic.List[string]]::new()
    $unknown = [System.Collections.Generic.List[string]]::new()
    $last = 0

    foreach ($match in $script:TbmTokenRegex.Matches($Template)) {
        $null = $builder.Append($Template, $last, $match.Index - $last)
        $last = $match.Index + $match.Length

        if ($match.Value -eq '{{') { $null = $builder.Append('{'); continue }
        if ($match.Value -eq '}}') { $null = $builder.Append('}'); continue }

        $name = $match.Groups[1].Value.Trim()
        $key = ConvertTo-TbmTokenKey -Name $name
        $resolved = $false
        $value = $null
        $columnName = $null
        if ($Context.Columns.TryGetValue($key, [ref]$columnName)) {
            $resolved = $true
            if ($null -ne $Row -and $Row.ContainsKey($columnName)) { $value = $Row[$columnName] }
        }
        else {
            $builtInValue = $null
            if ($Context.BuiltIns.TryGetValue($key, [ref]$builtInValue)) {
                $resolved = $true
                $value = $builtInValue
            }
        }

        if (-not $resolved) {
            $unknown.Add($name)
            $null = $builder.Append($match.Value)
            continue
        }

        if ([string]::IsNullOrWhiteSpace($value)) {
            if ($match.Groups[2].Success) { $value = $match.Groups[2].Value }
            else { $missing.Add($name); $value = '' }
        }
        else {
            $value = $value.Trim()
        }
        if ($HtmlEncodeValues) { $value = [System.Net.WebUtility]::HtmlEncode($value) }
        $null = $builder.Append($value)
    }
    $null = $builder.Append($Template, $last, $Template.Length - $last)

    return [pscustomobject]@{
        Text          = $builder.ToString()
        MissingTokens = $missing
        UnknownTokens = $unknown
    }
}

function ConvertTo-TbmGraphHtml {
    <#
    .SYNOPSIS
        Converts an expanded message to the HTML body sent through Graph.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Message,

        [Parameter(Mandatory)]
        [ValidateSet('Text', 'Html')]
        [string]$Format
    )
    if ($Format -eq 'Html') { return $Message }
    return [TeamsBlastMessenger.TbmText]::TextToHtml($Message)
}

function ConvertTo-TbmPlainText {
    <#
    .SYNOPSIS
        Converts an expanded message to plain text (deep links only support plain text).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Message,

        [Parameter(Mandatory)]
        [ValidateSet('Text', 'Html')]
        [string]$Format
    )
    if ($Format -eq 'Text') { return $Message }
    return [TeamsBlastMessenger.TbmText]::StripHtml($Message)
}

#endregion
