#region Manually entered recipients
<#
    Parses typed or pasted recipients into TbmManualEntry objects. Accepted forms (one or many per line,
    separated by new lines, semicolons, or commas between addresses):
        jane@contoso.com
        Jane Doe <jane@contoso.com>
        "Doe, Jane" <jane@contoso.com>      Doe, Jane <jane@contoso.com>
        Jane Doe<TAB>jane@contoso.com       (cells copied from Excel)
        Jane Doe                            (name only: looked up in the directory when Graph is connected)
#>

$script:TbmEmailRegex = [System.Text.RegularExpressions.Regex]::new("[A-Za-z0-9._%+'\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}", [System.Text.RegularExpressions.RegexOptions]::Compiled)
$script:TbmNameJunkRegex = [System.Text.RegularExpressions.Regex]::new('mailto:|[<>"()\[\]\t]', [System.Text.RegularExpressions.RegexOptions]::Compiled -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
$script:TbmSpaceRunRegex = [System.Text.RegularExpressions.Regex]::new('\s{2,}', [System.Text.RegularExpressions.RegexOptions]::Compiled)

function ConvertFrom-TbmManualRecipientText {
    <#
    .SYNOPSIS
        Splits free text into recipient entries and validates them (status OK / Invalid email / Duplicate / Needs email).
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[TeamsBlastMessenger.TbmManualEntry]])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Text
    )

    $entries = [System.Collections.Generic.List[TeamsBlastMessenger.TbmManualEntry]]::new()
    $lines = $Text.Replace("`r`n", "`n").Replace("`r", "`n").Split([char[]]@("`n", ';'), [System.StringSplitOptions]::RemoveEmptyEntries)
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $found = $script:TbmEmailRegex.Matches($line)

        if ($found.Count -eq 0) {
            # Skip a pasted header row such as "Name<TAB>Email".
            if ($line.IndexOf('email', [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or $line.IndexOf('e-mail', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { continue }
            # A line with only an "@" is a mistyped address, not a name.
            $entries.Add((New-TbmManualEntry -Raw $line.Trim() -NameText $line -Email $null -LooksLikeEmail $line.Contains('@')))
            continue
        }

        # One entry per address; the text before each address (after the previous one) is its name.
        $start = 0
        foreach ($match in $found) {
            $segment = $line.Substring($start, $match.Index + $match.Length - $start)
            $nameText = $line.Substring($start, $match.Index - $start)
            $start = $match.Index + $match.Length
            # "jane@x.com (Jane Doe)" / "jane@x.com<TAB>Jane Doe": with one address and no name before it, use the text after it.
            if ($found.Count -eq 1 -and [string]::IsNullOrWhiteSpace($script:TbmNameJunkRegex.Replace($nameText, ''))) {
                $nameText = $line.Substring($start)
                $segment = $line
            }
            $entries.Add((New-TbmManualEntry -Raw $segment.Trim(' ', ',', '>', "`t") -NameText $nameText -Email $match.Value))
        }
    }

    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($entry in $entries) {
        if ($entry.Status -eq 'OK' -and -not $seen.Add($entry.Email)) {
            $entry.Status = 'Duplicate'
            $entry.Detail = 'Listed more than once.'
        }
    }
    return , $entries
}

function New-TbmManualEntry {
    [CmdletBinding()]
    [OutputType([TeamsBlastMessenger.TbmManualEntry])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Raw,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$NameText,

        [Parameter()]
        [AllowNull()]
        [string]$Email,

        [Parameter()]
        [bool]$LooksLikeEmail = $false
    )

    $entry = [TeamsBlastMessenger.TbmManualEntry]::new()
    $entry.Raw = $Raw
    Set-TbmManualEntryName -Entry $entry -Name $NameText

    if ([string]::IsNullOrEmpty($Email)) {
        if ($LooksLikeEmail) {
            $entry.Name = ''; $entry.FirstName = ''; $entry.LastName = ''
            $entry.Email = $Raw
            $entry.Status = 'Invalid email'
            $entry.Detail = 'Not a valid email address.'
        }
        else {
            $entry.Status = 'Needs email'
            $entry.Detail = 'Name only. Connect to Graph and tick "Look up names" to find the address, or type it.'
        }
        return $entry
    }

    $entry.Email = $Email.Trim().Trim('.')
    if (Test-TbmEmailAddress -Value $entry.Email) {
        $entry.Status = 'OK'
    }
    else {
        $entry.Status = 'Invalid email'
        $entry.Detail = 'Not a valid email address.'
    }
    return $entry
}

function Set-TbmManualEntryName {
    <#
    .SYNOPSIS
        Cleans a display name and splits it into first/last ("Doe, Jane" and "Jane Doe" both give First=Jane, Last=Doe).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [TeamsBlastMessenger.TbmManualEntry]$Entry,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Name
    )

    $clean = $script:TbmNameJunkRegex.Replace($Name, ' ')
    $clean = $script:TbmSpaceRunRegex.Replace($clean, ' ').Trim(' ', ',', ':', '-', "'")
    $first = ''
    $last = ''
    $display = $clean
    $comma = $clean.IndexOf(',')
    if ($comma -gt 0) {
        $last = $clean.Substring(0, $comma).Trim()
        $first = $clean.Substring($comma + 1).Trim(' ', ',')
        $display = ('{0} {1}' -f $first, $last).Trim()
    }
    elseif ($clean.Length -gt 0) {
        $space = $clean.IndexOf(' ')
        if ($space -gt 0) {
            $first = $clean.Substring(0, $space)
            $last = $clean.Substring($space + 1).Trim()
        }
        else {
            $first = $clean
        }
    }
    $Entry.Name = $display
    $Entry.FirstName = $first
    $Entry.LastName = $last
}

function Resolve-TbmManualEntryNames {
    <#
    .SYNOPSIS
        Looks up name-only entries in the directory. A single match fills in the email; several matches are reported.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[TeamsBlastMessenger.TbmManualEntry]]$Entries
    )

    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($entry in $Entries) { if ($entry.Status -eq 'OK') { $null = $seen.Add($entry.Email) } }

    foreach ($entry in $Entries) {
        if ($entry.Status -ne 'Needs email' -or [string]::IsNullOrWhiteSpace($entry.Name)) { continue }
        try {
            $users = Find-TbmDirectoryUser -AuthState $AuthState -Name $entry.Name
        }
        catch {
            $entry.Status = 'Not found'
            $entry.Detail = 'Directory lookup failed: ' + (Get-TbmFailureMessage -Exception $_.Exception)
            continue
        }
        if ($users.Count -eq 0) {
            $entry.Status = 'Not found'
            $entry.Detail = 'No one with this name in the directory.'
            continue
        }
        if ($users.Count -gt 1) {
            $names = [System.Collections.Generic.List[string]]::new()
            foreach ($user in $users) { $names.Add(('{0} <{1}>' -f $user.DisplayName, $user.Email)) }
            $entry.Status = 'Ambiguous'
            $entry.Detail = 'Several matches; type the email instead: ' + [string]::Join('; ', $names)
            continue
        }
        $user = $users[0]
        $entry.Email = $user.Email
        Set-TbmManualEntryName -Entry $entry -Name $user.DisplayName
        if ($seen.Add($user.Email)) {
            $entry.Status = 'OK'
            $entry.Detail = 'Found in directory.'
        }
        else {
            $entry.Status = 'Duplicate'
            $entry.Detail = 'Listed more than once.'
        }
    }
}

function Find-TbmDirectoryUser {
    <#
    .SYNOPSIS
        Finds users by display name: exact match first, then a directory search. Returns up to 5 users with an email.
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[pscustomobject]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    $select = '$select=id,displayName,mail,userPrincipalName&$top=5'
    $filter = [System.Uri]::EscapeDataString(("displayName eq '{0}'" -f $Name.Replace("'", "''")))
    $response = Invoke-TbmGraphRequest -AuthState $AuthState -Path ('/users?$filter={0}&{1}' -f $filter, $select)
    $values = Get-TbmPropertyValue -InputObject $response -Name 'value'
    if ($null -eq $values -or @($values).Count -eq 0) {
        $search = [System.Uri]::EscapeDataString(('"displayName:{0}"' -f $Name.Replace('"', '')))
        $response = Invoke-TbmGraphRequest -AuthState $AuthState -Path ('/users?$search={0}&$count=true&{1}' -f $search, $select) -Headers @{ ConsistencyLevel = 'eventual' }
        $values = Get-TbmPropertyValue -InputObject $response -Name 'value'
    }

    $users = [System.Collections.Generic.List[pscustomobject]]::new()
    if ($null -eq $values) { return , $users }
    foreach ($value in $values) {
        $email = [string](Get-TbmPropertyValue -InputObject $value -Name 'mail')
        if ([string]::IsNullOrEmpty($email)) { $email = [string](Get-TbmPropertyValue -InputObject $value -Name 'userPrincipalName') }
        if ([string]::IsNullOrEmpty($email)) { continue }
        $users.Add([pscustomobject]@{ DisplayName = [string]$value.displayName; Email = $email })
    }
    return , $users
}

function Add-TbmManualEntriesToTable {
    <#
    .SYNOPSIS
        Adds valid entries to a recipient table, or builds a new table (Email, Name, First Name, Last Name) when none is given.
        Name parts go into matching columns of an existing spreadsheet when it has them.
    .OUTPUTS
        PSCustomObject: Table, RecipientColumn, Added (int).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[TeamsBlastMessenger.TbmManualEntry]]$Entries,

        [Parameter()]
        [System.Data.DataTable]$Table,

        [Parameter()]
        [string]$RecipientColumn
    )

    if ($null -eq $Table) {
        $rows = [System.Collections.Generic.List[string[]]]::new()
        $rows.Add([string[]]@('Email', 'Name', 'First Name', 'Last Name'))
        $Table = ConvertTo-TbmRecipientTable -Rows $rows
        $RecipientColumn = 'Email'
    }
    if ([string]::IsNullOrEmpty($RecipientColumn)) { throw [System.InvalidOperationException]::new('Choose the recipient (email) column before adding people.') }

    # Map name parts onto the spreadsheet's own columns, if present.
    $targets = @{ Name = $null; First = $null; Last = $null }
    foreach ($column in (Get-TbmDataColumnNames -Table $Table)) {
        switch (ConvertTo-TbmTokenKey -Name $column) {
            { $_ -in @('name', 'displayname', 'fullname') } { if ($null -eq $targets.Name) { $targets.Name = $column } }
            { $_ -in @('firstname', 'givenname', 'first', 'forename') } { if ($null -eq $targets.First) { $targets.First = $column } }
            { $_ -in @('lastname', 'surname', 'last', 'familyname') } { if ($null -eq $targets.Last) { $targets.Last = $column } }
        }
    }

    $added = 0
    foreach ($entry in $Entries) {
        if (-not $entry.IsValid) { continue }
        $row = $Table.NewRow()
        foreach ($column in $Table.Columns) { if ($column.DataType -eq [string]) { $row[$column] = '' } }
        $row[$script:TbmIncludeColumn] = $true
        $row[$RecipientColumn] = $entry.Email
        if ($null -ne $targets.Name) { $row[$targets.Name] = [string]$entry.Name }
        if ($null -ne $targets.First) { $row[$targets.First] = [string]$entry.FirstName }
        if ($null -ne $targets.Last) { $row[$targets.Last] = [string]$entry.LastName }
        $Table.Rows.Add($row)
        $added++
    }
    Write-TbmLog -Level Info -Source 'Recipients' -Message ('Added {0} manually entered recipient(s).' -f $added)
    return [pscustomobject]@{ Table = $Table; RecipientColumn = $RecipientColumn; Added = $added }
}

#endregion
