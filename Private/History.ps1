#region Send history
<#
    One JSON file per blast in %APPDATA%\TeamsBlastMessenger\History.
    During a send, each recipient outcome is appended to a ".journal" file (one JSON line, cheap and crash-safe);
    the full record is rewritten at start and end and the journal deleted. Loading merges any leftover journal,
    so an interrupted blast still shows exactly who was sent to.
#>

$script:TbmHistorySchemaVersion = 1

function Get-TbmBlastFilePath {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [TeamsBlastMessenger.TbmBlastRecord]$Blast
    )
    $name = '{0}_{1}.json' -f $Blast.CreatedUtc.ToString('yyyyMMdd-HHmmss', [System.Globalization.CultureInfo]::InvariantCulture), $Blast.Id
    return [System.IO.Path]::Combine((Get-TbmAppDataPath -Name 'History'), $name)
}

function ConvertTo-TbmRecipientHashtable {
    [CmdletBinding()]
    [OutputType([System.Collections.Specialized.OrderedDictionary])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [TeamsBlastMessenger.TbmRecipientResult]$Recipient
    )
    return [ordered]@{
        RowIndex        = $Recipient.RowIndex
        Email           = $Recipient.Email
        UserId          = $Recipient.UserId
        DisplayName     = $Recipient.DisplayName
        ChatId          = $Recipient.ChatId
        MessageId       = $Recipient.MessageId
        SentUtc         = ConvertTo-TbmIsoString -Value $Recipient.SentUtc
        Status          = $Recipient.Status.ToString()
        Error           = $Recipient.Error
        RenderedMessage = $Recipient.RenderedMessage
    }
}

function ConvertTo-TbmHistoryDocument {
    <#
    .SYNOPSIS
        Converts a blast record to an ordered dictionary with ISO-8601 UTC dates (stable across PS versions).
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Specialized.OrderedDictionary])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [TeamsBlastMessenger.TbmBlastRecord]$Blast
    )

    $recipients = [System.Collections.Generic.List[object]]::new()
    foreach ($recipient in $Blast.Recipients) { $recipients.Add((ConvertTo-TbmRecipientHashtable -Recipient $recipient)) }

    $interactions = [System.Collections.Generic.List[object]]::new()
    foreach ($interaction in $Blast.Interactions) {
        $replies = [System.Collections.Generic.List[object]]::new()
        foreach ($reply in $interaction.Replies) {
            $replies.Add([ordered]@{
                    MessageId     = $reply.MessageId
                    FromUserId    = $reply.FromUserId
                    FromName      = $reply.FromName
                    CreatedUtc    = ConvertTo-TbmIsoString -Value $reply.CreatedUtc
                    Text          = $reply.Text
                    IsDirectReply = $reply.IsDirectReply
                })
        }
        $interactions.Add([ordered]@{
                Email          = $interaction.Email
                DisplayName    = $interaction.DisplayName
                UserId         = $interaction.UserId
                ChatId         = $interaction.ChatId
                SentUtc        = ConvertTo-TbmIsoString -Value $interaction.SentUtc
                Responded      = $interaction.Responded
                ReplyCount     = $interaction.ReplyCount
                FirstReplyUtc  = ConvertTo-TbmIsoString -Value $interaction.FirstReplyUtc
                LastReplyUtc   = ConvertTo-TbmIsoString -Value $interaction.LastReplyUtc
                FirstReplyText = $interaction.FirstReplyText
                DirectReply    = $interaction.DirectReply
                Reactions      = $interaction.Reactions
                Error          = $interaction.Error
                Replies        = $replies.ToArray()
            })
    }

    return [ordered]@{
        SchemaVersion           = $script:TbmHistorySchemaVersion
        Id                      = $Blast.Id
        CreatedUtc              = ConvertTo-TbmIsoString -Value $Blast.CreatedUtc
        CompletedUtc            = ConvertTo-TbmIsoString -Value $Blast.CompletedUtc
        Mode                    = $Blast.Mode
        Transport               = $Blast.Transport
        Template                = $Blast.Template
        Format                  = $Blast.Format
        SourceFile              = $Blast.SourceFile
        Topic                   = $Blast.Topic
        SenderId                = $Blast.SenderId
        SenderUpn               = $Blast.SenderUpn
        IsTest                  = $Blast.IsTest
        LastInteractionCheckUtc = ConvertTo-TbmIsoString -Value $Blast.LastInteractionCheckUtc
        Recipients              = $recipients.ToArray()
        Interactions            = $interactions.ToArray()
    }
}

function Set-TbmRecipientFromData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [TeamsBlastMessenger.TbmRecipientResult]$Recipient,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [object]$Data
    )

    $Recipient.RowIndex = [int](Get-TbmPropertyValue -InputObject $Data -Name 'RowIndex')
    $Recipient.Email = [string](Get-TbmPropertyValue -InputObject $Data -Name 'Email')
    $Recipient.UserId = [string](Get-TbmPropertyValue -InputObject $Data -Name 'UserId')
    $Recipient.DisplayName = [string](Get-TbmPropertyValue -InputObject $Data -Name 'DisplayName')
    $Recipient.ChatId = [string](Get-TbmPropertyValue -InputObject $Data -Name 'ChatId')
    $Recipient.MessageId = [string](Get-TbmPropertyValue -InputObject $Data -Name 'MessageId')
    $Recipient.SentUtc = ConvertTo-TbmUtcDateTime -Value (Get-TbmPropertyValue -InputObject $Data -Name 'SentUtc')
    $Recipient.Error = [string](Get-TbmPropertyValue -InputObject $Data -Name 'Error')
    $Recipient.RenderedMessage = [string](Get-TbmPropertyValue -InputObject $Data -Name 'RenderedMessage')
    $status = [TeamsBlastMessenger.TbmRecipientStatus]::Pending
    $null = [System.Enum]::TryParse([string](Get-TbmPropertyValue -InputObject $Data -Name 'Status'), $true, [ref]$status)
    $Recipient.Status = $status
}

function ConvertFrom-TbmHistoryDocument {
    <#
    .SYNOPSIS
        Rebuilds a TbmBlastRecord from parsed history JSON.
    #>
    [CmdletBinding()]
    [OutputType([TeamsBlastMessenger.TbmBlastRecord])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [object]$Data
    )

    $blast = [TeamsBlastMessenger.TbmBlastRecord]::new()
    $blast.Id = [string]$Data.Id
    $blast.CreatedUtc = ConvertTo-TbmUtcDateTime -Value $Data.CreatedUtc
    $blast.CompletedUtc = ConvertTo-TbmUtcDateTime -Value (Get-TbmPropertyValue -InputObject $Data -Name 'CompletedUtc')
    $blast.Mode = [string](Get-TbmPropertyValue -InputObject $Data -Name 'Mode')
    $blast.Transport = [string](Get-TbmPropertyValue -InputObject $Data -Name 'Transport')
    $blast.Template = [string](Get-TbmPropertyValue -InputObject $Data -Name 'Template')
    $blast.Format = [string](Get-TbmPropertyValue -InputObject $Data -Name 'Format')
    $blast.SourceFile = [string](Get-TbmPropertyValue -InputObject $Data -Name 'SourceFile')
    $blast.Topic = [string](Get-TbmPropertyValue -InputObject $Data -Name 'Topic')
    $blast.SenderId = [string](Get-TbmPropertyValue -InputObject $Data -Name 'SenderId')
    $blast.SenderUpn = [string](Get-TbmPropertyValue -InputObject $Data -Name 'SenderUpn')
    $blast.IsTest = [bool](Get-TbmPropertyValue -InputObject $Data -Name 'IsTest')
    $blast.LastInteractionCheckUtc = ConvertTo-TbmUtcDateTime -Value (Get-TbmPropertyValue -InputObject $Data -Name 'LastInteractionCheckUtc')

    $recipients = Get-TbmPropertyValue -InputObject $Data -Name 'Recipients'
    if ($null -ne $recipients) {
        foreach ($item in $recipients) {
            $recipient = [TeamsBlastMessenger.TbmRecipientResult]::new()
            Set-TbmRecipientFromData -Recipient $recipient -Data $item
            $blast.Recipients.Add($recipient)
        }
    }

    $interactions = Get-TbmPropertyValue -InputObject $Data -Name 'Interactions'
    if ($null -ne $interactions) {
        foreach ($item in $interactions) {
            $interaction = [TeamsBlastMessenger.TbmInteractionResult]::new()
            $interaction.Email = [string](Get-TbmPropertyValue -InputObject $item -Name 'Email')
            $interaction.DisplayName = [string](Get-TbmPropertyValue -InputObject $item -Name 'DisplayName')
            $interaction.UserId = [string](Get-TbmPropertyValue -InputObject $item -Name 'UserId')
            $interaction.ChatId = [string](Get-TbmPropertyValue -InputObject $item -Name 'ChatId')
            $interaction.SentUtc = ConvertTo-TbmUtcDateTime -Value (Get-TbmPropertyValue -InputObject $item -Name 'SentUtc')
            $interaction.Responded = [bool](Get-TbmPropertyValue -InputObject $item -Name 'Responded')
            $interaction.ReplyCount = [int](Get-TbmPropertyValue -InputObject $item -Name 'ReplyCount')
            $interaction.FirstReplyUtc = ConvertTo-TbmUtcDateTime -Value (Get-TbmPropertyValue -InputObject $item -Name 'FirstReplyUtc')
            $interaction.LastReplyUtc = ConvertTo-TbmUtcDateTime -Value (Get-TbmPropertyValue -InputObject $item -Name 'LastReplyUtc')
            $interaction.FirstReplyText = [string](Get-TbmPropertyValue -InputObject $item -Name 'FirstReplyText')
            $interaction.DirectReply = [bool](Get-TbmPropertyValue -InputObject $item -Name 'DirectReply')
            $interaction.Reactions = [string](Get-TbmPropertyValue -InputObject $item -Name 'Reactions')
            $interaction.Error = [string](Get-TbmPropertyValue -InputObject $item -Name 'Error')
            $replies = Get-TbmPropertyValue -InputObject $item -Name 'Replies'
            if ($null -ne $replies) {
                foreach ($replyData in $replies) {
                    $reply = [TeamsBlastMessenger.TbmReplyMessage]::new()
                    $reply.MessageId = [string](Get-TbmPropertyValue -InputObject $replyData -Name 'MessageId')
                    $reply.FromUserId = [string](Get-TbmPropertyValue -InputObject $replyData -Name 'FromUserId')
                    $reply.FromName = [string](Get-TbmPropertyValue -InputObject $replyData -Name 'FromName')
                    $reply.CreatedUtc = ConvertTo-TbmUtcDateTime -Value (Get-TbmPropertyValue -InputObject $replyData -Name 'CreatedUtc')
                    $reply.Text = [string](Get-TbmPropertyValue -InputObject $replyData -Name 'Text')
                    $reply.IsDirectReply = [bool](Get-TbmPropertyValue -InputObject $replyData -Name 'IsDirectReply')
                    $interaction.Replies.Add($reply)
                }
            }
            $blast.Interactions.Add($interaction)
        }
    }
    return $blast
}

function Save-TbmBlastRecord {
    <#
    .SYNOPSIS
        Writes the full blast record (atomic) and removes its journal.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [TeamsBlastMessenger.TbmBlastRecord]$Blast
    )

    $path = Get-TbmBlastFilePath -Blast $Blast
    $json = ConvertTo-Json -InputObject (ConvertTo-TbmHistoryDocument -Blast $Blast) -Depth 8
    Write-TbmTextFileAtomic -Path $path -Content $json
    $journal = $path + '.journal'
    if ([System.IO.File]::Exists($journal)) { [System.IO.File]::Delete($journal) }
}

function Add-TbmBlastJournalEntry {
    <#
    .SYNOPSIS
        Appends one recipient outcome to the blast's journal (crash-safe progress record).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [TeamsBlastMessenger.TbmBlastRecord]$Blast,

        [Parameter(Mandatory)]
        [ValidateRange(0, [int]::MaxValue)]
        [int]$Index
    )

    $entry = ConvertTo-TbmRecipientHashtable -Recipient $Blast.Recipients[$Index]
    $entry['Index'] = $Index
    $line = (ConvertTo-Json -InputObject $entry -Compress -Depth 3) + [System.Environment]::NewLine
    [System.IO.File]::AppendAllText((Get-TbmBlastFilePath -Blast $Blast) + '.journal', $line, [System.Text.UTF8Encoding]::new($false))
}

function Read-TbmBlastRecord {
    <#
    .SYNOPSIS
        Loads a blast record from disk, applying any journal left by an interrupted send.
    #>
    [CmdletBinding()]
    [OutputType([TeamsBlastMessenger.TbmBlastRecord])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    $data = ConvertFrom-Json -InputObject ([System.IO.File]::ReadAllText($Path))
    $blast = ConvertFrom-TbmHistoryDocument -Data $data

    $journal = $Path + '.journal'
    if ([System.IO.File]::Exists($journal)) {
        foreach ($line in [System.IO.File]::ReadAllLines($journal)) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            try {
                $entry = ConvertFrom-Json -InputObject $line
                $index = [int]$entry.Index
                if ($index -ge 0 -and $index -lt $blast.Recipients.Count) {
                    Set-TbmRecipientFromData -Recipient $blast.Recipients[$index] -Data $entry
                }
            }
            catch {
                Write-TbmLog -Level Warn -Source 'History' -Message ('Skipping unreadable journal line in {0}.' -f $journal)
            }
        }
    }
    return $blast
}

function Get-TbmBlastHistory {
    <#
    .SYNOPSIS
        Loads all blast records, newest first. Unreadable files are logged and skipped.
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[TeamsBlastMessenger.TbmBlastRecord]])]
    param()

    $records = [System.Collections.Generic.List[TeamsBlastMessenger.TbmBlastRecord]]::new()
    $folder = Get-TbmAppDataPath -Name 'History'
    if (-not [System.IO.Directory]::Exists($folder)) { return , $records }
    foreach ($file in [System.IO.Directory]::GetFiles($folder, '*.json')) {
        try {
            $records.Add((Read-TbmBlastRecord -Path $file))
        }
        catch {
            Write-TbmLog -Level Warn -Source 'History' -Message ('Could not read history file {0}: {1}' -f $file, $_.Exception.Message)
        }
    }
    $records.Sort([System.Comparison[TeamsBlastMessenger.TbmBlastRecord]] { param($left, $right) $right.CreatedUtc.CompareTo($left.CreatedUtc) })
    return , $records
}

function Remove-TbmBlastRecord {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [TeamsBlastMessenger.TbmBlastRecord]$Blast
    )

    $path = Get-TbmBlastFilePath -Blast $Blast
    if (-not $PSCmdlet.ShouldProcess($path, 'Delete blast history')) { return }
    foreach ($file in @($path, $path + '.journal')) {
        if ([System.IO.File]::Exists($file)) { [System.IO.File]::Delete($file) }
    }
    Write-TbmLog -Level Info -Source 'History' -Message ('Deleted history record {0}.' -f $Blast.Id)
}

function Export-TbmBlastCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [TeamsBlastMessenger.TbmBlastRecord]$Blast,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    $rows = [System.Collections.Generic.List[string[]]]::new()
    foreach ($recipient in $Blast.Recipients) {
        $rows.Add([string[]]@(
                $recipient.Email, $recipient.DisplayName, $recipient.StatusText, $recipient.SentLocal,
                $recipient.ChatId, $recipient.MessageId, $recipient.Error,
                (ConvertTo-TbmPlainText -Message ([string]$recipient.RenderedMessage) -Format $(if ($Blast.Format -eq 'Html') { 'Html' } else { 'Text' }))
            ))
    }
    Write-TbmCsvFile -Path $Path -Rows $rows -Header @('Email', 'Display Name', 'Status', 'Sent (local)', 'Chat Id', 'Message Id', 'Error', 'Message')
}

function Export-TbmInteractionCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Collections.Generic.List[TeamsBlastMessenger.TbmInteractionResult]]$Interactions,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $rows = [System.Collections.Generic.List[string[]]]::new()
    foreach ($item in $Interactions) {
        $minutes = ''
        if ($null -ne $item.ResponseTime) { $minutes = $item.ResponseTime.TotalMinutes.ToString('0.##', $culture) }
        $allReplies = [System.Collections.Generic.List[string]]::new()
        foreach ($reply in $item.Replies) { $allReplies.Add(('[{0}] {1}' -f $reply.CreatedLocal, $reply.Text)) }
        $rows.Add([string[]]@(
                $item.Email, $item.DisplayName, $item.RespondedText, [string]$item.ReplyCount, $item.SentLocal, $item.FirstReplyLocal,
                $item.ResponseTimeDisplay, $minutes, $(if ($item.DirectReply) { 'Yes' } else { 'No' }), $item.Reactions,
                $item.FirstReplyText, [string]::Join([System.Environment]::NewLine, $allReplies), $item.Error
            ))
    }
    $header = @('Email', 'Display Name', 'Responded', 'Reply Count', 'Sent (local)', 'First Reply (local)', 'Response Time',
        'Response Time (minutes)', 'Quoted Reply', 'Reactions', 'First Reply', 'All Replies', 'Error')
    Write-TbmCsvFile -Path $Path -Rows $rows -Header $header
}

#endregion
