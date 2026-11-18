#region Interaction (reply) tracking
<#
    For each recipient a blast was sent to, find messages they posted in the blast's chat after the send:
      - Graph-sent blasts store ChatId/MessageId, so the chat is known.
      - Deep-link blasts store only SentUtc; the chat is found from /me/chats (members expanded).
    Messages are read with GET /chats/{id}/messages filtered to lastModifiedDateTime > SentUtc, then filtered
    client-side on createdDateTime > SentUtc, sender = recipient, messageType = message, not deleted.
    A reply that quotes the blast message (messageReference attachment) is flagged as a direct reply.
    Reactions on the blast message are read from the message itself.
#>

function Invoke-TbmInteractionWorker {
    <#
    .SYNOPSIS
        Background entry point: scans for replies, saves them into the blast's history and signals completion.
    .PARAMETER Job
        Hashtable: Blast (TbmBlastRecord), AuthState, Settings, Sync.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Job
    )

    $blast = [TeamsBlastMessenger.TbmBlastRecord]$Job.Blast
    $results = $null
    $errorMessage = $null
    try {
        $results = Invoke-TbmInteractionScan -Blast $blast -AuthState $Job.AuthState -Settings $Job.Settings -Sync $Job.Sync
        $blast.Interactions = $results
        $blast.LastInteractionCheckUtc = [datetime]::UtcNow
        Save-TbmBlastRecord -Blast $blast
    }
    catch [System.OperationCanceledException] {
        $errorMessage = 'Cancelled.'
        Write-TbmLog -Level Info -Source 'Interactions' -Message 'Interaction scan cancelled.'
    }
    catch {
        $errorMessage = Get-TbmFailureMessage -Exception $_.Exception
        Write-TbmLog -Source 'Interactions' -Message 'Interaction scan failed.' -ErrorRecord $_
    }
    finally {
        Add-TbmSyncEvent -Sync $Job.Sync -Type Done -Data @{ BlastId = $blast.Id; Results = $results; Error = $errorMessage }
    }
}

function Invoke-TbmInteractionScan {
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[TeamsBlastMessenger.TbmInteractionResult]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [TeamsBlastMessenger.TbmBlastRecord]$Blast,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Settings,

        [Parameter()]
        [hashtable]$Sync
    )

    $meId = [string]$AuthState.MeId
    if (-not [string]::IsNullOrEmpty($Blast.SenderId) -and -not [string]::Equals($Blast.SenderId, $meId, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw [System.InvalidOperationException]::new(('This blast was sent by {0}. Sign in as that account to check its replies.' -f $Blast.SenderUpn))
    }

    $results = [System.Collections.Generic.List[TeamsBlastMessenger.TbmInteractionResult]]::new()
    foreach ($recipient in $Blast.Recipients) {
        if ($recipient.Status -ne [TeamsBlastMessenger.TbmRecipientStatus]::Sent -or $null -eq $recipient.SentUtc) { continue }
        $result = [TeamsBlastMessenger.TbmInteractionResult]::new()
        $result.Email = $recipient.Email
        $result.DisplayName = $recipient.DisplayName
        $result.UserId = $recipient.UserId
        $result.ChatId = $recipient.ChatId
        $result.SentUtc = $recipient.SentUtc
        $results.Add($result)
    }
    Write-TbmLog -Level Info -Source 'Interactions' -Message ('Scanning replies for blast {0}: {1} sent recipient(s).' -f $Blast.Id, $results.Count)
    if ($results.Count -eq 0) { return , $results }

    # 1. Make sure every result has a user id (deep-link sends never resolved them).
    $userCache = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $step = 0
    foreach ($result in $results) {
        $step++
        if (-not [string]::IsNullOrEmpty($result.UserId)) { continue }
        Add-TbmSyncEvent -Sync $Sync -Type Progress -Data @{ Current = $step; Total = $results.Count; Text = ('Resolving {0}' -f $result.Email) }
        try {
            $user = Resolve-TbmGraphUser -AuthState $AuthState -Email $result.Email -Cache $userCache -Sync $Sync
            if ($null -eq $user) { $result.Error = 'Not found in the directory.'; continue }
            $result.UserId = $user.Id
            if ([string]::IsNullOrEmpty($result.DisplayName)) { $result.DisplayName = $user.DisplayName }
        }
        catch [System.OperationCanceledException] { throw }
        catch { $result.Error = Get-TbmFailureMessage -Exception $_.Exception }
    }

    # 2. Find chats for results that do not have one recorded.
    $needsChat = $false
    foreach ($result in $results) { if ([string]::IsNullOrEmpty($result.ChatId) -and -not [string]::IsNullOrEmpty($result.UserId)) { $needsChat = $true; break } }
    if ($needsChat) {
        Add-TbmSyncEvent -Sync $Sync -Type Progress -Data @{ Current = 0; Total = $results.Count; Text = 'Locating chats' }
        if ($Blast.Mode -eq 'Group') {
            $groupChatId = Find-TbmGroupChat -AuthState $AuthState -Blast $Blast -Results $results -Settings $Settings -Sync $Sync
            foreach ($result in $results) {
                if (-not [string]::IsNullOrEmpty($result.UserId)) {
                    if ([string]::IsNullOrEmpty($groupChatId)) { $result.Error = 'Group chat not found.' } else { $result.ChatId = $groupChatId }
                }
            }
        }
        else {
            $chatMap = Get-TbmOneOnOneChatMap -AuthState $AuthState -Settings $Settings -Sync $Sync
            foreach ($result in $results) {
                if (-not [string]::IsNullOrEmpty($result.ChatId) -or [string]::IsNullOrEmpty($result.UserId)) { continue }
                $chatId = $null
                if ($chatMap.TryGetValue($result.UserId, [ref]$chatId)) { $result.ChatId = $chatId }
                else { $result.Error = 'No 1:1 chat with this person was found.' }
            }
        }
    }

    # 3. Read each chat once and attribute messages.
    $byChat = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[TeamsBlastMessenger.TbmInteractionResult]]]::new([System.StringComparer]::Ordinal)
    foreach ($result in $results) {
        if ([string]::IsNullOrEmpty($result.ChatId)) { continue }
        if (-not $byChat.ContainsKey($result.ChatId)) { $byChat[$result.ChatId] = [System.Collections.Generic.List[TeamsBlastMessenger.TbmInteractionResult]]::new() }
        $byChat[$result.ChatId].Add($result)
    }

    $messageIdsByChat = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.HashSet[string]]]::new([System.StringComparer]::Ordinal)
    foreach ($recipient in $Blast.Recipients) {
        if ([string]::IsNullOrEmpty($recipient.ChatId) -or [string]::IsNullOrEmpty($recipient.MessageId)) { continue }
        if (-not $messageIdsByChat.ContainsKey($recipient.ChatId)) { $messageIdsByChat[$recipient.ChatId] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal) }
        $null = $messageIdsByChat[$recipient.ChatId].Add($recipient.MessageId)
    }

    $chatNumber = 0
    foreach ($pair in $byChat.GetEnumerator()) {
        $chatNumber++
        if ($null -ne $Sync -and [bool]$Sync['Cancel']) { throw [System.OperationCanceledException]::new('Cancelled.') }
        $chatId = $pair.Key
        $chatResults = $pair.Value
        Add-TbmSyncEvent -Sync $Sync -Type Progress -Data @{ Current = $chatNumber; Total = $byChat.Count; Text = ('Reading chat {0} of {1}' -f $chatNumber, $byChat.Count) }

        $since = [datetime]::MaxValue
        foreach ($result in $chatResults) { if ([datetime]$result.SentUtc -lt $since) { $since = [datetime]$result.SentUtc } }

        $blastMessageIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
        if ($messageIdsByChat.ContainsKey($chatId)) { $blastMessageIds = $messageIdsByChat[$chatId] }

        try {
            $messages = Get-TbmChatMessagesSince -AuthState $AuthState -ChatId $chatId -SinceUtc $since -Sync $Sync
            $reactions = Get-TbmBlastReactions -AuthState $AuthState -ChatId $chatId -MessageIds $blastMessageIds -Sync $Sync
            Set-TbmInteractionReplies -Results $chatResults -Messages $messages -MeId $meId -BlastMessageIds $blastMessageIds -Reactions $reactions -IsGroup ($Blast.Mode -eq 'Group')
        }
        catch [System.OperationCanceledException] { throw }
        catch {
            $reason = Get-TbmFailureMessage -Exception $_.Exception
            Write-TbmLog -Level Warn -Source 'Interactions' -Message ('Reading chat {0} failed: {1}' -f $chatId, $reason)
            foreach ($result in $chatResults) { $result.Error = $reason }
        }
    }

    $responded = 0
    foreach ($result in $results) { if ($result.Responded) { $responded++ } }
    Write-TbmLog -Level Info -Source 'Interactions' -Message ('Scan complete: {0} of {1} responded.' -f $responded, $results.Count)
    return , $results
}

function Set-TbmInteractionReplies {
    <#
    .SYNOPSIS
        Attributes chat messages to recipients and computes reply statistics. Pure function (no I/O) for testability.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Collections.Generic.List[TeamsBlastMessenger.TbmInteractionResult]]$Results,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$Messages,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$MeId,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$BlastMessageIds,

        [Parameter()]
        [System.Collections.Generic.Dictionary[string, string]]$Reactions,

        [Parameter()]
        [bool]$IsGroup = $false
    )

    foreach ($result in $Results) {
        $result.Replies.Clear()
        $result.Responded = $false
        $result.ReplyCount = 0
        $result.FirstReplyUtc = $null
        $result.LastReplyUtc = $null
        $result.FirstReplyText = ''
        $result.DirectReply = $false
        $result.Error = $null
        $reactionText = $null
        if ($null -ne $Reactions -and -not [string]::IsNullOrEmpty($result.UserId) -and $Reactions.TryGetValue($result.UserId, [ref]$reactionText)) {
            $result.Reactions = $reactionText
        }
        else {
            $result.Reactions = ''
        }
    }

    foreach ($message in $Messages) {
        if ([string](Get-TbmPropertyValue -InputObject $message -Name 'messageType') -ne 'message') { continue }
        if ($null -ne (Get-TbmPropertyValue -InputObject $message -Name 'deletedDateTime')) { continue }
        $fromId = [string](Get-TbmPropertyValue -InputObject $message -Name 'from.user.id')
        if ([string]::IsNullOrEmpty($fromId) -or [string]::Equals($fromId, $MeId, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        $createdUtc = ConvertTo-TbmUtcDateTime -Value (Get-TbmPropertyValue -InputObject $message -Name 'createdDateTime')
        if ($null -eq $createdUtc) { continue }

        $target = $null
        foreach ($result in $Results) {
            $sameUser = [string]::Equals($result.UserId, $fromId, [System.StringComparison]::OrdinalIgnoreCase)
            # In a 1:1 chat, anyone other than me is the recipient (covers unresolved ids).
            if ($sameUser -or (-not $IsGroup -and [string]::IsNullOrEmpty($result.UserId))) { $target = $result; break }
        }
        if ($null -eq $target -or $createdUtc -le [datetime]$target.SentUtc) { continue }

        $content = [string](Get-TbmPropertyValue -InputObject $message -Name 'body.content')
        $contentType = [string](Get-TbmPropertyValue -InputObject $message -Name 'body.contentType')
        $text = $content
        if ($contentType -eq 'html') { $text = [TeamsBlastMessenger.TbmText]::StripHtml($content) }

        $isDirect = $false
        $attachments = Get-TbmPropertyValue -InputObject $message -Name 'attachments'
        if ($null -ne $attachments) {
            foreach ($attachment in $attachments) {
                if ([string](Get-TbmPropertyValue -InputObject $attachment -Name 'contentType') -eq 'messageReference' -and
                    $BlastMessageIds.Contains([string](Get-TbmPropertyValue -InputObject $attachment -Name 'id'))) {
                    $isDirect = $true
                }
            }
        }
        if ([string]::IsNullOrWhiteSpace($text) -and $null -ne $attachments -and @($attachments).Count -gt 0) { $text = '[attachment]' }

        $reply = [TeamsBlastMessenger.TbmReplyMessage]::new()
        $reply.MessageId = [string](Get-TbmPropertyValue -InputObject $message -Name 'id')
        $reply.FromUserId = $fromId
        $reply.FromName = [string](Get-TbmPropertyValue -InputObject $message -Name 'from.user.displayName')
        $reply.CreatedUtc = $createdUtc
        $reply.Text = $text
        $reply.IsDirectReply = $isDirect
        $target.Replies.Add($reply)
        if ([string]::IsNullOrEmpty($target.UserId)) { $target.UserId = $fromId }
    }

    foreach ($result in $Results) {
        if ($result.Replies.Count -eq 0) { continue }
        $result.Replies.Sort([System.Comparison[TeamsBlastMessenger.TbmReplyMessage]] { param($left, $right) $left.CreatedUtc.CompareTo($right.CreatedUtc) })
        $first = $result.Replies[0]
        $result.Responded = $true
        $result.ReplyCount = $result.Replies.Count
        $result.FirstReplyUtc = $first.CreatedUtc
        $result.LastReplyUtc = $result.Replies[$result.Replies.Count - 1].CreatedUtc
        $result.FirstReplyText = $first.Text
        foreach ($reply in $result.Replies) { if ($reply.IsDirectReply) { $result.DirectReply = $true; break } }
        if ([string]::IsNullOrEmpty($result.DisplayName)) { $result.DisplayName = $first.FromName }
    }
}

function Get-TbmChatMessagesSince {
    <#
    .SYNOPSIS
        Returns chat messages created after SinceUtc, using the server-side lastModifiedDateTime filter when available.
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[object]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ChatId,

        [Parameter(Mandatory)]
        [datetime]$SinceUtc,

        [Parameter()]
        [hashtable]$Sync
    )

    $chatPath = '/chats/{0}/messages' -f [System.Uri]::EscapeDataString($ChatId)
    $since = ConvertTo-TbmIsoString -Value $SinceUtc
    $messages = $null
    try {
        $path = '{0}?$top=50&$orderby=lastModifiedDateTime desc&$filter=lastModifiedDateTime gt {1}' -f $chatPath, $since
        $messages = Invoke-TbmGraphPagedRequest -AuthState $AuthState -Path $path -MaxPages 40 -Sync $Sync
    }
    catch [TeamsBlastMessenger.TbmGraphException] {
        if ($_.Exception.StatusCode -ne 400) { throw }
        Write-TbmLog -Level Warn -Source 'Interactions' -Message 'Server-side message filter rejected; falling back to newest-first paging.'
        $path = '{0}?$top=50&$orderby=createdDateTime desc' -f $chatPath
        $stopWhen = {
            param($message)
            $created = ConvertTo-TbmUtcDateTime -Value (Get-TbmPropertyValue -InputObject $message -Name 'createdDateTime')
            # $SinceUtc resolves through the caller's scope (same module session state).
            return $null -ne $created -and $created -le $SinceUtc
        }
        $messages = Invoke-TbmGraphPagedRequest -AuthState $AuthState -Path $path -MaxPages 40 -Sync $Sync -StopWhen $stopWhen
    }

    $filtered = [System.Collections.Generic.List[object]]::new()
    foreach ($message in $messages) {
        $created = ConvertTo-TbmUtcDateTime -Value (Get-TbmPropertyValue -InputObject $message -Name 'createdDateTime')
        if ($null -ne $created -and $created -gt $SinceUtc) { $filtered.Add($message) }
    }
    return , $filtered
}

function Get-TbmBlastReactions {
    <#
    .SYNOPSIS
        Reads reactions on the blast message(s) in a chat. Returns userId -> "like, heart".
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.Dictionary[string, string]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ChatId,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$MessageIds,

        [Parameter()]
        [hashtable]$Sync
    )

    $byUser = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[string]]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($messageId in $MessageIds) {
        $path = '/chats/{0}/messages/{1}' -f [System.Uri]::EscapeDataString($ChatId), [System.Uri]::EscapeDataString($messageId)
        $message = Invoke-TbmGraphRequest -AuthState $AuthState -Path $path -AllowNotFound -Sync $Sync
        $reactions = Get-TbmPropertyValue -InputObject $message -Name 'reactions'
        if ($null -eq $reactions) { continue }
        foreach ($reaction in $reactions) {
            $userId = [string](Get-TbmPropertyValue -InputObject $reaction -Name 'user.user.id')
            if ([string]::IsNullOrEmpty($userId)) { continue }
            $name = [string](Get-TbmPropertyValue -InputObject $reaction -Name 'displayName')
            if ([string]::IsNullOrEmpty($name)) { $name = [string](Get-TbmPropertyValue -InputObject $reaction -Name 'reactionType') }
            if (-not $byUser.ContainsKey($userId)) { $byUser[$userId] = [System.Collections.Generic.List[string]]::new() }
            if (-not $byUser[$userId].Contains($name)) { $byUser[$userId].Add($name) }
        }
    }
    $result = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($pair in $byUser.GetEnumerator()) { $result[$pair.Key] = [string]::Join(', ', $pair.Value) }
    return , $result
}

function Get-TbmMyChats {
    <#
    .SYNOPSIS
        Lists the signed-in user's chats of a type with members expanded, most recently active first.
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[object]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [ValidateSet('oneOnOne', 'group')]
        [string]$ChatType,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Settings,

        [Parameter()]
        [hashtable]$Sync
    )

    $maxPages = [System.Math]::Max(1, [int]$Settings.InteractionMaxChatPages)
    $chats = $null
    try {
        $path = '/me/chats?$expand=members&$top=50&$orderby=lastMessagePreview/createdDateTime desc&$filter=chatType eq ''{0}''' -f $ChatType
        $chats = Invoke-TbmGraphPagedRequest -AuthState $AuthState -Path $path -MaxPages $maxPages -Sync $Sync
    }
    catch [TeamsBlastMessenger.TbmGraphException] {
        if ($_.Exception.StatusCode -ne 400) { throw }
        Write-TbmLog -Level Warn -Source 'Interactions' -Message 'Chat list filter rejected; listing all chats and filtering locally.'
        $chats = Invoke-TbmGraphPagedRequest -AuthState $AuthState -Path '/me/chats?$expand=members&$top=50' -MaxPages $maxPages -Sync $Sync
    }

    $filtered = [System.Collections.Generic.List[object]]::new()
    foreach ($chat in $chats) {
        if ([string](Get-TbmPropertyValue -InputObject $chat -Name 'chatType') -eq $ChatType) { $filtered.Add($chat) }
    }
    return , $filtered
}

function Get-TbmOneOnOneChatMap {
    <#
    .SYNOPSIS
        Maps the other participant's user id to the 1:1 chat id.
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.Dictionary[string, string]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Settings,

        [Parameter()]
        [hashtable]$Sync
    )

    $map = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($chat in (Get-TbmMyChats -AuthState $AuthState -ChatType oneOnOne -Settings $Settings -Sync $Sync)) {
        $members = Get-TbmPropertyValue -InputObject $chat -Name 'members'
        if ($null -eq $members) { continue }
        foreach ($member in $members) {
            $userId = [string](Get-TbmPropertyValue -InputObject $member -Name 'userId')
            if ([string]::IsNullOrEmpty($userId) -or [string]::Equals($userId, [string]$AuthState.MeId, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            if (-not $map.ContainsKey($userId)) { $map[$userId] = [string]$chat.id }
        }
    }
    Write-TbmLog -Level Info -Source 'Interactions' -Message ('Indexed {0} 1:1 chat(s).' -f $map.Count)
    return , $map
}

function Find-TbmGroupChat {
    <#
    .SYNOPSIS
        Finds the group chat whose members are exactly me + the blast's recipients (topic preferred), most recent first.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [TeamsBlastMessenger.TbmBlastRecord]$Blast,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Collections.Generic.List[TeamsBlastMessenger.TbmInteractionResult]]$Results,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Settings,

        [Parameter()]
        [hashtable]$Sync
    )

    $expected = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $null = $expected.Add([string]$AuthState.MeId)
    foreach ($result in $Results) { if (-not [string]::IsNullOrEmpty($result.UserId)) { $null = $expected.Add($result.UserId) } }

    $bestId = $null
    $bestScore = -1
    foreach ($chat in (Get-TbmMyChats -AuthState $AuthState -ChatType group -Settings $Settings -Sync $Sync)) {
        $memberIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $members = Get-TbmPropertyValue -InputObject $chat -Name 'members'
        if ($null -ne $members) { foreach ($member in $members) { $null = $memberIds.Add([string](Get-TbmPropertyValue -InputObject $member -Name 'userId')) } }
        if (-not $memberIds.SetEquals($expected)) { continue }

        $score = 1
        $topic = [string](Get-TbmPropertyValue -InputObject $chat -Name 'topic')
        if (-not [string]::IsNullOrWhiteSpace($Blast.Topic) -and [string]::Equals($topic.Trim(), $Blast.Topic.Trim(), [System.StringComparison]::OrdinalIgnoreCase)) { $score = 2 }
        if ($score -gt $bestScore) { $bestScore = $score; $bestId = [string]$chat.id }
    }
    if ($null -eq $bestId) { Write-TbmLog -Level Warn -Source 'Interactions' -Message 'No group chat with the same members was found.' }
    return $bestId
}

function Get-TbmInteractionSummary {
    <#
    .SYNOPSIS
        Response rate and average/median response time for display.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[TeamsBlastMessenger.TbmInteractionResult]]$Results
    )

    if ($Results.Count -eq 0) { return 'No sent recipients to check.' }
    $times = [System.Collections.Generic.List[System.TimeSpan]]::new()
    $responded = 0
    $reacted = 0
    foreach ($result in $Results) {
        if ($result.Responded) { $responded++; $times.Add([System.TimeSpan]$result.ResponseTime) }
        if (-not [string]::IsNullOrEmpty($result.Reactions)) { $reacted++ }
    }
    $rate = [double]$responded / $Results.Count
    $sentences = [System.Collections.Generic.List[string]]::new()
    $sentences.Add(('{0} of {1} responded ({2:P0}). {3} reacted.' -f $responded, $Results.Count, $rate, $reacted))
    if ($times.Count -gt 0) {
        $sentences.Add(('Average response {0}, median {1}.' -f [TeamsBlastMessenger.TbmText]::FormatDuration([TeamsBlastMessenger.TbmText]::Average($times)), [TeamsBlastMessenger.TbmText]::FormatDuration([TeamsBlastMessenger.TbmText]::Median($times))))
    }
    return [string]::Join(' ', $sentences)
}

#endregion
