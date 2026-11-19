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
    $unresolved = [System.Collections.Generic.List[string]]::new()
    foreach ($result in $results) { if ([string]::IsNullOrEmpty($result.UserId)) { $unresolved.Add($result.Email) } }
    if ($unresolved.Count -gt 1) {
        Add-TbmSyncEvent -Sync $Sync -Type Progress -Data @{ Current = 0; Total = $results.Count; Text = ('Resolving {0} recipient(s)' -f $unresolved.Count) }
        Resolve-TbmGraphUsers -AuthState $AuthState -Emails $unresolved.ToArray() -Cache $userCache -Sync $Sync
    }
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

    $chats = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($pair in $byChat.GetEnumerator()) {
        $since = [datetime]::MaxValue
        foreach ($result in $pair.Value) { if ([datetime]$result.SentUtc -lt $since) { $since = [datetime]$result.SentUtc } }
        $chats.Add(@{ ChatId = $pair.Key; SinceUtc = $since })
    }
    Add-TbmSyncEvent -Sync $Sync -Type Progress -Data @{ Current = 0; Total = $chats.Count; Text = ('Reading {0} chat(s)' -f $chats.Count) }
    $messagesByChat = Get-TbmChatMessagesBatch -AuthState $AuthState -Chats $chats -MaxPages 40 -Sync $Sync
    $reactionsByChat = Get-TbmBlastReactionsBatch -AuthState $AuthState -MessageIdsByChat $messageIdsByChat -Sync $Sync

    foreach ($pair in $byChat.GetEnumerator()) {
        $chatId = $pair.Key
        $chatResults = $pair.Value
        $read = $messagesByChat[$chatId]
        $reactionInfo = $null
        $reactions = $null
        $reason = $read.Error
        if ($reactionsByChat.TryGetValue($chatId, [ref]$reactionInfo)) {
            $reactions = $reactionInfo.Reactions
            if ($null -eq $reason) { $reason = $reactionInfo.Error }
        }
        if ($null -ne $reason) {
            Write-TbmLog -Level Warn -Source 'Interactions' -Message ('Reading chat {0} failed: {1}' -f $chatId, $reason)
            foreach ($result in $chatResults) { $result.Error = $reason }
            continue
        }
        $blastMessageIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
        if ($messageIdsByChat.ContainsKey($chatId)) { $blastMessageIds = $messageIdsByChat[$chatId] }
        Set-TbmInteractionReplies -Results $chatResults -Messages $read.Messages -MeId $meId -BlastMessageIds $blastMessageIds -Reactions $reactions -IsGroup ($Blast.Mode -eq 'Group')
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
        if ([string]([TeamsBlastMessenger.TbmObject]::GetValue($message, 'messageType')) -ne 'message') { continue }
        if ($null -ne ([TeamsBlastMessenger.TbmObject]::GetValue($message, 'deletedDateTime'))) { continue }
        $fromId = [string]([TeamsBlastMessenger.TbmObject]::GetValue($message, 'from.user.id'))
        if ([string]::IsNullOrEmpty($fromId) -or [string]::Equals($fromId, $MeId, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        $createdUtc = [TeamsBlastMessenger.TbmObject]::ToUtc([TeamsBlastMessenger.TbmObject]::GetValue($message, 'createdDateTime'))
        if ($null -eq $createdUtc) { continue }

        $target = $null
        foreach ($result in $Results) {
            $sameUser = [string]::Equals($result.UserId, $fromId, [System.StringComparison]::OrdinalIgnoreCase)
            # In a 1:1 chat, anyone other than me is the recipient (covers unresolved ids).
            if ($sameUser -or (-not $IsGroup -and [string]::IsNullOrEmpty($result.UserId))) { $target = $result; break }
        }
        if ($null -eq $target -or $createdUtc -le [datetime]$target.SentUtc) { continue }

        $content = [string]([TeamsBlastMessenger.TbmObject]::GetValue($message, 'body.content'))
        $contentType = [string]([TeamsBlastMessenger.TbmObject]::GetValue($message, 'body.contentType'))
        $text = $content
        if ($contentType -eq 'html') { $text = [TeamsBlastMessenger.TbmText]::StripHtml($content) }

        $isDirect = $false
        $attachments = [TeamsBlastMessenger.TbmObject]::GetValue($message, 'attachments')
        if ($null -ne $attachments) {
            foreach ($attachment in $attachments) {
                if ([string]([TeamsBlastMessenger.TbmObject]::GetValue($attachment, 'contentType')) -eq 'messageReference' -and
                    $BlastMessageIds.Contains([string]([TeamsBlastMessenger.TbmObject]::GetValue($attachment, 'id')))) {
                    $isDirect = $true
                }
            }
        }
        if ([string]::IsNullOrWhiteSpace($text) -and $null -ne $attachments -and @($attachments).Count -gt 0) { $text = '[attachment]' }

        $reply = [TeamsBlastMessenger.TbmReplyMessage]::new()
        $reply.MessageId = [string]([TeamsBlastMessenger.TbmObject]::GetValue($message, 'id'))
        $reply.FromUserId = $fromId
        $reply.FromName = [string]([TeamsBlastMessenger.TbmObject]::GetValue($message, 'from.user.displayName'))
        $reply.CreatedUtc = $createdUtc
        $reply.Text = $text
        $reply.IsDirectReply = $isDirect
        $target.Replies.Add($reply)
        if ([string]::IsNullOrEmpty($target.UserId)) { $target.UserId = $fromId }
    }

    $replyOrder = [TeamsBlastMessenger.TbmReplyByCreated]::new()
    foreach ($result in $Results) {
        if ($result.Replies.Count -eq 0) { continue }
        $result.Replies.Sort($replyOrder)
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

function Get-TbmChatMessagePath {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ChatId,

        [Parameter(Mandatory)]
        [datetime]$SinceUtc,

        [Parameter(Mandatory)]
        [bool]$ServerFilter
    )
    $chatPath = '/chats/{0}/messages' -f [System.Uri]::EscapeDataString($ChatId)
    if ($ServerFilter) { return '{0}?$top=50&$orderby=lastModifiedDateTime desc&$filter=lastModifiedDateTime gt {1}' -f $chatPath, (ConvertTo-TbmIsoString -Value $SinceUtc) }
    return '{0}?$top=50&$orderby=createdDateTime desc' -f $chatPath
}

function Get-TbmChatMessagesBatch {
    <#
    .SYNOPSIS
        Reads messages created after each chat's SinceUtc, for many chats at once, through Graph $batch.
    .DESCRIPTION
        Every round sends the next page of every unfinished chat in one set of batch calls, following
        @odata.nextLink until a chat's last page or MaxPages. The server-side lastModifiedDateTime filter is
        tried on the first chat alone; if Graph rejects it (400), every chat uses newest-first paging instead
        and stops at the first message older than its SinceUtc, so the failing request is made only once.
    .PARAMETER Chats
        List of hashtables: ChatId, SinceUtc.
    .OUTPUTS
        Dictionary of ChatId -> hashtable with Messages (List[object], created after SinceUtc) and Error (string or $null).
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.Dictionary[string, hashtable]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[hashtable]]$Chats,

        [Parameter()]
        [ValidateRange(1, 1000)]
        [int]$MaxPages = 40,

        [Parameter()]
        [hashtable]$Sync
    )

    $states = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($chat in $Chats) {
        $states.Add(@{ ChatId = [string]$chat.ChatId; Since = [datetime]$chat.SinceUtc; Next = $null; Filtered = $true; Pages = 0; Items = [System.Collections.Generic.List[object]]::new(); Error = $null; Done = $false })
    }

    $useServerFilter = $true
    $probed = $false
    while ($true) {
        $batch = [System.Collections.Generic.List[hashtable]]::new()
        $requests = [System.Collections.Generic.List[hashtable]]::new()
        for ($i = 0; $i -lt $states.Count; $i++) {
            $state = $states[$i]
            if ($state.Done) { continue }
            if ($null -eq $state.Next) {
                if (-not $probed -and $batch.Count -gt 0) { continue }
                $state.Filtered = $useServerFilter
                $state.Next = Get-TbmChatMessagePath -ChatId $state.ChatId -SinceUtc $state.Since -ServerFilter $useServerFilter
            }
            $batch.Add($state)
            $requests.Add(@{ Id = [string]$i; Path = [string]$state.Next })
            if (-not $probed) { break }
        }
        if ($requests.Count -eq 0) { break }

        $answers = Invoke-TbmGraphBatch -AuthState $AuthState -Requests $requests -Sync $Sync
        foreach ($request in $requests) {
            $state = $states[[int]$request.Id]
            $answer = $null
            if (-not $answers.TryGetValue($request.Id, [ref]$answer)) { $state.Error = 'No response from Microsoft Graph.'; $state.Done = $true; continue }

            if ($answer.Status -eq 400 -and $state.Filtered -and $state.Pages -eq 0) {
                if ($useServerFilter) { Write-TbmLog -Level Warn -Source 'Interactions' -Message 'Server-side message filter rejected; using newest-first paging for every chat.' }
                $useServerFilter = $false
                $state.Filtered = $false
                $state.Next = Get-TbmChatMessagePath -ChatId $state.ChatId -SinceUtc $state.Since -ServerFilter $false
                continue
            }
            if ($answer.Status -ne 200) {
                $state.Error = Get-TbmFailureMessage -Exception (New-TbmBatchItemException -Result $answer -Path $request.Path)
                $state.Done = $true
                continue
            }

            $state.Pages++
            $reachedOlder = $false
            $values = [TeamsBlastMessenger.TbmObject]::GetValue($answer.Body, 'value')
            if ($null -ne $values) {
                foreach ($message in $values) {
                    $state.Items.Add($message)
                    if (-not $state.Filtered) {
                        $created = [TeamsBlastMessenger.TbmObject]::ToUtc([TeamsBlastMessenger.TbmObject]::GetValue($message, 'createdDateTime'))
                        if ($null -ne $created -and $created -le $state.Since) { $reachedOlder = $true }
                    }
                }
            }
            $next = [string]([TeamsBlastMessenger.TbmObject]::GetValue($answer.Body, '@odata.nextLink'))
            if ($reachedOlder -or [string]::IsNullOrEmpty($next)) { $state.Done = $true; continue }
            if ($state.Pages -ge $MaxPages) {
                Write-TbmLog -Level Warn -Source 'Interactions' -Message ('Stopped reading chat {0} after {1} page(s); replies may be incomplete.' -f $state.ChatId, $MaxPages)
                $state.Done = $true
                continue
            }
            $state.Next = $next
        }
        $probed = $true

        $finished = 0
        foreach ($state in $states) { if ($state.Done) { $finished++ } }
        if ($null -ne $Sync) { Add-TbmSyncEvent -Sync $Sync -Type Progress -Data @{ Current = $finished; Total = $states.Count; Text = ('Read {0} of {1} chat(s)' -f $finished, $states.Count) } }
    }

    $output = [System.Collections.Generic.Dictionary[string, hashtable]]::new([System.StringComparer]::Ordinal)
    foreach ($state in $states) {
        $messages = [System.Collections.Generic.List[object]]::new()
        foreach ($message in $state.Items) {
            $created = [TeamsBlastMessenger.TbmObject]::ToUtc([TeamsBlastMessenger.TbmObject]::GetValue($message, 'createdDateTime'))
            if ($null -ne $created -and $created -gt $state.Since) { $messages.Add($message) }
        }
        $output[$state.ChatId] = @{ Messages = $messages; Error = $state.Error }
    }
    return , $output
}

function Get-TbmBlastReactionsBatch {
    <#
    .SYNOPSIS
        Reads reactions on the blast message(s) of many chats through Graph $batch.
    .OUTPUTS
        Dictionary of ChatId -> hashtable with Reactions (Dictionary userId -> "like, heart") and Error (string or $null).
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.Dictionary[string, hashtable]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Collections.Generic.Dictionary[string, System.Collections.Generic.HashSet[string]]]$MessageIdsByChat,

        [Parameter()]
        [hashtable]$Sync
    )

    $byUserByChat = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[string]]]]::new([System.StringComparer]::Ordinal)
    $errors = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::Ordinal)
    $requests = [System.Collections.Generic.List[hashtable]]::new()
    $chatOfRequest = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::Ordinal)
    foreach ($pair in $MessageIdsByChat.GetEnumerator()) {
        $byUserByChat[$pair.Key] = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[string]]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($messageId in $pair.Value) {
            $id = [string]$requests.Count
            $requests.Add(@{ Id = $id; Path = ('/chats/{0}/messages/{1}' -f [System.Uri]::EscapeDataString($pair.Key), [System.Uri]::EscapeDataString($messageId)) })
            $chatOfRequest[$id] = $pair.Key
        }
    }

    if ($requests.Count -gt 0) {
        $answers = Invoke-TbmGraphBatch -AuthState $AuthState -Requests $requests -Sync $Sync
        foreach ($request in $requests) {
            $chatId = $chatOfRequest[$request.Id]
            $answer = $null
            if (-not $answers.TryGetValue($request.Id, [ref]$answer) -or $answer.Status -eq 404) { continue }
            if ($answer.Status -ne 200) {
                $errors[$chatId] = Get-TbmFailureMessage -Exception (New-TbmBatchItemException -Result $answer -Path $request.Path)
                continue
            }
            $reactions = [TeamsBlastMessenger.TbmObject]::GetValue($answer.Body, 'reactions')
            if ($null -eq $reactions) { continue }
            $byUser = $byUserByChat[$chatId]
            foreach ($reaction in $reactions) {
                $userId = [string]([TeamsBlastMessenger.TbmObject]::GetValue($reaction, 'user.user.id'))
                if ([string]::IsNullOrEmpty($userId)) { continue }
                $name = [string]([TeamsBlastMessenger.TbmObject]::GetValue($reaction, 'displayName'))
                if ([string]::IsNullOrEmpty($name)) { $name = [string]([TeamsBlastMessenger.TbmObject]::GetValue($reaction, 'reactionType')) }
                if (-not $byUser.ContainsKey($userId)) { $byUser[$userId] = [System.Collections.Generic.List[string]]::new() }
                if (-not $byUser[$userId].Contains($name)) { $byUser[$userId].Add($name) }
            }
        }
    }

    $output = [System.Collections.Generic.Dictionary[string, hashtable]]::new([System.StringComparer]::Ordinal)
    foreach ($pair in $byUserByChat.GetEnumerator()) {
        $joined = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($user in $pair.Value.GetEnumerator()) { $joined[$user.Key] = [string]::Join(', ', $user.Value) }
        $reason = $null
        $null = $errors.TryGetValue($pair.Key, [ref]$reason)
        $output[$pair.Key] = @{ Reactions = $joined; Error = $reason }
    }
    return , $output
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
        if ([string]([TeamsBlastMessenger.TbmObject]::GetValue($chat, 'chatType')) -eq $ChatType) { $filtered.Add($chat) }
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
        $members = [TeamsBlastMessenger.TbmObject]::GetValue($chat, 'members')
        if ($null -eq $members) { continue }
        foreach ($member in $members) {
            $userId = [string]([TeamsBlastMessenger.TbmObject]::GetValue($member, 'userId'))
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
        $members = [TeamsBlastMessenger.TbmObject]::GetValue($chat, 'members')
        if ($null -ne $members) { foreach ($member in $members) { $null = $memberIds.Add([string]([TeamsBlastMessenger.TbmObject]::GetValue($member, 'userId'))) } }
        if (-not $memberIds.SetEquals($expected)) { continue }

        $score = 1
        $topic = [string]([TeamsBlastMessenger.TbmObject]::GetValue($chat, 'topic'))
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
