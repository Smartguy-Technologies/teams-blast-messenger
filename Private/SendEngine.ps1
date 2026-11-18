#region Background jobs
<#
    Long-running work (sending, interaction scans) runs in a dedicated runspace that imports this module.
    Workers never touch WPF objects: they enqueue events on $Sync.Events (a ConcurrentQueue) which the UI
    drains on a DispatcherTimer. Cancellation is a flag on the same synchronized hashtable.
#>

function New-TbmSyncState {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()
    return [hashtable]::Synchronized(@{
            Cancel = $false
            Events = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
        })
}

function Add-TbmSyncEvent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Sync,

        [Parameter(Mandatory)]
        [ValidateSet('Recipient', 'Progress', 'Done')]
        [string]$Type,

        [Parameter()]
        [hashtable]$Data
    )

    $syncEvent = @{ Type = $Type }
    if ($null -ne $Data) { foreach ($key in $Data.Keys) { $syncEvent[$key] = $Data[$key] } }
    $Sync.Events.Enqueue([pscustomobject]$syncEvent)
}

function Start-TbmBackgroundJob {
    <#
    .SYNOPSIS
        Runs an exported module command in a new STA runspace. Returns a handle for Complete-TbmBackgroundJob.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Invoke-TbmSendWorker', 'Invoke-TbmInteractionWorker')]
        [string]$CommandName,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Job
    )

    $sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
    if (Test-TbmIsWindows) {
        try { $sessionState.ExecutionPolicy = [Microsoft.PowerShell.ExecutionPolicy]::Bypass }
        catch { Write-TbmLog -Level Debug -Source 'Jobs' -Message 'Could not set runspace execution policy.' }
    }
    $sessionState.ImportPSModule([string[]]@($script:TbmManifestPath))

    $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($sessionState)
    if (Test-TbmIsWindows) { $runspace.ApartmentState = [System.Threading.ApartmentState]::STA }
    $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $runspace.Open()

    $powershell = [System.Management.Automation.PowerShell]::Create()
    $powershell.Runspace = $runspace
    $null = $powershell.AddCommand($CommandName).AddParameter('Job', $Job)
    $asyncResult = $powershell.BeginInvoke()
    Write-TbmLog -Level Debug -Source 'Jobs' -Message ('Started {0}.' -f $CommandName)
    return @{ PowerShell = $powershell; Runspace = $runspace; AsyncResult = $asyncResult; CommandName = $CommandName }
}

function Complete-TbmBackgroundJob {
    <#
    .SYNOPSIS
        Ends a finished job, logs any errors from its streams and disposes it.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Handle
    )

    try {
        $null = $Handle.PowerShell.EndInvoke($Handle.AsyncResult)
        foreach ($record in $Handle.PowerShell.Streams.Error) {
            Write-TbmLog -Source 'Jobs' -Message ('{0} reported an error.' -f $Handle.CommandName) -ErrorRecord $record
        }
    }
    catch {
        Write-TbmLog -Source 'Jobs' -Message ('{0} failed.' -f $Handle.CommandName) -ErrorRecord $_
    }
    finally {
        $Handle.PowerShell.Dispose()
        $Handle.Runspace.Dispose()
    }
}

#endregion

#region Send worker

function Invoke-TbmSendWorker {
    <#
    .SYNOPSIS
        Background entry point: sends a blast and records every outcome to history.
    .PARAMETER Job
        Hashtable: Blast (TbmBlastRecord), Items (List[hashtable] with Index, Email, Message),
        Settings (hashtable), AuthState (synchronized hashtable), Sync (synchronized hashtable).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Job
    )

    $blast = [TeamsBlastMessenger.TbmBlastRecord]$Job.Blast
    $sync = $Job.Sync
    Write-TbmLog -Level Info -Source 'Send' -Message ('Blast {0} started: {1} recipient(s), mode {2}, transport {3}{4}.' -f $blast.Id, $Job.Items.Count, $blast.Mode, $blast.Transport, $(if ($blast.IsTest) { ', TEST' } else { '' }))
    try {
        Save-TbmBlastRecord -Blast $blast
        if ($blast.Mode -eq 'Group') { Invoke-TbmGroupSend -Job $Job }
        else { Invoke-TbmIndividualSend -Job $Job }
    }
    catch {
        Write-TbmLog -Source 'Send' -Message ('Blast {0} stopped by an unexpected error.' -f $blast.Id) -ErrorRecord $_
        foreach ($item in $Job.Items) {
            $recipient = $blast.Recipients[$item.Index]
            if ($recipient.Status -eq [TeamsBlastMessenger.TbmRecipientStatus]::Pending) {
                Set-TbmRecipientOutcome -Job $Job -Index $item.Index -Status Failed -ErrorMessage $_.Exception.Message
            }
        }
    }
    finally {
        foreach ($item in $Job.Items) {
            $recipient = $blast.Recipients[$item.Index]
            if ($recipient.Status -eq [TeamsBlastMessenger.TbmRecipientStatus]::Pending) {
                Set-TbmRecipientOutcome -Job $Job -Index $item.Index -Status Cancelled -ErrorMessage 'Cancelled before sending.'
            }
        }
        $blast.CompletedUtc = [datetime]::UtcNow
        try { Save-TbmBlastRecord -Blast $blast }
        catch { Write-TbmLog -Source 'Send' -Message 'Could not save final history record.' -ErrorRecord $_ }
        Write-TbmLog -Level Info -Source 'Send' -Message ('Blast {0} finished: {1} sent, {2} failed, {3} total.' -f $blast.Id, $blast.SentCount, $blast.FailedCount, $blast.RecipientCount)
        Add-TbmSyncEvent -Sync $sync -Type Done -Data @{ BlastId = $blast.Id; Sent = $blast.SentCount; Failed = $blast.FailedCount; Total = $blast.RecipientCount }
    }
}

function Set-TbmRecipientOutcome {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Job,

        [Parameter(Mandatory)]
        [int]$Index,

        [Parameter(Mandatory)]
        [TeamsBlastMessenger.TbmRecipientStatus]$Status,

        [Parameter()]
        [string]$ErrorMessage
    )

    $blast = [TeamsBlastMessenger.TbmBlastRecord]$Job.Blast
    $recipient = $blast.Recipients[$Index]
    $recipient.Status = $Status
    $recipient.Error = $ErrorMessage
    try { Add-TbmBlastJournalEntry -Blast $blast -Index $Index }
    catch { Write-TbmLog -Level Warn -Source 'History' -Message ('Journal write failed: {0}' -f $_.Exception.Message) }

    $level = [TeamsBlastMessenger.TbmLogLevel]::Info
    if ($Status -eq [TeamsBlastMessenger.TbmRecipientStatus]::Failed) { $level = [TeamsBlastMessenger.TbmLogLevel]::Warn }
    $suffix = ''
    if (-not [string]::IsNullOrEmpty($ErrorMessage)) { $suffix = ': ' + $ErrorMessage }
    Write-TbmLog -Level $level -Source 'Send' -Message ('{0} -> {1}{2}' -f $recipient.Email, $Status, $suffix)

    Add-TbmSyncEvent -Sync $Job.Sync -Type Recipient -Data @{
        Index     = $Index
        RowIndex  = $recipient.RowIndex
        Email     = $recipient.Email
        Status    = $Status.ToString()
        Error     = $ErrorMessage
        SentLocal = $recipient.SentLocal
    }
}

function Get-TbmFailureMessage {
    <#
    .SYNOPSIS
        Produces a short, user-facing reason from an exception (Graph errors keep their code).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Exception]$Exception
    )

    $graphError = $Exception -as [TeamsBlastMessenger.TbmGraphException]
    if ($null -eq $graphError -and $null -ne $Exception.InnerException) { $graphError = $Exception.InnerException -as [TeamsBlastMessenger.TbmGraphException] }
    if ($null -ne $graphError) {
        switch ($graphError.StatusCode) {
            403 { return ('Access denied ({0}). The recipient may be external/blocked, or a permission is missing.' -f $graphError.ErrorCode) }
            404 { return ('Not found ({0}).' -f $graphError.ErrorCode) }
        }
        return ('Graph {0} {1}' -f $graphError.StatusCode, $graphError.ErrorCode)
    }
    return $Exception.GetBaseException().Message
}

function Send-TbmBlastItem {
    <#
    .SYNOPSIS
        Delivers one message to one or more recipients over the blast's transport and fills in the results.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Job,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [int[]]$Indexes,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Message,

        [Parameter()]
        [System.Collections.Generic.Dictionary[string, object]]$UserCache
    )

    $blast = [TeamsBlastMessenger.TbmBlastRecord]$Job.Blast
    $settings = $Job.Settings
    $sync = $Job.Sync
    $format = 'Text'
    if ($blast.Format -eq 'Html') { $format = 'Html' }

    if ($blast.Transport -like 'Graph*') {
        $userIds = [System.Collections.Generic.List[string]]::new()
        foreach ($index in $Indexes) {
            $recipient = $blast.Recipients[$index]
            if ([string]::IsNullOrEmpty($recipient.UserId)) {
                $user = Resolve-TbmGraphUser -AuthState $Job.AuthState -Email $recipient.Email -Cache $UserCache -Sync $sync
                if ($null -eq $user) { throw [System.InvalidOperationException]::new(('"{0}" was not found in the directory.' -f $recipient.Email)) }
                $recipient.UserId = $user.Id
                $recipient.DisplayName = $user.DisplayName
            }
            $userIds.Add($recipient.UserId)
        }
        $chatId = New-TbmGraphChat -AuthState $Job.AuthState -MemberUserIds $userIds.ToArray() -Topic $blast.Topic -Sync $sync
        $html = ConvertTo-TbmGraphHtml -Message $Message -Format $format
        $sent = Send-TbmGraphChatMessage -AuthState $Job.AuthState -ChatId $chatId -HtmlContent $html -Sync $sync
        $sentUtc = $sent.CreatedUtc
        if ($null -eq $sentUtc) { $sentUtc = [datetime]::UtcNow }
        foreach ($index in $Indexes) {
            $recipient = $blast.Recipients[$index]
            $recipient.ChatId = $chatId
            $recipient.MessageId = $sent.Id
            $recipient.SentUtc = $sentUtc
        }
        return
    }

    # Deep link (auto mode; manual mode is driven by the UI thread).
    $emails = [System.Collections.Generic.List[string]]::new()
    foreach ($index in $Indexes) { $emails.Add($blast.Recipients[$index].Email) }
    $plain = ConvertTo-TbmPlainText -Message $Message -Format $format
    $uri = New-TbmDeepLink -Emails $emails.ToArray() -Message $plain -Topic $blast.Topic -Target ([string]$settings.DeepLinkTarget)
    Start-TbmDeepLink -Uri $uri
    if (-not (Wait-TbmDelay -Milliseconds ([int]$settings.DeepLinkSendDelaySeconds * 1000) -Sync $sync)) {
        throw [System.OperationCanceledException]::new('Cancelled before Enter was sent; the chat may still contain the unsent text.')
    }
    Send-TbmTeamsEnterKey
    $sentUtc = [datetime]::UtcNow
    foreach ($index in $Indexes) { $blast.Recipients[$index].SentUtc = $sentUtc }
}

function Invoke-TbmIndividualSend {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Job
    )

    $sync = $Job.Sync
    $userCache = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $pacing = [int]$Job.Settings.PacingDelayMs
    if ($Job.Blast.Transport -like 'DeepLink*') { $pacing = [System.Math]::Max($pacing, 1500) }

    $position = 0
    foreach ($item in $Job.Items) {
        $position++
        if ([bool]$sync['Cancel']) { break }
        try {
            Send-TbmBlastItem -Job $Job -Indexes @([int]$item.Index) -Message ([string]$item.Message) -UserCache $userCache
            Set-TbmRecipientOutcome -Job $Job -Index $item.Index -Status Sent
        }
        catch [System.OperationCanceledException] {
            Set-TbmRecipientOutcome -Job $Job -Index $item.Index -Status Cancelled -ErrorMessage $_.Exception.Message
            break
        }
        catch {
            Write-TbmLog -Source 'Send' -Message ('Send to {0} failed.' -f $item.Email) -ErrorRecord $_
            Set-TbmRecipientOutcome -Job $Job -Index $item.Index -Status Failed -ErrorMessage (Get-TbmFailureMessage -Exception $_.Exception)
        }
        if ($position -lt $Job.Items.Count -and -not (Wait-TbmDelay -Milliseconds $pacing -Sync $sync)) { break }
    }
}

function Invoke-TbmGroupSend {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Job
    )

    $blast = [TeamsBlastMessenger.TbmBlastRecord]$Job.Blast
    $userCache = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $indexes = [System.Collections.Generic.List[int]]::new()

    # Resolve members first so one unknown address does not block the whole group.
    foreach ($item in $Job.Items) {
        if ([bool]$Job.Sync['Cancel']) { return }
        $recipient = $blast.Recipients[$item.Index]
        if ($blast.Transport -like 'Graph*') {
            try {
                $user = Resolve-TbmGraphUser -AuthState $Job.AuthState -Email $recipient.Email -Cache $userCache -Sync $Job.Sync
                if ($null -eq $user) {
                    Set-TbmRecipientOutcome -Job $Job -Index $item.Index -Status Failed -ErrorMessage 'Not found in the directory.'
                    continue
                }
                $recipient.UserId = $user.Id
                $recipient.DisplayName = $user.DisplayName
            }
            catch {
                Set-TbmRecipientOutcome -Job $Job -Index $item.Index -Status Failed -ErrorMessage (Get-TbmFailureMessage -Exception $_.Exception)
                continue
            }
        }
        $indexes.Add([int]$item.Index)
    }
    if ($indexes.Count -eq 0) { return }

    try {
        Send-TbmBlastItem -Job $Job -Indexes $indexes.ToArray() -Message ([string]$Job.Items[0].Message) -UserCache $userCache
        foreach ($index in $indexes) { Set-TbmRecipientOutcome -Job $Job -Index $index -Status Sent }
    }
    catch {
        Write-TbmLog -Source 'Send' -Message 'Group send failed.' -ErrorRecord $_
        $reason = Get-TbmFailureMessage -Exception $_.Exception
        $status = [TeamsBlastMessenger.TbmRecipientStatus]::Failed
        if ($_.Exception -is [System.OperationCanceledException]) { $status = [TeamsBlastMessenger.TbmRecipientStatus]::Cancelled }
        foreach ($index in $indexes) { Set-TbmRecipientOutcome -Job $Job -Index $index -Status $status -ErrorMessage $reason }
    }
}

#endregion
