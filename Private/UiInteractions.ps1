#region Interactions window

function Open-TbmInteractions {
    <#
    .SYNOPSIS
        Opens the interactions window for a blast, connecting to Graph first if needed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [TeamsBlastMessenger.TbmBlastRecord]$Blast,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Windows.Window]$Owner
    )

    if ($Blast.SentCount -eq 0) {
        $null = Show-TbmMessageBox -Text 'No messages in this blast were sent, so there are no replies to look for.' -Owner $Owner
        return
    }
    if (-not (Test-TbmGraphConnected -AuthState $script:TbmUi.AuthState) -and $Blast.Interactions.Count -eq 0) {
        $prompt = "Finding replies reads your Teams chats through Microsoft Graph, which isn't connected.`n`nConnect now?"
        if (-not (Confirm-TbmAction -Text $prompt)) { return }
        if (-not (Connect-TbmFromUi)) { return }
    }
    Show-TbmInteractionsWindow -Owner $Owner -Blast $Blast
}

function Show-TbmInteractionsWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Windows.Window]$Owner,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [TeamsBlastMessenger.TbmBlastRecord]$Blast
    )

    $ui = Get-TbmXamlWindow -Name 'InteractionsWindow' -Owner $Owner
    $interactionView = @{ Ui = $ui; Blast = $Blast; Job = $null; Sync = $null; Timer = $null }
    $c = $ui.Controls
    $c.TxtTitle.Text = 'Replies to "{0}"' -f $Blast.TemplatePreview
    $kind = 'individual messages'
    if ($Blast.Mode -eq 'Group') { $kind = 'a group chat' }
    $c.TxtChecked.Text = 'Sent {0} as {1} via {2} to {3} recipient(s).' -f $Blast.CreatedLocal, $kind, $Blast.Transport, $Blast.SentCount

    $interactionView.Show = {
        $view = $interactionView
        $controls = $view.Ui.Controls
        $results = $view.Blast.Interactions
        $controls.GridResults.ItemsSource = $null
        $controls.GridResults.ItemsSource = $results
        $controls.BtnExport.IsEnabled = $results.Count -gt 0
        $controls.TxtSummary.Text = ''
        $controls.TxtReplies.Text = ''
        if ($null -ne $view.Blast.LastInteractionCheckUtc) {
            $controls.TxtSummary.Text = Get-TbmInteractionSummary -Results $results
            $controls.TxtChecked.Text = 'Last checked {0}. Response times are measured from when each message was sent.' -f [TeamsBlastMessenger.TbmText]::FormatLocal($view.Blast.LastInteractionCheckUtc)
        }
    }

    $interactionView.Start = {
        $view = $interactionView
        if ($null -ne $view.Job) { return }
        if (-not (Test-TbmGraphConnected -AuthState $script:TbmUi.AuthState)) {
            if (-not (Confirm-TbmAction -Text 'Microsoft Graph is not connected. Connect now?')) { return }
            if (-not (Connect-TbmFromUi)) { return }
        }
        $null = Get-TbmAccessToken -AuthState $script:TbmUi.AuthState
        $view.Sync = New-TbmSyncState
        $job = @{ Blast = $view.Blast; AuthState = $script:TbmUi.AuthState; Settings = $script:TbmUi.Settings.Clone(); Sync = $view.Sync }
        $view.Job = Start-TbmBackgroundJob -CommandName 'Invoke-TbmInteractionWorker' -Job $job
        $controls = $view.Ui.Controls
        $controls.BtnScan.IsEnabled = $false
        $controls.BtnCancelScan.IsEnabled = $true
        $controls.PrgScan.IsIndeterminate = $true
        $controls.TxtProgress.Text = 'Reading chats...'
        $view.Timer.Start()
    }

    $interactionView.Timer = New-TbmDispatcherTimer -IntervalMilliseconds 250 -OnTick {
        $view = $interactionView
        if ($null -eq $view.Sync) { return }
        $item = $null
        while ($view.Sync.Events.TryDequeue([ref]$item)) {
            $controls = $view.Ui.Controls
            switch ($item.Type) {
                'Progress' {
                    $controls.TxtProgress.Text = [string]$item.Text
                    if ([int]$item.Total -gt 0 -and [int]$item.Current -gt 0) {
                        $controls.PrgScan.IsIndeterminate = $false
                        $controls.PrgScan.Maximum = [int]$item.Total
                        $controls.PrgScan.Value = [int]$item.Current
                    }
                }
                'Done' {
                    $view.Timer.Stop()
                    if ($null -ne $view.Job) { Complete-TbmBackgroundJob -Handle $view.Job }
                    $view.Job = $null
                    $view.Sync = $null
                    $controls.BtnScan.IsEnabled = $true
                    $controls.BtnCancelScan.IsEnabled = $false
                    $controls.PrgScan.IsIndeterminate = $false
                    $controls.PrgScan.Value = 0
                    if ([string]::IsNullOrEmpty([string]$item.Error)) {
                        $controls.TxtProgress.Text = 'Done.'
                    }
                    else {
                        $controls.TxtProgress.Text = 'Could not finish: ' + [string]$item.Error
                    }
                    & $view.Show
                    return
                }
            }
        }
    }

    $c.BtnScan.Add_Click({ Invoke-TbmUiAction -Name 'Check for replies' -Action { & $interactionView.Start } })
    $c.BtnCancelScan.Add_Click({
            if ($null -ne $interactionView.Sync) {
                $interactionView.Sync['Cancel'] = $true
                $interactionView.Ui.Controls.TxtProgress.Text = 'Cancelling...'
            }
        })
    $c.BtnClose.Add_Click({ $interactionView.Ui.Window.Close() })
    $c.BtnExport.Add_Click({
            Invoke-TbmUiAction -Name 'Export interactions' -Action {
                $blast = $interactionView.Blast
                $name = 'Interactions_{0}.csv' -f $blast.CreatedUtc.ToLocalTime().ToString('yyyyMMdd-HHmmss', [System.Globalization.CultureInfo]::InvariantCulture)
                $path = Show-TbmSaveFileDialog -FileName $name -Owner $interactionView.Ui.Window
                if ($null -ne $path) { Export-TbmInteractionCsv -Interactions $blast.Interactions -Path $path }
            }
        })
    $c.GridResults.Add_SelectionChanged({
            $result = $interactionView.Ui.Controls.GridResults.SelectedItem -as [TeamsBlastMessenger.TbmInteractionResult]
            $box = $interactionView.Ui.Controls.TxtReplies
            if ($null -eq $result) { $box.Text = ''; return }
            $interactionView.Ui.Controls.TxtRepliesTitle.Text = 'All replies from ' + $result.Recipient
            if ($result.Replies.Count -eq 0) {
                $notes = [System.Collections.Generic.List[string]]::new()
                if ([string]::IsNullOrEmpty($result.Error)) { $notes.Add('No replies since the message was sent.') } else { $notes.Add($result.Error) }
                if (-not [string]::IsNullOrEmpty($result.Reactions)) { $notes.Add('Reacted: ' + $result.Reactions) }
                $box.Text = [string]::Join(' ', $notes)
                return
            }
            $builder = [System.Text.StringBuilder]::new()
            foreach ($reply in $result.Replies) {
                $elapsed = [TeamsBlastMessenger.TbmText]::FormatDuration($reply.CreatedUtc - [datetime]$result.SentUtc)
                $quoted = ''
                if ($reply.IsDirectReply) { $quoted = '  (quoted your message)' }
                $null = $builder.AppendFormat('[{0}  +{1}] {2}{3}', $reply.CreatedLocal, $elapsed, $reply.FromName, $quoted).AppendLine()
                $null = $builder.AppendLine($reply.Text).AppendLine()
            }
            $box.Text = $builder.ToString()
        })
    $ui.Window.Add_Closing({
            param($source, $closingArgs)
            if ($null -eq $interactionView.Job) { return }
            $interactionView.Sync['Cancel'] = $true
            $interactionView.Timer.Stop()
            Wait-TbmBackgroundJob -Handle $interactionView.Job -TimeoutMilliseconds 15000
            $interactionView.Job = $null
        })
    $ui.Window.Add_Loaded({
            Invoke-TbmUiAction -Name 'Load interactions' -Action {
                & $interactionView.Show
                if ($null -eq $interactionView.Blast.LastInteractionCheckUtc) { & $interactionView.Start }
            }
        })
    $null = $ui.Window.ShowDialog()
}

#endregion
