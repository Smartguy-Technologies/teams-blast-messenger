#region History window

function Show-TbmHistoryWindow {
    <#
    .SYNOPSIS
        Shows previous blasts. Returns @{ Action = 'Retry' | 'Reuse'; Blast = ... } when the user picks one, else $null.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Windows.Window]$Owner
    )

    $ui = Get-TbmXamlWindow -Name 'HistoryWindow' -Owner $Owner
    $historyView = @{ Ui = $ui; Records = $null; Result = $null }

    $refresh = {
        $historyView.Records = Get-TbmBlastHistory
        $controls = $historyView.Ui.Controls
        $controls.GridBlasts.ItemsSource = $null
        $controls.GridBlasts.ItemsSource = $historyView.Records
        $controls.GridRecipients.ItemsSource = $null
        $controls.TxtMessage.Text = ''
        $messages = 0
        foreach ($record in $historyView.Records) { $messages += $record.SentCount }
        $controls.TxtSummary.Text = '{0} blast(s), {1} recipient message(s) delivered. Stored in {2}' -f $historyView.Records.Count, $messages, (Get-TbmAppDataPath -Name 'History')
        Update-TbmHistoryButtons -View $historyView
    }
    $historyView.Refresh = $refresh

    $c = $ui.Controls
    $c.GridBlasts.Add_SelectionChanged({
            Invoke-TbmUiAction -Name 'Select blast' -Action {
                $blast = Get-TbmSelectedHistoryBlast -View $historyView
                $historyView.Ui.Controls.GridRecipients.ItemsSource = $null
                if ($null -ne $blast) { $historyView.Ui.Controls.GridRecipients.ItemsSource = $blast.Recipients }
                $historyView.Ui.Controls.TxtMessage.Text = ''
                Update-TbmHistoryButtons -View $historyView
            }
        })
    $c.GridRecipients.Add_SelectionChanged({
            Invoke-TbmUiAction -Name 'Select recipient' -Action {
                $blast = Get-TbmSelectedHistoryBlast -View $historyView
                $recipient = $historyView.Ui.Controls.GridRecipients.SelectedItem -as [TeamsBlastMessenger.TbmRecipientResult]
                if ($null -eq $blast -or $null -eq $recipient) { return }
                $format = 'Text'
                if ($blast.Format -eq 'Html') { $format = 'Html' }
                $historyView.Ui.Controls.TxtMessage.Text = ConvertTo-TbmPlainText -Message ([string]$recipient.RenderedMessage) -Format $format
            }
        })
    $c.BtnRefresh.Add_Click({ Invoke-TbmUiAction -Name 'Refresh history' -Action { & $historyView.Refresh } })
    $c.BtnOpenFolder.Add_Click({ Invoke-TbmUiAction -Name 'Open history folder' -Action { Open-TbmFolder -Path (Get-TbmAppDataPath -Name 'History') } })
    $c.BtnInteractions.Add_Click({
            Invoke-TbmUiAction -Name 'Check interactions' -Action {
                $blast = Get-TbmSelectedHistoryBlast -View $historyView
                if ($null -eq $blast) { return }
                Open-TbmInteractions -Blast $blast -Owner $historyView.Ui.Window
                $historyView.Ui.Controls.GridBlasts.Items.Refresh()
            }
        })
    $c.BtnRetry.Add_Click({
            Invoke-TbmUiAction -Name 'Retry failed' -Action {
                $blast = Get-TbmSelectedHistoryBlast -View $historyView
                if ($null -eq $blast) { return }
                if ((Get-TbmRetryableRecipients -Blast $blast).Count -eq 0) {
                    $null = Show-TbmMessageBox -Text 'Every recipient in this blast was sent or skipped; there is nothing to retry.' -Owner $historyView.Ui.Window
                    return
                }
                $historyView.Result = @{ Action = 'Retry'; Blast = $blast }
                $historyView.Ui.Window.Close()
            }
        })
    $c.BtnReuse.Add_Click({
            Invoke-TbmUiAction -Name 'Reuse message' -Action {
                $blast = Get-TbmSelectedHistoryBlast -View $historyView
                if ($null -eq $blast) { return }
                $historyView.Result = @{ Action = 'Reuse'; Blast = $blast }
                $historyView.Ui.Window.Close()
            }
        })
    $c.BtnExport.Add_Click({
            Invoke-TbmUiAction -Name 'Export history' -Action {
                $blast = Get-TbmSelectedHistoryBlast -View $historyView
                if ($null -eq $blast) { return }
                $name = 'Blast_{0}.csv' -f $blast.CreatedUtc.ToLocalTime().ToString('yyyyMMdd-HHmmss', [System.Globalization.CultureInfo]::InvariantCulture)
                $path = Show-TbmSaveFileDialog -FileName $name -Owner $historyView.Ui.Window
                if ($null -ne $path) { Export-TbmBlastCsv -Blast $blast -Path $path }
            }
        })
    $c.BtnDelete.Add_Click({
            Invoke-TbmUiAction -Name 'Delete history' -Action {
                $blast = Get-TbmSelectedHistoryBlast -View $historyView
                if ($null -eq $blast) { return }
                if (-not (Confirm-TbmAction -Text ('Delete the history record for the blast sent {0} to {1} recipient(s)? Messages already sent in Teams are not affected.' -f $blast.CreatedLocal, $blast.RecipientCount) -Icon ([System.Windows.MessageBoxImage]::Warning))) { return }
                Remove-TbmBlastRecord -Blast $blast -Confirm:$false
                & $historyView.Refresh
            }
        })

    $ui.Window.Add_Loaded({ Invoke-TbmUiAction -Name 'Load history' -Action { & $historyView.Refresh } })
    $null = $ui.Window.ShowDialog()
    return $historyView.Result
}

function Get-TbmSelectedHistoryBlast {
    <#
    .SYNOPSIS
        The selected blast, fully loaded (the list is built from summaries).
    #>
    [CmdletBinding()]
    [OutputType([TeamsBlastMessenger.TbmBlastRecord])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$View
    )
    $blast = $View.Ui.Controls.GridBlasts.SelectedItem -as [TeamsBlastMessenger.TbmBlastRecord]
    if ($null -eq $blast) { return $null }
    return Expand-TbmBlastRecord -Blast $blast
}

function Update-TbmHistoryButtons {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$View
    )
    $controls = $View.Ui.Controls
    $hasSelection = $null -ne $controls.GridBlasts.SelectedItem
    foreach ($name in @('BtnInteractions', 'BtnRetry', 'BtnReuse', 'BtnExport', 'BtnDelete')) { $controls[$name].IsEnabled = $hasSelection }
}

function Get-TbmRetryableRecipients {
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[TeamsBlastMessenger.TbmRecipientResult]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [TeamsBlastMessenger.TbmBlastRecord]$Blast
    )
    $list = [System.Collections.Generic.List[TeamsBlastMessenger.TbmRecipientResult]]::new()
    foreach ($recipient in $Blast.Recipients) {
        if ($recipient.Status -eq [TeamsBlastMessenger.TbmRecipientStatus]::Failed -or
            $recipient.Status -eq [TeamsBlastMessenger.TbmRecipientStatus]::Cancelled -or
            $recipient.Status -eq [TeamsBlastMessenger.TbmRecipientStatus]::Pending) {
            $list.Add($recipient)
        }
    }
    return , $list
}

#endregion
