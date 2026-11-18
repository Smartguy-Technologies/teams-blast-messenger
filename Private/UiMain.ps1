#region Main window: wiring
<#
    All main-window state lives in $script:TbmUi (created by Show-TbmMainWindow). Every handler runs on the
    UI thread and is wrapped in Invoke-TbmUiAction. Background sends report back through $script:TbmUi.Sync,
    drained by Invoke-TbmMainTimerTick.
#>

function Register-TbmMainWindowEvents {
    [CmdletBinding()]
    param()

    $c = $script:TbmUi.C
    $window = $script:TbmUi.Window

    $window.Dispatcher.Add_UnhandledException({
            param($source, $exceptionArgs)
            Write-TbmLog -Level Error -Source 'UI' -Message ('Unhandled UI exception: ' + $exceptionArgs.Exception.ToString())
            $exceptionArgs.Handled = $true
            [void](Show-TbmMessageBox -Text ('Unexpected error: ' + $exceptionArgs.Exception.Message) -Icon ([System.Windows.MessageBoxImage]::Error))
        })
    $window.Add_Loaded({ Invoke-TbmUiAction -Name 'Startup' -Action { Initialize-TbmMainWindowState } })
    $window.Add_Closing({
            param($source, $closingArgs)
            Invoke-TbmUiAction -Name 'Close' -Action {
                if ($script:TbmUi.Busy) {
                    if (-not (Confirm-TbmAction -Text 'A send is in progress. Stop after the current recipient and close?' -Icon ([System.Windows.MessageBoxImage]::Warning))) {
                        $closingArgs.Cancel = $true
                        return
                    }
                    Stop-TbmCurrentSend
                    if ($null -ne $script:TbmUi.Job) {
                        Wait-TbmBackgroundJob -Handle $script:TbmUi.Job
                        $script:TbmUi.Job = $null
                    }
                }
                Save-TbmDraft
                $script:TbmUi.Settings.Transport = Get-TbmComboTag -ComboBox $script:TbmUi.C.CmbTransport
                $script:TbmUi.Settings.MessageFormat = Get-TbmComboTag -ComboBox $script:TbmUi.C.CmbFormat
                Save-TbmSettings -Settings $script:TbmUi.Settings
            }
        })

    $c.BtnConnect.Add_Click({ Invoke-TbmUiAction -Name 'Connect' -Action { [void](Connect-TbmFromUi) } })
    $c.BtnDisconnect.Add_Click({
            Invoke-TbmUiAction -Name 'Disconnect' -Action {
                Disconnect-TbmGraph -AuthState $script:TbmUi.AuthState
                Update-TbmConnectionDisplay
                Update-TbmPreview
            }
        })
    $c.BtnSettings.Add_Click({
            Invoke-TbmUiAction -Name 'Settings' -Action {
                if (Show-TbmSettingsWindow -Owner $script:TbmUi.Window -Settings $script:TbmUi.Settings) {
                    Update-TbmConnectionDisplay
                    Update-TbmPreview
                }
            }
        })
    $c.BtnHistory.Add_Click({ Invoke-TbmUiAction -Name 'History' -Action { Show-TbmHistoryFromUi } })
    $c.BtnInteractions.Add_Click({ Invoke-TbmUiAction -Name 'Interactions' -Action { Open-TbmLatestInteractions } })
    $c.BtnOpenLogs.Add_Click({ Invoke-TbmUiAction -Name 'Open logs' -Action { Open-TbmFolder -Path (Get-TbmAppDataPath -Name 'Logs') } })

    $c.BtnBrowse.Add_Click({ Invoke-TbmUiAction -Name 'Open spreadsheet' -Action { Select-TbmRecipientFile } })
    $c.BtnReload.Add_Click({
            Invoke-TbmUiAction -Name 'Reload' -Action {
                if ($null -eq $script:TbmUi.SourceFile) { return }
                if ($script:TbmUi.ManualCount -gt 0 -and -not (Confirm-TbmAction -Text ('Reloading the file removes the {0} recipient(s) you added manually. Continue?' -f $script:TbmUi.ManualCount) -Icon ([System.Windows.MessageBoxImage]::Warning))) { return }
                Import-TbmUiRecipients -Path $script:TbmUi.SourceFile -SheetName $script:TbmUi.SheetName
            }
        })
    $c.BtnAddManual.Add_Click({ Invoke-TbmUiAction -Name 'Add recipients manually' -Action { Add-TbmManualRecipientsFromUi } })
    $c.CmbSheet.Add_SelectionChanged({
            Invoke-TbmUiAction -Name 'Change sheet' -Action {
                if ($script:TbmUi.Suppress -or $null -eq $script:TbmUi.SourceFile) { return }
                $sheet = [string]$script:TbmUi.C.CmbSheet.SelectedItem
                if (-not [string]::IsNullOrEmpty($sheet) -and $sheet -ne $script:TbmUi.SheetName) { Import-TbmUiRecipients -Path $script:TbmUi.SourceFile -SheetName $sheet }
            }
        })
    $c.CmbRecipientColumn.Add_SelectionChanged({
            Invoke-TbmUiAction -Name 'Change recipient column' -Action {
                if ($script:TbmUi.Suppress -or $null -eq $script:TbmUi.Table) { return }
                $column = [string]$script:TbmUi.C.CmbRecipientColumn.SelectedItem
                if ([string]::IsNullOrEmpty($column)) { return }
                $script:TbmUi.RecipientColumn = $column
                Invoke-TbmSuppressed { [void](Update-TbmRecipientValidation -Table $script:TbmUi.Table -RecipientColumn $column) }
                Update-TbmRecipientSummary
                Update-TbmPreview
            }
        })
    $c.BtnSelectAll.Add_Click({ Invoke-TbmUiAction -Name 'Select valid' -Action { Set-TbmAllIncluded -Include $true } })
    $c.BtnSelectNone.Add_Click({ Invoke-TbmUiAction -Name 'Select none' -Action { Set-TbmAllIncluded -Include $false } })
    $c.GridRecipients.Add_AutoGeneratingColumn({
            param($source, $columnArgs)
            Set-TbmGeneratedColumn -ColumnArgs $columnArgs
        })
    $c.GridRecipients.Add_SelectionChanged({ Invoke-TbmUiAction -Name 'Preview' -Action { Update-TbmPreview } })

    $groupToggled = {
        Invoke-TbmUiAction -Name 'Change mode' -Action {
            $script:TbmUi.C.TxtTopic.IsEnabled = [bool]$script:TbmUi.C.RadGroup.IsChecked
            Update-TbmPreview
        }
    }
    $c.RadGroup.Add_Checked($groupToggled)
    $c.RadIndividual.Add_Checked($groupToggled)
    $c.CmbFormat.Add_SelectionChanged({ Invoke-TbmUiAction -Name 'Change format' -Action { Update-TbmPreview } })
    $c.TxtMessage.Add_TextChanged({ Invoke-TbmUiAction -Name 'Preview' -Action { Update-TbmPreview } })
    $c.LstTokens.Add_MouseDoubleClick({ Invoke-TbmUiAction -Name 'Insert token' -Action { Add-TbmTokenAtCaret } })

    $c.BtnPreviewAll.Add_Click({ Invoke-TbmUiAction -Name 'Dry run' -Action { Invoke-TbmDryRun } })
    $c.BtnTest.Add_Click({ Invoke-TbmUiAction -Name 'Test send' -Action { Start-TbmTestSend } })
    $c.BtnSend.Add_Click({ Invoke-TbmUiAction -Name 'Send' -Action { Start-TbmSendFromUi } })
    $c.BtnCancel.Add_Click({ Invoke-TbmUiAction -Name 'Cancel' -Action { Stop-TbmCurrentSend } })
    $c.BtnManualNext.Add_Click({ Invoke-TbmUiAction -Name 'Next recipient' -Action { Complete-TbmManualStep -Status Sent } })
    $c.BtnManualSkip.Add_Click({ Invoke-TbmUiAction -Name 'Skip recipient' -Action { Complete-TbmManualStep -Status Skipped } })
}

function Initialize-TbmMainWindowState {
    [CmdletBinding()]
    param()

    $ui = $script:TbmUi
    Set-TbmComboByTag -ComboBox $ui.C.CmbTransport -Value ([string]$ui.Settings.Transport)
    Set-TbmComboByTag -ComboBox $ui.C.CmbFormat -Value ([string]$ui.Settings.MessageFormat)
    [void](Set-TbmPolicyLock -Control $ui.C.CmbTransport -SettingName 'Transport')
    [void](Set-TbmPolicyLock -Control $ui.C.CmbFormat -SettingName 'MessageFormat')
    Restore-TbmDraft
    Update-TbmTokenList
    Update-TbmConnectionDisplay
    Update-TbmPreview

    if ((Resolve-TbmAuthMode -Settings $ui.Settings) -eq 'GraphRest') {
        $ui.C.TxtProgress.Text = 'Restoring Microsoft Graph session…'
        if (Connect-TbmGraphRestSilent -AuthState $ui.AuthState -Settings $ui.Settings) { Update-TbmConnectionDisplay; Update-TbmPreview }
        $ui.C.TxtProgress.Text = 'Ready.'
    }
}

function Invoke-TbmSuppressed {
    <#
    .SYNOPSIS
        Runs a block with change events suppressed (bulk updates).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [scriptblock]$Action
    )
    $previous = $script:TbmUi.Suppress
    $script:TbmUi.Suppress = $true
    try { & $Action } finally { $script:TbmUi.Suppress = $previous }
}

#endregion

#region Main window: connection

function Connect-TbmFromUi {
    <#
    .SYNOPSIS
        Signs in with the configured method. Returns $true when connected.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $ui = $script:TbmUi
    $mode = Resolve-TbmAuthMode -Settings $ui.Settings
    if ($null -eq $mode) {
        $text = "Microsoft Graph isn't set up on this PC yet. Do one of these:`n`n" +
        "  - Install the Graph PowerShell SDK:  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`n" +
        "  - Enter an Entra app registration Client ID in Settings (see README).`n`n" +
        'Until then you can still send with the deep-link options.'
        [void](Show-TbmMessageBox -Text $text -Icon ([System.Windows.MessageBoxImage]::Information))
        return $false
    }

    $owner = Get-TbmActiveWindow
    $owner.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        if ($mode -eq 'GraphSdk') {
            Connect-TbmGraphSdk -AuthState $ui.AuthState -Settings $ui.Settings
        }
        elseif (-not (Connect-TbmGraphRestSilent -AuthState $ui.AuthState -Settings $ui.Settings)) {
            $owner.Cursor = $null
            if (-not (Show-TbmDeviceCodeWindow -Owner $owner -Settings $ui.Settings -AuthState $ui.AuthState)) { return $false }
        }
    }
    finally {
        $owner.Cursor = $null
        Update-TbmConnectionDisplay
        Update-TbmPreview
    }
    return Test-TbmGraphConnected -AuthState $ui.AuthState
}

function Update-TbmConnectionDisplay {
    [CmdletBinding()]
    param()

    $ui = $script:TbmUi
    $converter = [System.Windows.Media.BrushConverter]::new()
    if (Test-TbmGraphConnected -AuthState $ui.AuthState) {
        $method = 'Graph SDK'
        if ($ui.AuthState.Mode -eq 'GraphRest') { $method = 'app registration' }
        $ui.C.EllConnection.Fill = $converter.ConvertFromString('#3FB950')
        $ui.C.TxtConnection.Text = 'Connected as {0} ({1}) via {2}' -f $ui.AuthState.MeName, $ui.AuthState.MeUpn, $method
        $ui.C.BtnConnect.Visibility = [System.Windows.Visibility]::Collapsed
        $ui.C.BtnDisconnect.Visibility = [System.Windows.Visibility]::Visible
    }
    else {
        $ui.C.EllConnection.Fill = $converter.ConvertFromString('#F0B429')
        $ui.C.TxtConnection.Text = 'Not connected to Graph · deep-link fallback available'
        $ui.C.BtnConnect.Visibility = [System.Windows.Visibility]::Visible
        $ui.C.BtnDisconnect.Visibility = [System.Windows.Visibility]::Collapsed
    }
}

function Resolve-TbmUiTransport {
    <#
    .SYNOPSIS
        Maps the "Send via" choice to a concrete transport: GraphSdk, GraphRest, DeepLinkAuto or DeepLinkManual.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    $ui = $script:TbmUi
    $connected = Test-TbmGraphConnected -AuthState $ui.AuthState
    switch (Get-TbmComboTag -ComboBox $ui.C.CmbTransport) {
        'Graph' {
            if (-not $connected) { throw [System.InvalidOperationException]::new('Connect to Microsoft Graph first, or choose a deep-link option under "Send via".') }
            return [string]$ui.AuthState.Mode
        }
        'DeepLinkAuto' { return 'DeepLinkAuto' }
        'DeepLinkManual' { return 'DeepLinkManual' }
        default {
            if ($connected) { return [string]$ui.AuthState.Mode }
            return 'DeepLinkManual'
        }
    }
}

function Get-TbmTransportLabel {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Transport
    )
    switch ($Transport) {
        'GraphSdk' { return 'Microsoft Graph (SDK)' }
        'GraphRest' { return 'Microsoft Graph (app registration)' }
        'DeepLinkAuto' { return 'deep link with auto-send' }
        'DeepLinkManual' { return 'deep link (you press Send)' }
    }
    return $Transport
}

#endregion

#region Main window: recipients

function Select-TbmRecipientFile {
    [CmdletBinding()]
    param()

    $dialog = [Microsoft.Win32.OpenFileDialog]::new()
    $dialog.Filter = 'Spreadsheets (*.csv;*.xlsx;*.xlsm;*.xls;*.tsv;*.txt)|*.csv;*.xlsx;*.xlsm;*.xls;*.tsv;*.txt|All files (*.*)|*.*'
    $lastFolder = [string]$script:TbmUi.Settings.LastFolder
    if (-not [string]::IsNullOrEmpty($lastFolder) -and [System.IO.Directory]::Exists($lastFolder)) { $dialog.InitialDirectory = $lastFolder }
    if ($dialog.ShowDialog($script:TbmUi.Window) -ne $true) { return }
    Import-TbmUiRecipients -Path $dialog.FileName
}

function Import-TbmUiRecipients {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter()]
        [string]$SheetName
    )

    $ui = $script:TbmUi
    $ui.Window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        $result = Import-TbmRecipientFile -Path $Path -SheetName $SheetName
    }
    finally {
        $ui.Window.Cursor = $null
    }

    $ui.SourceFile = $Path
    $ui.SheetName = $result.SheetName
    $ui.Settings.LastFolder = [System.IO.Path]::GetDirectoryName($Path)
    Set-TbmUiRecipientTable -Table $result.Table -RecipientColumn (Find-TbmRecipientColumn -Table $result.Table) -Sheets $result.Sheets -SheetName $result.SheetName -SourceLabel $Path

    if ($null -eq $ui.RecipientColumn) {
        [void](Show-TbmMessageBox -Text 'No email column was recognised. Choose the column that holds Teams email addresses under "Recipient column".' -Icon ([System.Windows.MessageBoxImage]::Warning))
    }
}

function Set-TbmUiRecipientTable {
    <#
    .SYNOPSIS
        Makes a recipient table (imported or manual) the current list and refreshes the grid, tokens and preview.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Data.DataTable]$Table,

        [Parameter()]
        [string]$RecipientColumn,

        [Parameter()]
        [string[]]$Sheets,

        [Parameter()]
        [string]$SheetName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$SourceLabel
    )

    $ui = $script:TbmUi
    $ui.Table = $Table
    $ui.ColumnNames = Get-TbmDataColumnNames -Table $Table
    $ui.RecipientColumn = $null
    if (-not [string]::IsNullOrEmpty($RecipientColumn)) { $ui.RecipientColumn = $RecipientColumn }
    $ui.SourceLabel = $SourceLabel
    $ui.ManualCount = 0

    Invoke-TbmSuppressed {
        $controls = $script:TbmUi.C
        $controls.CmbSheet.ItemsSource = $Sheets
        $controls.CmbSheet.SelectedItem = $SheetName
        $controls.CmbRecipientColumn.ItemsSource = $script:TbmUi.ColumnNames
        $controls.CmbRecipientColumn.SelectedItem = $script:TbmUi.RecipientColumn
        if ($null -ne $script:TbmUi.RecipientColumn) {
            [void](Update-TbmRecipientValidation -Table $script:TbmUi.Table -RecipientColumn $script:TbmUi.RecipientColumn)
        }
        $controls.GridRecipients.ItemsSource = $null
        $controls.GridRecipients.ItemsSource = $script:TbmUi.Table.DefaultView
    }

    $Table.Add_ColumnChanged({
            param($source, $changeArgs)
            if (-not $script:TbmUi.Suppress -and $changeArgs.Column.ColumnName -eq $script:TbmIncludeColumn) { Update-TbmRecipientSummary }
        })

    $ui.C.TxtFile.Text = $SourceLabel
    $ui.C.TxtFile.ClearValue([System.Windows.Controls.Control]::ForegroundProperty)
    Update-TbmFileControls
    Update-TbmRecipientSummary
    Update-TbmTokenList
    Update-TbmPreview
}

function Add-TbmManualRecipientsFromUi {
    <#
    .SYNOPSIS
        Opens the manual entry dialog and adds the people to the current list, or starts a new list.
    #>
    [CmdletBinding()]
    param()

    $ui = $script:TbmUi
    $hasList = $null -ne $ui.Table -and $null -ne $ui.RecipientColumn
    $choice = Show-TbmManualRecipientsWindow -Owner $ui.Window -AuthState $ui.AuthState -HasList $hasList
    if ($null -eq $choice) { return }

    if ($choice.Replace -or -not $hasList) {
        if ($null -ne $ui.Table -and $ui.Table.Rows.Count -gt 0 -and -not (Confirm-TbmAction -Text 'Replace the current recipient list with the people you entered?')) { return }
        $result = Add-TbmManualEntriesToTable -Entries $choice.Entries
        $ui.SourceFile = $null
        $ui.SheetName = $null
        Set-TbmUiRecipientTable -Table $result.Table -RecipientColumn $result.RecipientColumn -SourceLabel ('Entered manually ({0})' -f $result.Added)
        $ui.ManualCount = $result.Added
        return
    }

    [void]$ui.C.GridRecipients.CommitEdit()
    $previous = $ui.Suppress
    $ui.Suppress = $true
    try {
        $result = Add-TbmManualEntriesToTable -Entries $choice.Entries -Table $ui.Table -RecipientColumn $ui.RecipientColumn
        [void](Update-TbmRecipientValidation -Table $ui.Table -RecipientColumn $ui.RecipientColumn)
    }
    finally {
        $ui.Suppress = $previous
    }
    $ui.ManualCount += $result.Added
    $baseLabel = $ui.SourceFile
    if ([string]::IsNullOrEmpty($baseLabel)) { $baseLabel = 'Entered manually' }
    $ui.SourceLabel = '{0} + {1} added manually' -f $baseLabel, $ui.ManualCount
    $ui.C.TxtFile.Text = $ui.SourceLabel
    Update-TbmRecipientSummary
    Update-TbmPreview
    $ui.C.GridRecipients.ScrollIntoView($ui.Table.DefaultView[$ui.Table.DefaultView.Count - 1])
}

function Set-TbmGeneratedColumn {
    <#
    .SYNOPSIS
        Customises auto-generated DataGrid columns: instant-commit "Send" checkbox, read-only Status,
        literal headers (no access-key underscores) and indexer bindings safe for any column name.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Windows.Controls.DataGridAutoGeneratingColumnEventArgs]$ColumnArgs
    )

    $name = [string]$ColumnArgs.PropertyName
    if ($name -eq $script:TbmIncludeColumn) {
        $template = '<DataTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"><CheckBox HorizontalAlignment="Center" VerticalAlignment="Center" IsChecked="{Binding [TBM_Include], Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"/></DataTemplate>'
        $column = [System.Windows.Controls.DataGridTemplateColumn]::new()
        $column.Header = 'Send'
        $column.CellTemplate = [System.Windows.Markup.XamlReader]::Parse($template)
        $column.SortMemberPath = $name
        $column.Width = [System.Windows.Controls.DataGridLength]::new(48)
        $ColumnArgs.Column = $column
        return
    }

    $headerText = $name
    if ($name -eq $script:TbmStatusColumn) { $headerText = 'Status' }
    $header = [System.Windows.Controls.TextBlock]::new()
    $header.Text = $headerText
    $ColumnArgs.Column.Header = $header
    $ColumnArgs.Column.SortMemberPath = $name

    $boundColumn = $ColumnArgs.Column -as [System.Windows.Controls.DataGridBoundColumn]
    if ($null -ne $boundColumn) {
        $escaped = $name.Replace('^', '^^').Replace('[', '^[').Replace(']', '^]').Replace(',', '^,')
        $binding = [System.Windows.Data.Binding]::new('[' + $escaped + ']')
        $binding.Mode = [System.Windows.Data.BindingMode]::TwoWay
        $boundColumn.Binding = $binding
    }
    if ($name -eq $script:TbmStatusColumn) {
        $ColumnArgs.Column.IsReadOnly = $true
        $ColumnArgs.Column.Width = [System.Windows.Controls.DataGridLength]::new(100)
    }
}

function Update-TbmFileControls {
    [CmdletBinding()]
    param()
    $ui = $script:TbmUi
    $loaded = $null -ne $ui.Table -and -not $ui.Busy
    $ui.C.BtnReload.IsEnabled = $loaded -and $null -ne $ui.SourceFile
    $ui.C.BtnAddManual.IsEnabled = -not $ui.Busy
    $ui.C.CmbRecipientColumn.IsEnabled = $loaded
    $ui.C.BtnSelectAll.IsEnabled = $loaded
    $ui.C.BtnSelectNone.IsEnabled = $loaded
    $sheetCount = 0
    if ($null -ne $ui.C.CmbSheet.ItemsSource) { $sheetCount = @($ui.C.CmbSheet.ItemsSource).Count }
    $ui.C.CmbSheet.IsEnabled = $loaded -and $sheetCount -gt 1
}

function Set-TbmAllIncluded {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [bool]$Include
    )
    $ui = $script:TbmUi
    if ($null -eq $ui.Table) { return }
    [void]$ui.C.GridRecipients.CommitEdit()
    Invoke-TbmSuppressed {
        foreach ($row in $script:TbmUi.Table.Rows) {
            $status = [string]$row[$script:TbmStatusColumn]
            $valid = $status -ne 'Invalid email' -and $status -ne 'Duplicate'
            $row[$script:TbmIncludeColumn] = $Include -and $valid
        }
    }
    Update-TbmRecipientSummary
    Update-TbmPreview
}

function Update-TbmRecipientSummary {
    [CmdletBinding()]
    param()
    $ui = $script:TbmUi
    if ($null -eq $ui.Table) { $ui.C.TxtRecipientSummary.Text = 'No recipients yet. Import a CSV or Excel file (recommended), or use Add manually.'; return }
    $selected = 0; $invalid = 0; $duplicates = 0
    foreach ($row in $ui.Table.Rows) {
        if ([bool]$row[$script:TbmIncludeColumn]) { $selected++ }
        switch ([string]$row[$script:TbmStatusColumn]) {
            'Invalid email' { $invalid++ }
            'Duplicate' { $duplicates++ }
        }
    }
    $text = '{0} row(s) · {1} selected' -f $ui.Table.Rows.Count, $selected
    if ($invalid -gt 0) { $text += ' · {0} invalid email' -f $invalid }
    if ($duplicates -gt 0) { $text += ' · {0} duplicate' -f $duplicates }
    if ($null -ne $ui.RecipientColumn) { $text += ' · sending to column "{0}"' -f $ui.RecipientColumn }
    $ui.C.TxtRecipientSummary.Text = $text
}

function Get-TbmIncludedRows {
    <#
    .SYNOPSIS
        Ticked rows with a valid email, in grid (table) order.
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[System.Data.DataRow]])]
    param()

    $ui = $script:TbmUi
    $rows = [System.Collections.Generic.List[System.Data.DataRow]]::new()
    if ($null -eq $ui.Table -or $null -eq $ui.RecipientColumn) { return , $rows }
    [void]$ui.C.GridRecipients.CommitEdit()
    [void]$ui.C.GridRecipients.CommitEdit()
    foreach ($row in $ui.Table.Rows) {
        if ($row.HasVersion([System.Data.DataRowVersion]::Proposed)) { $row.EndEdit() }
        if (-not [bool]$row[$script:TbmIncludeColumn]) { continue }
        if (-not (Test-TbmEmailAddress -Value ([string]$row[$ui.RecipientColumn]))) { continue }
        $rows.Add($row)
    }
    return , $rows
}

#endregion

#region Main window: message editor and preview

function Update-TbmTokenList {
    [CmdletBinding()]
    param()
    $ui = $script:TbmUi
    $tokens = [System.Collections.Generic.List[string]]::new()
    foreach ($column in $ui.ColumnNames) { $tokens.Add('{' + $column + '}') }
    foreach ($name in $script:TbmBuiltInTokenNames) { $tokens.Add('{' + $name + '}') }
    $ui.C.LstTokens.ItemsSource = $tokens
}

function Add-TbmTokenAtCaret {
    [CmdletBinding()]
    param()
    $ui = $script:TbmUi
    $token = [string]$ui.C.LstTokens.SelectedItem
    if ([string]::IsNullOrEmpty($token)) { return }
    $box = $ui.C.TxtMessage
    $start = $box.SelectionStart
    $box.SelectedText = $token
    $box.SelectionLength = 0
    $box.CaretIndex = $start + $token.Length
    [void]$box.Focus()
}

function Get-TbmUiFormat {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    if ((Get-TbmComboTag -ComboBox $script:TbmUi.C.CmbFormat) -eq 'Html') { return 'Html' }
    return 'Text'
}

function Get-TbmUiPreviewRow {
    <#
    .SYNOPSIS
        The selected grid row, else the first ticked row, else the first row; $null without data.
    #>
    [CmdletBinding()]
    [OutputType([System.Data.DataRow])]
    param()
    $ui = $script:TbmUi
    if ($null -eq $ui.Table -or $ui.Table.Rows.Count -eq 0) { return $null }
    $view = $ui.C.GridRecipients.SelectedItem -as [System.Data.DataRowView]
    if ($null -ne $view) { return $view.Row }
    foreach ($row in $ui.Table.Rows) { if ([bool]$row[$script:TbmIncludeColumn]) { return $row } }
    return $ui.Table.Rows[0]
}

function New-TbmUiTokenContext {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()
    return New-TbmTokenContext -ColumnNames $script:TbmUi.ColumnNames -BuiltIns (Get-TbmBuiltInTokenValues -AuthState $script:TbmUi.AuthState)
}

function Update-TbmPreview {
    [CmdletBinding()]
    param()

    $ui = $script:TbmUi
    if ($null -eq $ui -or $ui.Suppress -or -not $ui.Window.IsLoaded) { return }
    $template = $ui.C.TxtMessage.Text
    $format = Get-TbmUiFormat
    $context = New-TbmUiTokenContext
    $row = Get-TbmUiPreviewRow
    $values = $null
    if ($null -ne $row) { $values = ConvertTo-TbmRowDictionary -Row $row -ColumnNames $ui.ColumnNames }

    $expanded = Expand-TbmTemplate -Template $template -Context $context -Row $values -HtmlEncodeValues:($format -eq 'Html')
    $ui.C.TxtPreview.Text = ConvertTo-TbmPlainText -Message $expanded.Text -Format $format
    if ($null -ne $row -and $null -ne $ui.RecipientColumn) {
        $ui.C.TxtPreviewRecipient.Text = 'to ' + [string]$row[$ui.RecipientColumn]
    }
    else {
        $ui.C.TxtPreviewRecipient.Text = '(load a spreadsheet to preview with real data)'
    }

    $notes = [System.Collections.Generic.List[string]]::new()
    $check = Test-TbmTemplate -Template $template -Context $context
    if ($check.UnknownTokens.Count -gt 0) {
        $notes.Add(('No matching column for {0}; these will be sent exactly as typed.' -f [string]::Join(', ', (ConvertTo-TbmTokenDisplay -Names $check.UnknownTokens))))
    }
    if ($null -ne $row -and $expanded.MissingTokens.Count -gt 0) {
        $notes.Add(('Empty for this recipient: {0}. Tip: add a fallback, e.g. {{{1}|there}}.' -f [string]::Join(', ', (ConvertTo-TbmTokenDisplay -Names $expanded.MissingTokens)), $expanded.MissingTokens[0]))
    }
    if ([bool]$ui.C.RadGroup.IsChecked -and $check.UsesRowData) {
        $notes.Add('Group mode sends one identical message to everyone; column tokens will use the first selected recipient''s values.')
    }
    foreach ($collision in $context.Collisions) { $notes.Add(('Columns {0} match the same token; the first one is used.' -f $collision)) }
    $ui.C.TxtValidation.Text = [string]::Join("`n", $notes)
}

function ConvertTo-TbmTokenDisplay {
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string]]$Names
    )
    $display = [System.Collections.Generic.List[string]]::new()
    foreach ($name in $Names) { if (-not $display.Contains('{' + $name + '}')) { $display.Add('{' + $name + '}') } }
    return , $display.ToArray()
}

function Save-TbmDraft {
    [CmdletBinding()]
    param()
    $ui = $script:TbmUi
    $draft = [ordered]@{
        Template = $ui.C.TxtMessage.Text
        Mode     = $(if ([bool]$ui.C.RadGroup.IsChecked) { 'Group' } else { 'Individual' })
        Topic    = $ui.C.TxtTopic.Text
    }
    Write-TbmTextFileAtomic -Path ([System.IO.Path]::Combine((Get-TbmAppDataPath -Name 'Cache'), 'draft.json')) -Content (ConvertTo-Json -InputObject $draft)
}

function Restore-TbmDraft {
    [CmdletBinding()]
    param()
    $path = [System.IO.Path]::Combine((Get-TbmAppDataPath -Name 'Cache'), 'draft.json')
    if (-not [System.IO.File]::Exists($path)) { return }
    try {
        $draft = ConvertFrom-Json -InputObject ([System.IO.File]::ReadAllText($path))
        $template = [string](Get-TbmPropertyValue -InputObject $draft -Name 'Template')
        if (-not [string]::IsNullOrEmpty($template)) { $script:TbmUi.C.TxtMessage.Text = $template }
        if ([string](Get-TbmPropertyValue -InputObject $draft -Name 'Mode') -eq 'Group') { $script:TbmUi.C.RadGroup.IsChecked = $true }
        $script:TbmUi.C.TxtTopic.Text = [string](Get-TbmPropertyValue -InputObject $draft -Name 'Topic')
    }
    catch {
        Write-TbmLog -Level Warn -Source 'UI' -Message ('Could not restore the message draft: {0}' -f $_.Exception.Message)
    }
}

#endregion

#region Main window: building and running a send

function New-TbmSendPlan {
    <#
    .SYNOPSIS
        Validates the current message/recipients with the user and builds the blast record and work items.
        Returns $null if the user declines a warning.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter()]
        [switch]$IsTest,

        [Parameter()]
        [System.Data.DataRow]$TestRow,

        [Parameter()]
        [string]$TestRecipient
    )

    $ui = $script:TbmUi
    $template = $ui.C.TxtMessage.Text
    if ([string]::IsNullOrWhiteSpace($template)) { throw [System.InvalidOperationException]::new('Write a message first.') }
    $format = Get-TbmUiFormat
    $mode = 'Individual'
    if (-not $IsTest -and [bool]$ui.C.RadGroup.IsChecked) { $mode = 'Group' }
    $transport = Resolve-TbmUiTransport
    $context = New-TbmUiTokenContext
    $check = Test-TbmTemplate -Template $template -Context $context

    $sources = [System.Collections.Generic.List[object]]::new()
    if ($IsTest) {
        $sources.Add($TestRow)
    }
    else {
        if ($null -eq $ui.Table) { throw [System.InvalidOperationException]::new('Import a recipient spreadsheet (or add people manually) first.') }
        if ($null -eq $ui.RecipientColumn) { throw [System.InvalidOperationException]::new('Choose the recipient (email) column first.') }
        foreach ($row in (Get-TbmIncludedRows)) { $sources.Add($row) }
        if ($sources.Count -eq 0) { throw [System.InvalidOperationException]::new('Tick at least one recipient with a valid email address.') }
    }

    if ($check.UnknownTokens.Count -gt 0) {
        $text = 'These tokens do not match any column and will be sent exactly as typed: {0}.{1}{1}Send anyway?' -f [string]::Join(', ', (ConvertTo-TbmTokenDisplay -Names $check.UnknownTokens)), "`n"
        if (-not (Confirm-TbmAction -Text $text -Icon ([System.Windows.MessageBoxImage]::Warning))) { return $null }
    }
    if ($mode -eq 'Group') {
        if ($sources.Count -gt 249) { throw [System.InvalidOperationException]::new('A Teams group chat holds at most 250 people including you. Select 249 recipients or fewer, or send individually.') }
        if ($check.UsesRowData) {
            $text = "Group mode sends one identical message. Column tokens will use the first selected recipient's values.`n`nContinue?"
            if (-not (Confirm-TbmAction -Text $text)) { return $null }
        }
    }

    $blast = [TeamsBlastMessenger.TbmBlastRecord]::new()
    $blast.Mode = $mode
    $blast.Transport = $transport
    $blast.Template = $template
    $blast.Format = $format
    $blast.SourceFile = [string]$ui.SourceLabel
    $blast.IsTest = [bool]$IsTest
    if ($mode -eq 'Group') { $blast.Topic = $ui.C.TxtTopic.Text.Trim() }
    if ($transport -like 'Graph*') {
        $blast.SenderId = [string]$ui.AuthState.MeId
        $blast.SenderUpn = [string]$ui.AuthState.MeUpn
    }

    $items = [System.Collections.Generic.List[hashtable]]::new()
    $missingCounts = [System.Collections.Generic.Dictionary[string, int]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $messagesWithMissing = 0
    $groupMessage = $null
    foreach ($source in $sources) {
        $row = $source -as [System.Data.DataRow]
        $email = $TestRecipient
        if (-not $IsTest) { $email = ([string]$row[$ui.RecipientColumn]).Trim() }

        $message = $groupMessage
        if ($null -eq $message) {
            $values = $null
            if ($null -ne $row) { $values = ConvertTo-TbmRowDictionary -Row $row -ColumnNames $ui.ColumnNames }
            $expanded = Expand-TbmTemplate -Template $template -Context $context -Row $values -HtmlEncodeValues:($format -eq 'Html')
            $message = $expanded.Text
            if ($null -ne $row -and $expanded.MissingTokens.Count -gt 0) {
                $messagesWithMissing++
                foreach ($name in $expanded.MissingTokens) {
                    $count = 0
                    [void]$missingCounts.TryGetValue($name, [ref]$count)
                    $missingCounts[$name] = $count + 1
                }
            }
            if ($mode -eq 'Group') { $groupMessage = $message }
        }

        $recipient = [TeamsBlastMessenger.TbmRecipientResult]::new()
        $recipient.RowIndex = -1
        if (-not $IsTest -and $null -ne $row) { $recipient.RowIndex = $ui.Table.Rows.IndexOf($row) }
        $recipient.Email = $email
        $recipient.RenderedMessage = $message
        $blast.Recipients.Add($recipient)
        $items.Add(@{ Index = $blast.Recipients.Count - 1; Email = $email; Message = $message })
    }

    if ($messagesWithMissing -gt 0) {
        $parts = [System.Collections.Generic.List[string]]::new()
        foreach ($pair in $missingCounts.GetEnumerator()) { $parts.Add(('{{{0}}} empty in {1} row(s)' -f $pair.Key, $pair.Value)) }
        $text = "{0} message(s) have blank values:`n  {1}`n`nThose spots will be left empty. Continue?" -f $messagesWithMissing, [string]::Join("`n  ", $parts)
        if (-not (Confirm-TbmAction -Text $text -Icon ([System.Windows.MessageBoxImage]::Warning))) { return $null }
    }

    if ($transport -like 'DeepLink*') {
        $maxLength = [int]$ui.Settings.DeepLinkMaxUrlLength
        $tooLong = 0
        if ($mode -eq 'Group') {
            $emails = [System.Collections.Generic.List[string]]::new()
            foreach ($item in $items) { $emails.Add($item.Email) }
            $plain = ConvertTo-TbmPlainText -Message ([string]$items[0].Message) -Format $format
            if ((New-TbmDeepLink -Emails $emails.ToArray() -Message $plain -Topic $blast.Topic).Length -gt $maxLength) { $tooLong = 1 }
        }
        else {
            foreach ($item in $items) {
                $plain = ConvertTo-TbmPlainText -Message ([string]$item.Message) -Format $format
                if ((New-TbmDeepLink -Emails @([string]$item.Email) -Message $plain).Length -gt $maxLength) { $tooLong++ }
            }
        }
        if ($tooLong -gt 0) {
            $text = "{0} deep link(s) are longer than {1} characters. Teams may cut the message short or not open the chat. Shorten the message (or connect to Graph) for reliable delivery.`n`nContinue anyway?" -f $tooLong, $maxLength
            if (-not (Confirm-TbmAction -Text $text -Icon ([System.Windows.MessageBoxImage]::Warning))) { return $null }
        }
        if ($format -eq 'Html') { Write-TbmLog -Level Info -Source 'Send' -Message 'Deep links support plain text only; HTML formatting will be removed.' }
    }

    return @{ Blast = $blast; Items = $items }
}

function Start-TbmSendFromUi {
    [CmdletBinding()]
    param()

    $ui = $script:TbmUi
    if ($ui.Busy) { return }
    $plan = New-TbmSendPlan
    if ($null -eq $plan) { return }
    $blast = $plan.Blast
    $label = Get-TbmTransportLabel -Transport $blast.Transport

    if ($blast.Mode -eq 'Group') {
        $topic = ''
        if (-not [string]::IsNullOrEmpty($blast.Topic)) { $topic = ' named "{0}"' -f $blast.Topic }
        $text = 'Create one group chat{0} with {1} people and send the message via {2}?' -f $topic, $blast.RecipientCount, $label
    }
    else {
        $text = 'Send {0} individual message(s) via {1}?' -f $blast.RecipientCount, $label
    }
    switch ($blast.Transport) {
        { $_ -like 'Graph*' } { $text += "`n`nMessages are sent as {0} ({1})." -f $ui.AuthState.MeName, $ui.AuthState.MeUpn }
        'DeepLinkAuto' { $text += "`n`nTeams opens each chat and Enter is pressed after {0} seconds. Don't use the keyboard or mouse until it finishes (Cancel stops it)." -f $ui.Settings.DeepLinkSendDelaySeconds }
        'DeepLinkManual' { $text += "`n`nEach chat opens in Teams with the message filled in. Check it, press Send in Teams, then click 'Sent – Next' here." }
    }
    if (-not (Confirm-TbmAction -Text $text)) { return }
    Start-TbmBlastRun -Plan $plan
}

function Start-TbmTestSend {
    [CmdletBinding()]
    param()

    $ui = $script:TbmUi
    if ($ui.Busy) { return }
    $recipient = ([string]$ui.Settings.TestRecipient).Trim()
    if ([string]::IsNullOrEmpty($recipient)) { $recipient = [string]$ui.AuthState.MeUpn }
    if ([string]::IsNullOrEmpty($recipient)) {
        throw [System.InvalidOperationException]::new('Set a test recipient in Settings, or connect to Graph to send the test to yourself.')
    }
    $plan = New-TbmSendPlan -IsTest -TestRow (Get-TbmUiPreviewRow) -TestRecipient $recipient
    if ($null -eq $plan) { return }
    $text = 'Send the previewed message to {0} as a test via {1}?' -f $recipient, (Get-TbmTransportLabel -Transport $plan.Blast.Transport)
    if (-not (Confirm-TbmAction -Text $text)) { return }
    Start-TbmBlastRun -Plan $plan
}

function Start-TbmRetry {
    <#
    .SYNOPSIS
        Re-sends a previous blast's failed/cancelled recipients with their original rendered messages.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [TeamsBlastMessenger.TbmBlastRecord]$SourceBlast
    )

    $ui = $script:TbmUi
    if ($ui.Busy) { return }
    $recipients = Get-TbmRetryableRecipients -Blast $SourceBlast
    if ($recipients.Count -eq 0) { return }
    $transport = Resolve-TbmUiTransport

    $blast = [TeamsBlastMessenger.TbmBlastRecord]::new()
    $blast.Mode = 'Individual'
    $blast.Transport = $transport
    $blast.Template = $SourceBlast.Template
    $blast.Format = $SourceBlast.Format
    $blast.SourceFile = $SourceBlast.SourceFile
    if ($transport -like 'Graph*') {
        $blast.SenderId = [string]$ui.AuthState.MeId
        $blast.SenderUpn = [string]$ui.AuthState.MeUpn
    }
    $items = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($source in $recipients) {
        $recipient = [TeamsBlastMessenger.TbmRecipientResult]::new()
        $recipient.RowIndex = -1
        $recipient.Email = $source.Email
        $recipient.RenderedMessage = $source.RenderedMessage
        $blast.Recipients.Add($recipient)
        $items.Add(@{ Index = $blast.Recipients.Count - 1; Email = $source.Email; Message = [string]$source.RenderedMessage })
    }
    $text = 'Retry {0} recipient(s) individually via {1}, using the messages originally rendered for them?' -f $items.Count, (Get-TbmTransportLabel -Transport $transport)
    if (-not (Confirm-TbmAction -Text $text)) { return }
    Start-TbmBlastRun -Plan @{ Blast = $blast; Items = $items }
}

function Start-TbmBlastRun {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Plan
    )

    $ui = $script:TbmUi
    $blast = [TeamsBlastMessenger.TbmBlastRecord]$Plan.Blast
    if ($blast.Transport -like 'Graph*') { [void](Get-TbmAccessToken -AuthState $ui.AuthState) }

    $sync = New-TbmSyncState
    $job = @{ Blast = $blast; Items = $Plan.Items; Settings = $ui.Settings.Clone(); AuthState = $ui.AuthState; Sync = $sync }
    $ui.Sync = $sync
    $ui.Blast = $blast
    $ui.Progress = @{ Done = 0; Sent = 0; Failed = 0; Total = $blast.RecipientCount }
    $ui.C.PrgSend.Maximum = [System.Math]::Max(1, $blast.RecipientCount)
    $ui.C.PrgSend.Value = 0
    foreach ($recipient in $blast.Recipients) { Set-TbmGridStatus -RowIndex $recipient.RowIndex -Status 'Queued' }

    Set-TbmUiBusy -Busy $true
    if ($blast.Transport -eq 'DeepLinkManual') {
        Start-TbmManualSession -Job $job
    }
    else {
        $ui.Job = Start-TbmBackgroundJob -CommandName 'Invoke-TbmSendWorker' -Job $job
        $ui.C.TxtProgress.Text = 'Sending to {0} recipient(s)…' -f $blast.RecipientCount
    }
}

function Stop-TbmCurrentSend {
    [CmdletBinding()]
    param()
    $ui = $script:TbmUi
    if ($null -ne $ui.Manual) {
        Stop-TbmManualSession
        return
    }
    if ($null -ne $ui.Sync) {
        $ui.Sync['Cancel'] = $true
        $ui.C.BtnCancel.IsEnabled = $false
        $ui.C.TxtProgress.Text = 'Cancelling after the current recipient…'
        Write-TbmLog -Level Info -Source 'Send' -Message 'Cancellation requested.'
    }
}

function Set-TbmGridStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$RowIndex,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Status
    )
    $table = $script:TbmUi.Table
    if ($null -eq $table -or $RowIndex -lt 0 -or $RowIndex -ge $table.Rows.Count) { return }
    $table.Rows[$RowIndex][$script:TbmStatusColumn] = $Status
}

function Set-TbmUiBusy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [bool]$Busy
    )

    $ui = $script:TbmUi
    $ui.Busy = $Busy
    foreach ($name in @('BtnBrowse', 'BtnSend', 'BtnTest', 'BtnPreviewAll', 'BtnConnect', 'BtnDisconnect', 'BtnSettings', 'BtnHistory', 'BtnInteractions', 'RadIndividual', 'RadGroup')) {
        $ui.C[$name].IsEnabled = -not $Busy
    }
    # Boxes locked by organization policy stay locked after a send.
    $ui.C.CmbTransport.IsEnabled = (-not $Busy) -and -not (Test-TbmSettingEnforced -Name 'Transport')
    $ui.C.CmbFormat.IsEnabled = (-not $Busy) -and -not (Test-TbmSettingEnforced -Name 'MessageFormat')
    $ui.C.TxtTopic.IsEnabled = (-not $Busy) -and [bool]$ui.C.RadGroup.IsChecked
    $ui.C.TxtMessage.IsReadOnly = $Busy
    $ui.C.BtnCancel.IsEnabled = $Busy
    $manualVisibility = [System.Windows.Visibility]::Collapsed
    if ($Busy -and $null -ne $ui.Manual) { $manualVisibility = [System.Windows.Visibility]::Visible }
    $ui.C.BtnManualNext.Visibility = $manualVisibility
    $ui.C.BtnManualSkip.Visibility = $manualVisibility
    Update-TbmFileControls
}

#endregion

#region Main window: deep link manual mode (UI-driven)

function Start-TbmManualSession {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Job
    )

    $blast = [TeamsBlastMessenger.TbmBlastRecord]$Job.Blast
    Save-TbmBlastRecord -Blast $blast
    $steps = [System.Collections.Generic.List[hashtable]]::new()
    if ($blast.Mode -eq 'Group') {
        $indexes = [System.Collections.Generic.List[int]]::new()
        foreach ($item in $Job.Items) { $indexes.Add([int]$item.Index) }
        $steps.Add(@{ Indexes = $indexes.ToArray(); Message = [string]$Job.Items[0].Message })
    }
    else {
        foreach ($item in $Job.Items) { $steps.Add(@{ Indexes = @([int]$item.Index); Message = [string]$item.Message }) }
    }
    $script:TbmUi.Manual = @{ Job = $Job; Steps = $steps; Position = 0 }
    Write-TbmLog -Level Info -Source 'Send' -Message ('Blast {0} started (manual deep link): {1} step(s).' -f $blast.Id, $steps.Count)
    Set-TbmUiBusy -Busy $true
    Open-TbmManualStep
}

function Open-TbmManualStep {
    [CmdletBinding()]
    param()

    $ui = $script:TbmUi
    $manual = $ui.Manual
    $blast = [TeamsBlastMessenger.TbmBlastRecord]$manual.Job.Blast
    $step = $manual.Steps[$manual.Position]
    $emails = [System.Collections.Generic.List[string]]::new()
    foreach ($index in $step.Indexes) {
        $emails.Add($blast.Recipients[$index].Email)
        Set-TbmGridStatus -RowIndex $blast.Recipients[$index].RowIndex -Status 'Sending'
    }
    $format = 'Text'
    if ($blast.Format -eq 'Html') { $format = 'Html' }
    $plain = ConvertTo-TbmPlainText -Message $step.Message -Format $format
    $uri = New-TbmDeepLink -Emails $emails.ToArray() -Message $plain -Topic $blast.Topic -Target ([string]$manual.Job.Settings.DeepLinkTarget)
    Start-TbmDeepLink -Uri $uri

    $who = $emails[0]
    if ($emails.Count -gt 1) { $who = '{0} people' -f $emails.Count }
    $ui.C.TxtProgress.Text = '{0} of {1}: {2}. Check the message in Teams, press Send there, then click "Sent – Next".' -f ($manual.Position + 1), $manual.Steps.Count, $who
}

function Complete-TbmManualStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Sent', 'Skipped')]
        [string]$Status
    )

    $ui = $script:TbmUi
    $manual = $ui.Manual
    if ($null -eq $manual) { return }
    $blast = [TeamsBlastMessenger.TbmBlastRecord]$manual.Job.Blast
    $step = $manual.Steps[$manual.Position]
    $now = [datetime]::UtcNow
    foreach ($index in $step.Indexes) {
        if ($Status -eq 'Sent') { $blast.Recipients[$index].SentUtc = $now }
        $reason = $null
        if ($Status -eq 'Skipped') { $reason = 'Skipped by user.' }
        Set-TbmRecipientOutcome -Job $manual.Job -Index $index -Status $Status -ErrorMessage $reason
    }
    $manual.Position++
    if ($manual.Position -ge $manual.Steps.Count) {
        Stop-TbmManualSession
        return
    }
    Open-TbmManualStep
}

function Stop-TbmManualSession {
    [CmdletBinding()]
    param()

    $ui = $script:TbmUi
    $manual = $ui.Manual
    if ($null -eq $manual) { return }
    $blast = [TeamsBlastMessenger.TbmBlastRecord]$manual.Job.Blast
    foreach ($item in $manual.Job.Items) {
        if ($blast.Recipients[$item.Index].Status -eq [TeamsBlastMessenger.TbmRecipientStatus]::Pending) {
            Set-TbmRecipientOutcome -Job $manual.Job -Index $item.Index -Status Cancelled -ErrorMessage 'Cancelled before sending.'
        }
    }
    $blast.CompletedUtc = [datetime]::UtcNow
    Save-TbmBlastRecord -Blast $blast
    Write-TbmLog -Level Info -Source 'Send' -Message ('Blast {0} finished: {1} sent, {2} failed, {3} total.' -f $blast.Id, $blast.SentCount, $blast.FailedCount, $blast.RecipientCount)
    Add-TbmSyncEvent -Sync $manual.Job.Sync -Type Done -Data @{ BlastId = $blast.Id; Sent = $blast.SentCount; Failed = $blast.FailedCount; Total = $blast.RecipientCount }
}

#endregion

#region Main window: timer, dry run, history and interactions

function Invoke-TbmMainTimerTick {
    [CmdletBinding()]
    param()

    $ui = $script:TbmUi
    if ($ui.InTick) { return }
    $ui.InTick = $true
    try {
        Update-TbmLogPane
        if ($null -eq $ui.Sync) { return }
        $item = $null
        while ($null -ne $ui.Sync -and $ui.Sync.Events.TryDequeue([ref]$item)) {
            switch ($item.Type) {
                'Recipient' {
                    Set-TbmGridStatus -RowIndex ([int]$item.RowIndex) -Status ([string]$item.Status)
                    $ui.Progress.Done++
                    if ($item.Status -eq 'Sent') { $ui.Progress.Sent++ }
                    if ($item.Status -eq 'Failed') { $ui.Progress.Failed++ }
                    $ui.C.PrgSend.Value = [System.Math]::Min($ui.C.PrgSend.Maximum, $ui.Progress.Done)
                    if ($null -eq $ui.Manual) {
                        $ui.C.TxtProgress.Text = '{0} of {1} processed · {2} sent · {3} failed' -f $ui.Progress.Done, $ui.Progress.Total, $ui.Progress.Sent, $ui.Progress.Failed
                    }
                }
                'Done' { Complete-TbmRunUi -DoneEvent $item }
            }
        }
    }
    catch {
        Write-TbmLog -Source 'UI' -Message 'Timer update failed.' -ErrorRecord $_
    }
    finally {
        $ui.InTick = $false
    }
}

function Complete-TbmRunUi {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [object]$DoneEvent
    )

    $ui = $script:TbmUi
    if ($null -ne $ui.Job) {
        Complete-TbmBackgroundJob -Handle $ui.Job
        $ui.Job = $null
    }
    $blast = $ui.Blast
    $ui.Sync = $null
    $ui.Manual = $null
    if ($null -ne $blast -and -not $blast.IsTest) { $ui.LastBlast = $blast }
    Set-TbmUiBusy -Busy $false
    $ui.C.PrgSend.Value = $ui.C.PrgSend.Maximum
    $prefix = 'Finished'
    if ($null -ne $blast -and $blast.IsTest) { $prefix = 'Test finished' }
    $ui.C.TxtProgress.Text = '{0}: {1} sent, {2} failed, {3} total. Saved to History.' -f $prefix, $DoneEvent.Sent, $DoneEvent.Failed, $DoneEvent.Total
}

function Update-TbmLogPane {
    [CmdletBinding()]
    param()
    $entries = [TeamsBlastMessenger.TbmLogger]::DrainPending()
    if ($entries.Count -eq 0) { return }
    $ui = $script:TbmUi
    foreach ($entry in $entries) {
        $line = $entry.ToDisplayLine()
        $newline = $line.IndexOf("`n")
        if ($newline -gt 0) { $line = $line.Substring(0, $newline).TrimEnd() + ' (details in log file)' }
        $ui.LogLines.Enqueue($line)
    }
    while ($ui.LogLines.Count -gt 800) { [void]$ui.LogLines.Dequeue() }
    $ui.C.TxtLog.Text = [string]::Join([System.Environment]::NewLine, $ui.LogLines.ToArray())
    $ui.C.TxtLog.ScrollToEnd()
}

function Invoke-TbmDryRun {
    <#
    .SYNOPSIS
        Renders every selected message without sending and marks rows Ready / Missing.
    #>
    [CmdletBinding()]
    param()

    $ui = $script:TbmUi
    if ($null -eq $ui.Table) { throw [System.InvalidOperationException]::new('Load a recipient spreadsheet first.') }
    if ($null -eq $ui.RecipientColumn) { throw [System.InvalidOperationException]::new('Choose the recipient (email) column first.') }
    $template = $ui.C.TxtMessage.Text
    $format = Get-TbmUiFormat
    $context = New-TbmUiTokenContext
    $check = Test-TbmTemplate -Template $template -Context $context
    $rows = Get-TbmIncludedRows
    $ready = 0
    $missing = 0
    $longest = 0
    foreach ($row in $rows) {
        $values = ConvertTo-TbmRowDictionary -Row $row -ColumnNames $ui.ColumnNames
        $expanded = Expand-TbmTemplate -Template $template -Context $context -Row $values -HtmlEncodeValues:($format -eq 'Html')
        $longest = [System.Math]::Max($longest, $expanded.Text.Length)
        if ($expanded.MissingTokens.Count -gt 0) {
            $missing++
            $row[$script:TbmStatusColumn] = 'Missing: ' + [string]::Join(', ', $expanded.MissingTokens)
        }
        else {
            $ready++
            $row[$script:TbmStatusColumn] = 'Ready'
        }
    }
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add(('{0} message(s) rendered; nothing was sent.' -f $rows.Count))
    $lines.Add(('  Ready: {0}' -f $ready))
    $lines.Add(('  With blank values: {0}' -f $missing))
    if ($check.UnknownTokens.Count -gt 0) { $lines.Add(('  Unknown tokens (sent as typed): {0}' -f [string]::Join(', ', (ConvertTo-TbmTokenDisplay -Names $check.UnknownTokens)))) }
    $lines.Add(('  Longest message: {0} characters' -f $longest))
    $lines.Add(('  Would send via: {0}' -f (Get-TbmTransportLabel -Transport (Resolve-TbmUiTransport))))
    Write-TbmLog -Level Info -Source 'DryRun' -Message ('Dry run: {0} ready, {1} with blank values.' -f $ready, $missing)
    [void](Show-TbmMessageBox -Text ([string]::Join("`n", $lines)) -Title 'Dry run')
}

function Show-TbmHistoryFromUi {
    [CmdletBinding()]
    param()
    $ui = $script:TbmUi
    $choice = Show-TbmHistoryWindow -Owner $ui.Window
    if ($null -eq $choice) { return }
    $blast = [TeamsBlastMessenger.TbmBlastRecord]$choice.Blast
    switch ($choice.Action) {
        'Reuse' {
            $ui.C.TxtMessage.Text = $blast.Template
            Set-TbmComboByTag -ComboBox $ui.C.CmbFormat -Value $blast.Format
            if ($blast.Mode -eq 'Group') { $ui.C.RadGroup.IsChecked = $true } else { $ui.C.RadIndividual.IsChecked = $true }
            $ui.C.TxtTopic.Text = [string]$blast.Topic
            Update-TbmPreview
        }
        'Retry' { Start-TbmRetry -SourceBlast $blast }
    }
}

function Open-TbmLatestInteractions {
    [CmdletBinding()]
    param()
    $ui = $script:TbmUi
    $blast = $ui.LastBlast
    if ($null -eq $blast) {
        foreach ($record in (Get-TbmBlastHistory)) {
            if (-not $record.IsTest -and $record.SentCount -gt 0) { $blast = $record; break }
        }
    }
    if ($null -eq $blast) {
        [void](Show-TbmMessageBox -Text 'There are no sent blasts yet. Send one first, or open History to pick an older blast.')
        return
    }
    Open-TbmInteractions -Blast $blast -Owner $ui.Window
}

#endregion
