#region WPF infrastructure

function Initialize-TbmWpf {
    <#
    .SYNOPSIS
        Loads WPF assemblies and ensures an Application object holds the shared styles (StaticResource lookups).
    #>
    [CmdletBinding()]
    param()

    foreach ($assemblyName in @('PresentationFramework', 'PresentationCore', 'WindowsBase', 'System.Xaml', 'System.Windows.Forms')) {
        Add-Type -AssemblyName $assemblyName
    }

    $application = [System.Windows.Application]::Current
    if ($null -eq $application) {
        $application = [System.Windows.Application]::new()
        $application.ShutdownMode = [System.Windows.ShutdownMode]::OnExplicitShutdown
    }
    $stylesPath = [System.IO.Path]::Combine($script:TbmModuleRoot, 'Xaml', 'Styles.xaml')
    $alreadyLoaded = $false
    foreach ($dictionary in $application.Resources.MergedDictionaries) {
        if ($dictionary.Contains('Accent')) { $alreadyLoaded = $true; break }
    }
    if (-not $alreadyLoaded) {
        $application.Resources.MergedDictionaries.Add((ConvertFrom-TbmXamlFile -Path $stylesPath))
    }
}

function ConvertFrom-TbmXamlFile {
    [CmdletBinding()]
    [OutputType([object])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    $document = [System.Xml.XmlDocument]::new()
    $document.XmlResolver = $null
    $document.LoadXml([System.IO.File]::ReadAllText($Path))
    $reader = [System.Xml.XmlNodeReader]::new($document)
    try {
        return [System.Windows.Markup.XamlReader]::Load($reader)
    }
    finally {
        $reader.Dispose()
    }
}

function Get-TbmXamlWindow {
    <#
    .SYNOPSIS
        Loads Xaml\<Name>.xaml and returns the Window plus a hashtable of every element that has a Name.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('MainWindow', 'HistoryWindow', 'InteractionsWindow', 'SettingsWindow', 'DeviceCodeWindow', 'ManualRecipientsWindow')]
        [string]$Name,

        [Parameter()]
        [System.Windows.Window]$Owner
    )

    $path = [System.IO.Path]::Combine($script:TbmModuleRoot, 'Xaml', $Name + '.xaml')
    $document = [System.Xml.XmlDocument]::new()
    $document.XmlResolver = $null
    $document.LoadXml([System.IO.File]::ReadAllText($path))
    $reader = [System.Xml.XmlNodeReader]::new($document)
    try {
        $window = [System.Windows.Markup.XamlReader]::Load($reader)
    }
    finally {
        $reader.Dispose()
    }

    $controls = @{}
    foreach ($node in $document.SelectNodes('//*[@Name]')) {
        $elementName = $node.GetAttribute('Name')
        $element = $window.FindName($elementName)
        if ($null -ne $element) { $controls[$elementName] = $element }
    }
    if ($null -ne $Owner) { $window.Owner = $Owner }
    return @{ Window = $window; Controls = $controls }
}

function Get-TbmActiveWindow {
    [CmdletBinding()]
    [OutputType([System.Windows.Window])]
    param()
    $application = [System.Windows.Application]::Current
    if ($null -ne $application) {
        foreach ($window in $application.Windows) { if ($window.IsActive) { return $window } }
    }
    if ($null -ne $script:TbmUi -and $null -ne $script:TbmUi.Window) { return $script:TbmUi.Window }
    return $null
}

function Show-TbmMessageBox {
    [CmdletBinding()]
    [OutputType([System.Windows.MessageBoxResult])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Text,

        [Parameter()]
        [string]$Title = 'Teams Blast Messenger',

        [Parameter()]
        [System.Windows.MessageBoxButton]$Buttons = [System.Windows.MessageBoxButton]::OK,

        [Parameter()]
        [System.Windows.MessageBoxImage]$Icon = [System.Windows.MessageBoxImage]::Information,

        [Parameter()]
        [System.Windows.Window]$Owner
    )

    if ($null -eq $Owner) { $Owner = Get-TbmActiveWindow }
    if ($null -ne $Owner) { return [System.Windows.MessageBox]::Show($Owner, $Text, $Title, $Buttons, $Icon) }
    return [System.Windows.MessageBox]::Show($Text, $Title, $Buttons, $Icon)
}

function Confirm-TbmAction {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Text,

        [Parameter()]
        [System.Windows.MessageBoxImage]$Icon = [System.Windows.MessageBoxImage]::Question
    )
    $answer = Show-TbmMessageBox -Text $Text -Buttons ([System.Windows.MessageBoxButton]::YesNo) -Icon $Icon
    return $answer -eq [System.Windows.MessageBoxResult]::Yes
}

function Invoke-TbmUiAction {
    <#
    .SYNOPSIS
        Runs a UI event handler body; any exception is logged with detail and shown as a friendly message.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [scriptblock]$Action
    )

    try {
        & $Action
    }
    catch {
        Write-TbmLog -Source 'UI' -Message ('{0} failed.' -f $Name) -ErrorRecord $_
        [void](Show-TbmMessageBox -Text ("{0} failed:`n`n{1}" -f $Name, $_.Exception.Message) -Icon ([System.Windows.MessageBoxImage]::Error))
    }
}

function Set-TbmComboByTag {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Windows.Controls.ComboBox]$ComboBox,

        [Parameter()]
        [AllowEmptyString()]
        [string]$Value
    )
    foreach ($item in $ComboBox.Items) {
        $key = [string]$item.Tag
        if ([string]::IsNullOrEmpty($key)) { $key = [string]$item.Content }
        if ([string]::Equals($key, $Value, [System.StringComparison]::OrdinalIgnoreCase)) { $ComboBox.SelectedItem = $item; return }
    }
    if ($ComboBox.Items.Count -gt 0 -and $ComboBox.SelectedIndex -lt 0) { $ComboBox.SelectedIndex = 0 }
}

function Get-TbmComboTag {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Windows.Controls.ComboBox]$ComboBox
    )
    $item = $ComboBox.SelectedItem
    if ($null -eq $item) { return '' }
    if ($item -is [System.Windows.Controls.ComboBoxItem]) {
        if (-not [string]::IsNullOrEmpty([string]$item.Tag)) { return [string]$item.Tag }
        return [string]$item.Content
    }
    return [string]$item
}

function Show-TbmSaveFileDialog {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$FileName,

        [Parameter()]
        [string]$Filter = 'CSV file (*.csv)|*.csv',

        [Parameter()]
        [System.Windows.Window]$Owner
    )
    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.FileName = $FileName
    $dialog.Filter = $Filter
    $dialog.InitialDirectory = Get-TbmAppDataPath -Name 'Exports'
    $dialog.OverwritePrompt = $true
    if ($null -eq $Owner) { $Owner = Get-TbmActiveWindow }
    if ($dialog.ShowDialog($Owner) -eq $true) { return $dialog.FileName }
    return $null
}

function Open-TbmFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )
    [void][System.IO.Directory]::CreateDirectory($Path)
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new('explorer.exe', ('"{0}"' -f $Path))
    $startInfo.UseShellExecute = $true
    $process = [System.Diagnostics.Process]::Start($startInfo)
    if ($null -ne $process) { $process.Dispose() }
}

function New-TbmDispatcherTimer {
    [CmdletBinding()]
    [OutputType([System.Windows.Threading.DispatcherTimer])]
    param(
        [Parameter(Mandatory)]
        [ValidateRange(50, 60000)]
        [int]$IntervalMilliseconds,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [scriptblock]$OnTick
    )
    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [System.TimeSpan]::FromMilliseconds($IntervalMilliseconds)
    $timer.Add_Tick($OnTick)
    return $timer
}

function Wait-TbmBackgroundJob {
    <#
    .SYNOPSIS
        Waits (bounded) for a background job after cancellation, then disposes it.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Handle,

        [Parameter()]
        [ValidateRange(0, 120000)]
        [int]$TimeoutMilliseconds = 20000
    )
    if ($Handle.AsyncResult.AsyncWaitHandle.WaitOne($TimeoutMilliseconds)) {
        Complete-TbmBackgroundJob -Handle $Handle
    }
    else {
        Write-TbmLog -Level Warn -Source 'Jobs' -Message ('{0} did not stop within {1} ms; abandoning it.' -f $Handle.CommandName, $TimeoutMilliseconds)
        try { $Handle.PowerShell.Stop() } catch { Write-TbmLog -Level Debug -Source 'Jobs' -Message 'Stop failed.' }
    }
}

#endregion

#region Settings dialog

function Show-TbmSettingsWindow {
    <#
    .SYNOPSIS
        Edits settings in place. Returns $true when saved.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Windows.Window]$Owner,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Settings
    )

    $ui = Get-TbmXamlWindow -Name 'SettingsWindow' -Owner $Owner
    $settingsView = @{ Ui = $ui; Settings = $Settings; Saved = $false }
    $c = $ui.Controls

    Set-TbmComboByTag -ComboBox $c.CmbAuthMode -Value ([string]$Settings.AuthMode)
    $c.TxtClientId.Text = [string]$Settings.ClientId
    $c.TxtTenantId.Text = [string]$Settings.TenantId
    $c.TxtPacing.Text = [string]$Settings.PacingDelayMs
    $c.TxtTestRecipient.Text = [string]$Settings.TestRecipient
    Set-TbmComboByTag -ComboBox $c.CmbDeepLinkTarget -Value ([string]$Settings.DeepLinkTarget)
    $c.TxtDeepLinkDelay.Text = [string]$Settings.DeepLinkSendDelaySeconds
    $c.TxtMaxUrl.Text = [string]$Settings.DeepLinkMaxUrlLength
    Set-TbmComboByTag -ComboBox $c.CmbLogLevel -Value ([string]$Settings.LogLevel)
    $c.TxtRetention.Text = [string]$Settings.LogRetentionDays
    $c.TxtChatPages.Text = [string]$Settings.InteractionMaxChatPages

    $sdkText = 'Graph SDK module: not installed (Install-Module Microsoft.Graph.Authentication -Scope CurrentUser).'
    $sdkModules = Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication'
    if ($null -ne $sdkModules -and @($sdkModules).Count -gt 0) { $sdkText = 'Graph SDK module: installed (v{0}).' -f @($sdkModules)[0].Version }
    $c.TxtGraphInfo.Text = $sdkText + "`nApp registration: public client, 'Allow public client flows' = Yes, delegated permissions " + [string]::Join(', ', $script:TbmGraphScopes) + '.'
    $c.TxtAppData.Text = 'Logs, history and settings are stored in ' + (Get-TbmAppDataPath -Name 'Root')

    $c.BtnOpenAppData.Add_Click({ Invoke-TbmUiAction -Name 'Open data folder' -Action { Open-TbmFolder -Path (Get-TbmAppDataPath -Name 'Root') } })
    $c.BtnSave.Add_Click({
            Invoke-TbmUiAction -Name 'Save settings' -Action {
                $controls = $settingsView.Ui.Controls
                $errors = [System.Collections.Generic.List[string]]::new()
                $values = @{}
                foreach ($spec in @(
                        @{ Name = 'PacingDelayMs'; Box = 'TxtPacing'; Min = 0; Max = 60000; Label = 'Delay between messages' },
                        @{ Name = 'DeepLinkSendDelaySeconds'; Box = 'TxtDeepLinkDelay'; Min = 1; Max = 120; Label = 'Auto-send wait' },
                        @{ Name = 'DeepLinkMaxUrlLength'; Box = 'TxtMaxUrl'; Min = 500; Max = 32000; Label = 'Max deep-link length' },
                        @{ Name = 'LogRetentionDays'; Box = 'TxtRetention'; Min = 1; Max = 3650; Label = 'Keep logs' },
                        @{ Name = 'InteractionMaxChatPages'; Box = 'TxtChatPages'; Min = 1; Max = 400; Label = 'Chat pages to search' })) {
                    $number = 0
                    if (-not [int]::TryParse($controls[$spec.Box].Text.Trim(), [ref]$number) -or $number -lt $spec.Min -or $number -gt $spec.Max) {
                        $errors.Add(('{0} must be a whole number from {1} to {2}.' -f $spec.Label, $spec.Min, $spec.Max))
                    }
                    else { $values[$spec.Name] = $number }
                }
                $clientId = $controls.TxtClientId.Text.Trim()
                $parsedGuid = [guid]::Empty
                if (-not [string]::IsNullOrEmpty($clientId) -and -not [guid]::TryParse($clientId, [ref]$parsedGuid)) { $errors.Add('Client ID must be a GUID (Application ID from the app registration).') }
                $testRecipient = $controls.TxtTestRecipient.Text.Trim()
                if (-not [string]::IsNullOrEmpty($testRecipient) -and -not (Test-TbmEmailAddress -Value $testRecipient)) { $errors.Add('Test recipient must be an email address.') }
                if ($errors.Count -gt 0) {
                    [void](Show-TbmMessageBox -Text ([string]::Join("`n", $errors)) -Icon ([System.Windows.MessageBoxImage]::Warning) -Owner $settingsView.Ui.Window)
                    return
                }

                $target = $settingsView.Settings
                foreach ($key in $values.Keys) { $target[$key] = $values[$key] }
                $target.AuthMode = Get-TbmComboTag -ComboBox $controls.CmbAuthMode
                $target.ClientId = $clientId
                $tenant = $controls.TxtTenantId.Text.Trim()
                if ([string]::IsNullOrEmpty($tenant)) { $tenant = 'organizations' }
                $target.TenantId = $tenant
                $target.TestRecipient = $testRecipient
                $target.DeepLinkTarget = Get-TbmComboTag -ComboBox $controls.CmbDeepLinkTarget
                $target.LogLevel = Get-TbmComboTag -ComboBox $controls.CmbLogLevel
                Save-TbmSettings -Settings $target

                $level = [TeamsBlastMessenger.TbmLogLevel]::Info
                [void][System.Enum]::TryParse([string]$target.LogLevel, $true, [ref]$level)
                [TeamsBlastMessenger.TbmLogger]::MinimumLevel = $level
                $settingsView.Saved = $true
                $settingsView.Ui.Window.DialogResult = $true
            }
        })

    [void]$ui.Window.ShowDialog()
    return [bool]$settingsView.Saved
}

#endregion

#region Device code sign-in dialog

function Show-TbmDeviceCodeWindow {
    <#
    .SYNOPSIS
        Runs the device code flow with a small dialog. Returns $true when signed in.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Windows.Window]$Owner,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Settings,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState
    )

    $deviceCode = Start-TbmDeviceCodeLogin -Settings $Settings
    $ui = Get-TbmXamlWindow -Name 'DeviceCodeWindow' -Owner $Owner
    $deviceView = @{
        Ui        = $ui
        Code      = $deviceCode
        AuthState = $AuthState
        Interval  = $deviceCode.IntervalSeconds
        NextPoll  = [datetime]::UtcNow.AddSeconds($deviceCode.IntervalSeconds)
        Success   = $false
        Timer     = $null
    }
    $ui.Controls.TxtInstructions.Text = ('Open {0} and enter the code below. Sign in with the account you send Teams messages from, and approve the requested permissions ({1}).' -f $deviceCode.VerificationUri, [string]::Join(', ', $script:TbmGraphScopes))
    $ui.Controls.TxtUserCode.Text = $deviceCode.UserCode

    $ui.Controls.BtnCopyOpen.Add_Click({
            Invoke-TbmUiAction -Name 'Open sign-in page' -Action {
                [System.Windows.Clipboard]::SetText($deviceView.Code.UserCode)
                Start-TbmDeepLink -Uri $deviceView.Code.VerificationUri
                $deviceView.Ui.Controls.TxtStatus.Text = 'Code copied. Paste it on the sign-in page, then finish signing in…'
            }
        })

    $deviceView.Timer = New-TbmDispatcherTimer -IntervalMilliseconds 1000 -OnTick {
        if ([datetime]::UtcNow -lt $deviceView.NextPoll) { return }
        try {
            $state = Request-TbmDeviceCodeToken -AuthState $deviceView.AuthState -DeviceCode $deviceView.Code
            switch ($state) {
                'Success' {
                    $deviceView.Timer.Stop()
                    $deviceView.Success = $true
                    $deviceView.Ui.Window.DialogResult = $true
                }
                'SlowDown' { $deviceView.Interval += 5 }
            }
            $deviceView.NextPoll = [datetime]::UtcNow.AddSeconds($deviceView.Interval)
        }
        catch {
            $deviceView.Timer.Stop()
            Write-TbmLog -Source 'Auth' -Message 'Device code sign-in failed.' -ErrorRecord $_
            $deviceView.Ui.Controls.TxtStatus.Foreground = [System.Windows.Media.Brushes]::Firebrick
            $deviceView.Ui.Controls.TxtStatus.Text = $_.Exception.Message
        }
    }
    $ui.Window.Add_Closed({ $deviceView.Timer.Stop() })
    $deviceView.Timer.Start()
    [void]$ui.Window.ShowDialog()
    return [bool]$deviceView.Success
}

#endregion

#region Manual recipients dialog

function Show-TbmManualRecipientsWindow {
    <#
    .SYNOPSIS
        Lets the user type or paste recipients. Returns @{ Entries; Replace } for the valid entries, or $null if cancelled.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Windows.Window]$Owner,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter()]
        [bool]$HasList = $false
    )

    $ui = Get-TbmXamlWindow -Name 'ManualRecipientsWindow' -Owner $Owner
    $manualView = @{ Ui = $ui; AuthState = $AuthState; Entries = $null; Result = $null }
    $c = $ui.Controls

    $connected = Test-TbmGraphConnected -AuthState $AuthState
    $c.ChkLookup.IsEnabled = $connected
    $c.ChkLookup.IsChecked = $connected
    if (-not $connected) { $c.ChkLookup.ToolTip = 'Connect to Microsoft Graph to look up people by name.' }
    if (-not $HasList) {
        $c.RadAppend.IsEnabled = $false
        $c.RadReplace.IsChecked = $true
    }

    $c.TxtInput.Add_TextChanged({
            $manualView.Entries = $null
            $manualView.Ui.Controls.BtnAdd.IsEnabled = $false
            $manualView.Ui.Controls.BtnAdd.Content = 'Add recipients'
        })
    $c.BtnCheck.Add_Click({
            Invoke-TbmUiAction -Name 'Check recipients' -Action {
                $controls = $manualView.Ui.Controls
                $entries = ConvertFrom-TbmManualRecipientText -Text $controls.TxtInput.Text
                if ([bool]$controls.ChkLookup.IsChecked -and (Test-TbmGraphConnected -AuthState $manualView.AuthState)) {
                    $manualView.Ui.Window.Cursor = [System.Windows.Input.Cursors]::Wait
                    try { Resolve-TbmManualEntryNames -AuthState $manualView.AuthState -Entries $entries }
                    finally { $manualView.Ui.Window.Cursor = $null }
                }
                $valid = 0
                foreach ($entry in $entries) { if ($entry.IsValid) { $valid++ } }
                $manualView.Entries = $entries
                $controls.GridEntries.ItemsSource = $null
                $controls.GridEntries.ItemsSource = $entries
                $controls.TxtSummary.Text = '{0} entr(ies) found · {1} ready to add · {2} need attention (highlighted; they will not be added).' -f $entries.Count, $valid, ($entries.Count - $valid)
                $controls.BtnAdd.IsEnabled = $valid -gt 0
                $controls.BtnAdd.Content = 'Add {0} recipient(s)' -f $valid
            }
        })
    $c.BtnAdd.Add_Click({
            Invoke-TbmUiAction -Name 'Add recipients' -Action {
                if ($null -eq $manualView.Entries) { return }
                $manualView.Result = @{ Entries = $manualView.Entries; Replace = [bool]$manualView.Ui.Controls.RadReplace.IsChecked }
                $manualView.Ui.Window.DialogResult = $true
            }
        })
    $ui.Window.Add_Loaded({ [void]$manualView.Ui.Controls.TxtInput.Focus() })
    [void]$ui.Window.ShowDialog()
    return $manualView.Result
}

#endregion
