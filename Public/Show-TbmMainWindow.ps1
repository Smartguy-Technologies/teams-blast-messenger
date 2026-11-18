function Show-TbmMainWindow {
    <#
    .SYNOPSIS
        Opens the Teams Blast Messenger main window (blocks until it is closed).
    .DESCRIPTION
        Requires Windows and an STA thread (Windows PowerShell 5.1 console is STA by default;
        Start-TeamsBlastMessenger.ps1 relaunches with -STA when needed).
    .EXAMPLE
        Import-Module .\TeamsBlastMessenger.psd1; Show-TbmMainWindow
    #>
    [CmdletBinding()]
    param()

    if (-not (Test-TbmIsWindows)) { throw [System.PlatformNotSupportedException]::new('Teams Blast Messenger requires Windows (WPF).') }
    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
        throw [System.InvalidOperationException]::new('WPF requires an STA thread. Start with: powershell.exe -STA -File .\Start-TeamsBlastMessenger.ps1')
    }

    $settings = Initialize-TbmEnvironment
    Initialize-TbmWpf
    $ui = Get-TbmXamlWindow -Name 'MainWindow'

    $script:TbmUi = @{
        Window          = $ui.Window
        C               = $ui.Controls
        Settings        = $settings
        AuthState       = New-TbmAuthState
        Table           = $null
        SourceFile      = $null
        SourceLabel     = ''
        ManualCount     = 0
        SheetName       = $null
        ColumnNames     = [string[]]@()
        RecipientColumn = $null
        Busy            = $false
        Job             = $null
        Sync            = $null
        Blast           = $null
        Manual          = $null
        LastBlast       = $null
        Suppress        = $false
        InTick          = $false
        Progress        = @{ Done = 0; Sent = 0; Failed = 0; Total = 0 }
        LogLines        = [System.Collections.Generic.Queue[string]]::new()
        Timer           = $null
    }

    Register-TbmMainWindowEvents
    $script:TbmUi.Timer = New-TbmDispatcherTimer -IntervalMilliseconds 250 -OnTick { Invoke-TbmMainTimerTick }
    $script:TbmUi.Timer.Start()

    try {
        $null = $ui.Window.ShowDialog()
    }
    finally {
        $script:TbmUi.Timer.Stop()
        Write-TbmLog -Level Info -Source 'App' -Message 'Teams Blast Messenger closed.'
    }
}
