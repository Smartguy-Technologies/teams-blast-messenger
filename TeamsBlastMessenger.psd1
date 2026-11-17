@{
    RootModule           = 'TeamsBlastMessenger.psm1'
    ModuleVersion        = '1.0.0'
    GUID                 = '7d0f5a8e-3c4b-4f3e-9a51-2b6c1e9f4d27'
    Author               = 'Teams Blast Messenger'
    Description          = 'Send personalized Microsoft Teams messages to many recipients from a spreadsheet, via Microsoft Graph or a deep-link fallback, with send history and reply tracking.'
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    FunctionsToExport    = @(
        'Show-TbmMainWindow',
        'Invoke-TbmSendWorker',
        'Invoke-TbmInteractionWorker'
    )
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()
    PrivateData          = @{
        PSData = @{
            Tags = @('Teams', 'Graph', 'WPF', 'Messaging')
        }
    }
}
