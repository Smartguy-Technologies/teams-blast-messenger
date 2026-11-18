#Requires -Version 5.1
<#
.SYNOPSIS
    Launches Teams Blast Messenger.
.DESCRIPTION
    Ensures the session is STA (required by WPF), imports the module from this folder and opens the main window.
    Data (settings, logs, history) is stored in %APPDATA%\TeamsBlastMessenger.
.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Start-TeamsBlastMessenger.ps1
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    $executable = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $arguments = '-NoProfile -STA -ExecutionPolicy Bypass -File "{0}"' -f $PSCommandPath
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new($executable, $arguments)
    $startInfo.UseShellExecute = $false
    $null = [System.Diagnostics.Process]::Start($startInfo)
    return
}

try {
    Import-Module -Name ([System.IO.Path]::Combine($PSScriptRoot, 'TeamsBlastMessenger.psd1')) -Force
    Show-TbmMainWindow
}
catch {
    $message = 'Teams Blast Messenger could not start: {0}' -f $_.Exception.Message
    try {
        Add-Type -AssemblyName 'PresentationFramework'
        $null = [System.Windows.MessageBox]::Show($message, 'Teams Blast Messenger', 'OK', 'Error')
    }
    catch {
        Write-Error -Message $message
    }
    throw
}
