#region Deep-link fallback
<#
    Used when Graph is not configured. Teams deep links open a 1:1 or group chat with the message pre-filled:
        msteams:/l/chat/0/0?users=a@x.com,b@x.com&topicName=Topic&message=Hello
    Auto mode then focuses the Teams window and presses Enter; it verifies Teams is the foreground
    process before sending the keystroke so Enter never lands in another application.
#>

$script:TbmTeamsProcessNames = @('ms-teams', 'MSTeams', 'Teams')

function New-TbmDeepLink {
    <#
    .SYNOPSIS
        Builds a Teams chat deep link with a pre-filled message.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Emails,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Message,

        [Parameter()]
        [string]$Topic,

        [Parameter()]
        [ValidateSet('Desktop', 'Web')]
        [string]$Target = 'Desktop'
    )

    $base = 'https://teams.microsoft.com/l/chat/0/0'
    if ($Target -eq 'Desktop') { $base = 'msteams:/l/chat/0/0' }

    $encodedUsers = [System.Collections.Generic.List[string]]::new()
    foreach ($email in $Emails) { $encodedUsers.Add([System.Uri]::EscapeDataString($email.Trim())) }

    $query = [System.Collections.Generic.List[string]]::new()
    $query.Add('users=' + [string]::Join(',', $encodedUsers))
    if ($Emails.Length -gt 1 -and -not [string]::IsNullOrWhiteSpace($Topic)) {
        $query.Add('topicName=' + [System.Uri]::EscapeDataString($Topic.Trim()))
    }
    if (-not [string]::IsNullOrEmpty($Message)) {
        $normalized = $Message.Replace("`r`n", "`n").Replace("`r", "`n")
        $query.Add('message=' + [System.Uri]::EscapeDataString($normalized))
    }
    return $base + '?' + [string]::Join('&', $query)
}

function Start-TbmDeepLink {
    <#
    .SYNOPSIS
        Opens a deep link with the registered handler (Teams desktop for msteams:, default browser for https:).
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri
    )

    if (-not $PSCmdlet.ShouldProcess($Uri, 'Open Teams deep link')) { return }
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new($Uri)
    $startInfo.UseShellExecute = $true
    $process = [System.Diagnostics.Process]::Start($startInfo)
    if ($null -ne $process) { $process.Dispose() }
    Write-TbmLog -Level Debug -Source 'DeepLink' -Message ('Opened deep link ({0} chars).' -f $Uri.Length)
}

function Get-TbmTeamsProcessInfo {
    <#
    .SYNOPSIS
        Returns the Teams process ids and the first main window handle found.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    $ids = [System.Collections.Generic.HashSet[int]]::new()
    $handle = [System.IntPtr]::Zero
    foreach ($name in $script:TbmTeamsProcessNames) {
        foreach ($process in [System.Diagnostics.Process]::GetProcessesByName($name)) {
            try {
                [void]$ids.Add($process.Id)
                if ($handle -eq [System.IntPtr]::Zero -and $process.MainWindowHandle -ne [System.IntPtr]::Zero) { $handle = $process.MainWindowHandle }
            }
            finally {
                $process.Dispose()
            }
        }
    }
    return [pscustomobject]@{ ProcessIds = $ids; WindowHandle = $handle }
}

function Send-TbmTeamsEnterKey {
    <#
    .SYNOPSIS
        Brings Teams to the foreground and presses Enter to send the pre-filled message.
        Throws if Teams is not running or could not be made the foreground window.
    #>
    [CmdletBinding()]
    param()

    if (-not (Test-TbmIsWindows)) { throw [System.PlatformNotSupportedException]::new('Auto-send requires Windows.') }
    Add-Type -AssemblyName 'System.Windows.Forms'

    $teams = Get-TbmTeamsProcessInfo
    if ($teams.ProcessIds.Count -eq 0) { throw [System.InvalidOperationException]::new('Microsoft Teams is not running.') }

    $foregroundId = [TeamsBlastMessenger.TbmNative]::GetForegroundProcessId()
    if (-not $teams.ProcessIds.Contains($foregroundId)) {
        [void][TeamsBlastMessenger.TbmNative]::Activate($teams.WindowHandle)
        [System.Threading.Thread]::Sleep(400)
        $foregroundId = [TeamsBlastMessenger.TbmNative]::GetForegroundProcessId()
    }
    if (-not $teams.ProcessIds.Contains($foregroundId)) {
        throw [System.InvalidOperationException]::new('Teams was not the active window, so Enter was not sent. Increase the deep-link delay or use manual mode.')
    }
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
}

#endregion
