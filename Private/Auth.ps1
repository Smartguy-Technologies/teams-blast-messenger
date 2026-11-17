#region Authentication
<#
    Two Graph sign-in paths share one "auth state" (a synchronized hashtable passed to worker runspaces):
      GraphSdk  - Microsoft.Graph.Authentication module (Connect-MgGraph). The access token is extracted
                  from an Invoke-MgGraphRequest response so all calls use the same HTTP client.
      GraphRest - Your own Entra app registration (public client) using the OAuth 2.0 device code flow.
                  The refresh token is cached with DPAPI (CurrentUser) under %APPDATA%.
    Delegated permissions are required: Teams chat messages cannot be sent with application permissions.
#>

$script:TbmGraphScopes = @('User.Read', 'User.ReadBasic.All', 'Chat.Create', 'ChatMessage.Send', 'Chat.Read')
$script:TbmDpapiEntropy = [System.Text.Encoding]::UTF8.GetBytes('TeamsBlastMessenger.v1')

function New-TbmAuthState {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()

    return [hashtable]::Synchronized(@{
            Mode         = 'None'     # None | GraphSdk | GraphRest
            AccessToken  = $null
            ExpiresUtc   = [datetime]::MinValue
            RefreshToken = $null
            ClientId     = $null
            TenantId     = $null
            MeId         = $null
            MeUpn        = $null
            MeName       = $null
        })
}

function Test-TbmGraphConnected {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState
    )
    return $AuthState.Mode -ne 'None' -and -not [string]::IsNullOrEmpty([string]$AuthState.MeId)
}

function Test-TbmGraphSdkAvailable {
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    $modules = Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication'
    return $null -ne $modules -and @($modules).Count -gt 0
}

function Resolve-TbmAuthMode {
    <#
    .SYNOPSIS
        Picks GraphRest or GraphSdk from settings and what is installed. Returns $null if neither is usable.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Settings
    )

    $hasClientId = -not [string]::IsNullOrWhiteSpace([string]$Settings.ClientId)
    switch ([string]$Settings.AuthMode) {
        'GraphRest' { if ($hasClientId) { return 'GraphRest' } return $null }
        'GraphSdk' { if (Test-TbmGraphSdkAvailable) { return 'GraphSdk' } return $null }
        default {
            if ($hasClientId) { return 'GraphRest' }
            if (Test-TbmGraphSdkAvailable) { return 'GraphSdk' }
            return $null
        }
    }
}

function Get-TbmJwtExpiryUtc {
    <#
    .SYNOPSIS
        Reads the "exp" claim from a JWT access token without validating it.
    #>
    [CmdletBinding()]
    [OutputType([datetime])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Token
    )

    $fallback = [datetime]::UtcNow.AddMinutes(45)
    $parts = $Token.Split('.')
    if ($parts.Length -lt 2) { return $fallback }
    $payload = $parts[1].Replace('-', '+').Replace('_', '/')
    switch ($payload.Length % 4) {
        2 { $payload += '==' }
        3 { $payload += '=' }
    }
    try {
        $json = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($payload))
        $match = [System.Text.RegularExpressions.Regex]::Match($json, '"exp"\s*:\s*(\d+)')
        if ($match.Success) {
            return [System.DateTimeOffset]::FromUnixTimeSeconds([long]$match.Groups[1].Value).UtcDateTime
        }
    }
    catch {
        Write-TbmLog -Level Debug -Source 'Auth' -Message 'Access token is not a decodable JWT; assuming 45 minute lifetime.'
    }
    return $fallback
}

function Get-TbmAccessToken {
    <#
    .SYNOPSIS
        Returns a valid access token, refreshing it when it expires within five minutes (or when forced).
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter()]
        [switch]$ForceRefresh
    )

    if ($AuthState.Mode -eq 'None') { throw [System.InvalidOperationException]::new('Not connected to Microsoft Graph.') }

    [System.Threading.Monitor]::Enter($AuthState.SyncRoot)
    try {
        $needsRefresh = $ForceRefresh -or [string]::IsNullOrEmpty([string]$AuthState.AccessToken) -or ([datetime]$AuthState.ExpiresUtc -le [datetime]::UtcNow.AddMinutes(5))
        if ($needsRefresh) {
            if ($AuthState.Mode -eq 'GraphRest') { Update-TbmRestToken -AuthState $AuthState }
            else { Update-TbmSdkToken -AuthState $AuthState }
        }
        return [string]$AuthState.AccessToken
    }
    finally {
        [System.Threading.Monitor]::Exit($AuthState.SyncRoot)
    }
}

function Initialize-TbmSignedInUser {
    <#
    .SYNOPSIS
        Populates MeId / MeUpn / MeName from GET /me.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState
    )

    $me = Invoke-TbmGraphRequest -AuthState $AuthState -Path '/me?$select=id,displayName,userPrincipalName,mail'
    $AuthState.MeId = [string]$me.id
    $AuthState.MeUpn = [string](Get-TbmPropertyValue -InputObject $me -Name 'userPrincipalName')
    $AuthState.MeName = [string](Get-TbmPropertyValue -InputObject $me -Name 'displayName')
    Write-TbmLog -Level Info -Source 'Auth' -Message ('Signed in as {0} ({1}) via {2}.' -f $AuthState.MeName, $AuthState.MeUpn, $AuthState.Mode)
}

#region Graph PowerShell SDK

function Connect-TbmGraphSdk {
    <#
    .SYNOPSIS
        Interactive sign-in with Connect-MgGraph, then captures the token and signed-in user.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Settings
    )

    Import-Module -Name 'Microsoft.Graph.Authentication' -ErrorAction Stop
    $command = Get-Command -Name 'Connect-MgGraph' -ErrorAction Stop
    $connectParameters = @{ Scopes = $script:TbmGraphScopes; ErrorAction = 'Stop' }
    if ($command.Parameters.ContainsKey('NoWelcome')) { $connectParameters['NoWelcome'] = $true }
    if (-not [string]::IsNullOrWhiteSpace([string]$Settings.ClientId)) {
        $connectParameters['ClientId'] = [string]$Settings.ClientId
        $connectParameters['TenantId'] = [string]$Settings.TenantId
    }
    Write-TbmLog -Level Info -Source 'Auth' -Message 'Connecting with Microsoft Graph PowerShell SDK.'
    $null = Connect-MgGraph @connectParameters

    $AuthState.Mode = 'GraphSdk'
    try {
        Update-TbmSdkToken -AuthState $AuthState
        Initialize-TbmSignedInUser -AuthState $AuthState
    }
    catch {
        $AuthState.Mode = 'None'
        throw
    }
}

function Update-TbmSdkToken {
    <#
    .SYNOPSIS
        Obtains the SDK's current access token by issuing a lightweight request and reading its Authorization header.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState
    )

    if ($null -eq (Get-Command -Name 'Invoke-MgGraphRequest' -ErrorAction SilentlyContinue)) {
        Import-Module -Name 'Microsoft.Graph.Authentication' -ErrorAction Stop
    }
    $response = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me?$select=id' -OutputType HttpResponseMessage -ErrorAction Stop
    try {
        $authorization = $response.RequestMessage.Headers.Authorization
        if ($null -eq $authorization -or [string]::IsNullOrEmpty($authorization.Parameter)) {
            throw [System.InvalidOperationException]::new('The Graph SDK did not expose an access token. Reconnect, or use an app registration (Client ID) instead.')
        }
        $AuthState.AccessToken = $authorization.Parameter
        $AuthState.ExpiresUtc = Get-TbmJwtExpiryUtc -Token $authorization.Parameter
        Write-TbmLog -Level Debug -Source 'Auth' -Message ('SDK token acquired; expires {0:u}.' -f $AuthState.ExpiresUtc)
    }
    finally {
        $response.Dispose()
    }
}

#endregion

#region REST (device code flow)

function Get-TbmAuthorityUri {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [ValidateSet('devicecode', 'token')]
        [string]$Endpoint
    )
    return 'https://login.microsoftonline.com/{0}/oauth2/v2.0/{1}' -f [System.Uri]::EscapeDataString($TenantId.Trim()), $Endpoint
}

function Get-TbmRestScopeString {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    $scopes = [System.Collections.Generic.List[string]]::new()
    $scopes.Add('offline_access')
    foreach ($scope in $script:TbmGraphScopes) { $scopes.Add('https://graph.microsoft.com/' + $scope) }
    return [string]::Join(' ', $scopes)
}

function Start-TbmDeviceCodeLogin {
    <#
    .SYNOPSIS
        Starts the device code flow and returns the user code / verification URL to show to the user.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Settings
    )

    $clientId = ([string]$Settings.ClientId).Trim()
    $tenantId = ([string]$Settings.TenantId).Trim()
    if ([string]::IsNullOrEmpty($tenantId)) { $tenantId = 'organizations' }

    $response = Invoke-TbmHttpRequest -Method POST -Uri (Get-TbmAuthorityUri -TenantId $tenantId -Endpoint 'devicecode') -FormBody @{
        client_id = $clientId
        scope     = Get-TbmRestScopeString
    }
    $data = ConvertFrom-Json -InputObject $response.Content
    if ($response.StatusCode -ne 200) {
        $description = [string](Get-TbmPropertyValue -InputObject $data -Name 'error_description')
        throw [System.InvalidOperationException]::new('Device code request failed: ' + $description)
    }
    Write-TbmLog -Level Info -Source 'Auth' -Message 'Device code sign-in started.'
    return [pscustomobject]@{
        ClientId        = $clientId
        TenantId        = $tenantId
        DeviceCode      = [string]$data.device_code
        UserCode        = [string]$data.user_code
        VerificationUri = [string]$data.verification_uri
        Message         = [string]$data.message
        ExpiresUtc      = [datetime]::UtcNow.AddSeconds([int]$data.expires_in)
        IntervalSeconds = [System.Math]::Max(1, [int]$data.interval)
    }
}

function Request-TbmDeviceCodeToken {
    <#
    .SYNOPSIS
        Polls the token endpoint once. Returns 'Pending', 'SlowDown' or 'Success'; throws when the flow fails.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [pscustomobject]$DeviceCode
    )

    if ([datetime]::UtcNow -gt $DeviceCode.ExpiresUtc) { throw [System.TimeoutException]::new('The sign-in code expired. Start again.') }

    $response = Invoke-TbmHttpRequest -Method POST -Uri (Get-TbmAuthorityUri -TenantId $DeviceCode.TenantId -Endpoint 'token') -FormBody @{
        grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
        client_id   = $DeviceCode.ClientId
        device_code = $DeviceCode.DeviceCode
    }
    $data = ConvertFrom-Json -InputObject $response.Content
    if ($response.StatusCode -eq 200) {
        $AuthState.ClientId = $DeviceCode.ClientId
        $AuthState.TenantId = $DeviceCode.TenantId
        Set-TbmRestTokenResponse -AuthState $AuthState -TokenResponse $data
        $AuthState.Mode = 'GraphRest'
        try {
            Initialize-TbmSignedInUser -AuthState $AuthState
        }
        catch {
            $AuthState.Mode = 'None'
            throw
        }
        Save-TbmTokenCache -AuthState $AuthState
        return 'Success'
    }

    $errorCode = [string](Get-TbmPropertyValue -InputObject $data -Name 'error')
    switch ($errorCode) {
        'authorization_pending' { return 'Pending' }
        'slow_down' { return 'SlowDown' }
        default {
            $description = [string](Get-TbmPropertyValue -InputObject $data -Name 'error_description')
            Write-TbmLog -Level Warn -Source 'Auth' -Message ('Device code sign-in failed: {0}' -f $errorCode)
            throw [System.InvalidOperationException]::new(('Sign-in failed ({0}): {1}' -f $errorCode, $description))
        }
    }
}

function Set-TbmRestTokenResponse {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [object]$TokenResponse
    )

    $AuthState.AccessToken = [string]$TokenResponse.access_token
    $AuthState.ExpiresUtc = [datetime]::UtcNow.AddSeconds([int]$TokenResponse.expires_in)
    $newRefresh = [string](Get-TbmPropertyValue -InputObject $TokenResponse -Name 'refresh_token')
    if (-not [string]::IsNullOrEmpty($newRefresh)) { $AuthState.RefreshToken = $newRefresh }
}

function Update-TbmRestToken {
    <#
    .SYNOPSIS
        Redeems the refresh token for a new access token (and rotated refresh token).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState
    )

    if ([string]::IsNullOrEmpty([string]$AuthState.RefreshToken)) {
        throw [System.InvalidOperationException]::new('No refresh token available. Please connect again.')
    }
    $response = Invoke-TbmHttpRequest -Method POST -Uri (Get-TbmAuthorityUri -TenantId $AuthState.TenantId -Endpoint 'token') -FormBody @{
        grant_type    = 'refresh_token'
        client_id     = $AuthState.ClientId
        refresh_token = $AuthState.RefreshToken
        scope         = Get-TbmRestScopeString
    }
    $data = ConvertFrom-Json -InputObject $response.Content
    if ($response.StatusCode -ne 200) {
        $errorCode = [string](Get-TbmPropertyValue -InputObject $data -Name 'error')
        Write-TbmLog -Level Warn -Source 'Auth' -Message ('Token refresh failed: {0}' -f $errorCode)
        throw [System.InvalidOperationException]::new(('Your Graph session expired ({0}). Please connect again.' -f $errorCode))
    }
    Set-TbmRestTokenResponse -AuthState $AuthState -TokenResponse $data
    Save-TbmTokenCache -AuthState $AuthState
    Write-TbmLog -Level Debug -Source 'Auth' -Message ('REST token refreshed; expires {0:u}.' -f $AuthState.ExpiresUtc)
}

function Test-TbmDpapiAvailable {
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    return (Test-TbmIsWindows) -and $null -ne ('System.Security.Cryptography.ProtectedData' -as [type])
}

function Save-TbmTokenCache {
    <#
    .SYNOPSIS
        Persists the refresh token encrypted with DPAPI (current user). Skipped where DPAPI is unavailable.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState
    )

    if (-not (Test-TbmDpapiAvailable)) {
        Write-TbmLog -Level Warn -Source 'Auth' -Message 'DPAPI unavailable; refresh token not persisted.'
        return
    }
    $payload = ConvertTo-Json -Compress -InputObject @{
        ClientId     = $AuthState.ClientId
        TenantId     = $AuthState.TenantId
        RefreshToken = $AuthState.RefreshToken
    }
    $plain = [System.Text.Encoding]::UTF8.GetBytes($payload)
    $protected = [System.Security.Cryptography.ProtectedData]::Protect($plain, $script:TbmDpapiEntropy, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
    $path = Get-TbmAppDataPath -Name 'TokenCache'
    [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($path))
    [System.IO.File]::WriteAllBytes($path, $protected)
}

function Clear-TbmTokenCache {
    [CmdletBinding()]
    param()
    $path = Get-TbmAppDataPath -Name 'TokenCache'
    if ([System.IO.File]::Exists($path)) {
        [System.IO.File]::Delete($path)
        Write-TbmLog -Level Info -Source 'Auth' -Message 'Token cache cleared.'
    }
}

function Connect-TbmGraphRestSilent {
    <#
    .SYNOPSIS
        Restores a REST session from the DPAPI token cache. Returns $true on success.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$Settings
    )

    $path = Get-TbmAppDataPath -Name 'TokenCache'
    if (-not (Test-TbmDpapiAvailable) -or -not [System.IO.File]::Exists($path)) { return $false }

    try {
        $protected = [System.IO.File]::ReadAllBytes($path)
        $plain = [System.Security.Cryptography.ProtectedData]::Unprotect($protected, $script:TbmDpapiEntropy, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        $cache = ConvertFrom-Json -InputObject ([System.Text.Encoding]::UTF8.GetString($plain))
        $sameApp = [string]::Equals([string]$cache.ClientId, ([string]$Settings.ClientId).Trim(), [System.StringComparison]::OrdinalIgnoreCase) -and
                   [string]::Equals([string]$cache.TenantId, ([string]$Settings.TenantId).Trim(), [System.StringComparison]::OrdinalIgnoreCase)
        if (-not $sameApp) {
            Write-TbmLog -Level Info -Source 'Auth' -Message 'Cached token belongs to a different app registration; ignoring it.'
            return $false
        }
        $AuthState.ClientId = [string]$cache.ClientId
        $AuthState.TenantId = [string]$cache.TenantId
        $AuthState.RefreshToken = [string]$cache.RefreshToken
        $AuthState.Mode = 'GraphRest'
        Update-TbmRestToken -AuthState $AuthState
        Initialize-TbmSignedInUser -AuthState $AuthState
        return $true
    }
    catch {
        Write-TbmLog -Level Warn -Source 'Auth' -Message ('Silent sign-in failed: {0}' -f $_.Exception.Message)
        Reset-TbmAuthState -AuthState $AuthState
        return $false
    }
}

#endregion

function Reset-TbmAuthState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState
    )
    $AuthState.Mode = 'None'
    $AuthState.AccessToken = $null
    $AuthState.ExpiresUtc = [datetime]::MinValue
    $AuthState.RefreshToken = $null
    $AuthState.MeId = $null
    $AuthState.MeUpn = $null
    $AuthState.MeName = $null
}

function Disconnect-TbmGraph {
    <#
    .SYNOPSIS
        Signs out: clears in-memory tokens and the token cache, and disconnects the SDK if it was used.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState
    )

    if ($AuthState.Mode -eq 'GraphSdk' -and $null -ne (Get-Command -Name 'Disconnect-MgGraph' -ErrorAction SilentlyContinue)) {
        try { $null = Disconnect-MgGraph -ErrorAction Stop }
        catch { Write-TbmLog -Level Warn -Source 'Auth' -Message ('Disconnect-MgGraph: {0}' -f $_.Exception.Message) }
    }
    Clear-TbmTokenCache
    Reset-TbmAuthState -AuthState $AuthState
    Write-TbmLog -Level Info -Source 'Auth' -Message 'Disconnected from Microsoft Graph.'
}

#endregion
