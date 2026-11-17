#region HTTP transport

$script:TbmGraphBaseUri = 'https://graph.microsoft.com/v1.0'
$script:TbmHttpClient = $null

function Get-TbmHttpClient {
    <#
    .SYNOPSIS
        Returns the runspace-wide HttpClient (one instance avoids socket exhaustion).
    #>
    [CmdletBinding()]
    [OutputType([System.Net.Http.HttpClient])]
    param()

    if ($null -eq $script:TbmHttpClient) {
        # Windows PowerShell 5.1 may default to TLS 1.0; Graph and Entra require TLS 1.2+.
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
        $client = [System.Net.Http.HttpClient]::new()
        $client.Timeout = [System.TimeSpan]::FromSeconds(100)
        [void]$client.DefaultRequestHeaders.TryAddWithoutValidation('User-Agent', 'TeamsBlastMessenger/1.0')
        $script:TbmHttpClient = $client
    }
    return $script:TbmHttpClient
}

function Invoke-TbmHttpRequest {
    <#
    .SYNOPSIS
        Sends one HTTP request and returns status, body and throttling headers without throwing on HTTP errors.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')]
        [string]$Method,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri,

        [Parameter()]
        [hashtable]$Headers,

        [Parameter()]
        [string]$JsonBody,

        [Parameter()]
        [hashtable]$FormBody
    )

    $client = Get-TbmHttpClient
    $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::new($Method), $Uri)
    $response = $null
    try {
        if ($null -ne $Headers) {
            foreach ($key in $Headers.Keys) {
                [void]$request.Headers.TryAddWithoutValidation([string]$key, [string]$Headers[$key])
            }
        }
        if ($PSBoundParameters.ContainsKey('JsonBody')) {
            $request.Content = [System.Net.Http.StringContent]::new($JsonBody, [System.Text.Encoding]::UTF8, 'application/json')
        }
        elseif ($null -ne $FormBody) {
            $pairs = [System.Collections.Generic.List[System.Collections.Generic.KeyValuePair[string, string]]]::new()
            foreach ($key in $FormBody.Keys) {
                $pairs.Add([System.Collections.Generic.KeyValuePair[string, string]]::new([string]$key, [string]$FormBody[$key]))
            }
            $request.Content = [System.Net.Http.FormUrlEncodedContent]::new($pairs)
        }

        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $response = $client.SendAsync($request).GetAwaiter().GetResult()
        $content = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        $stopwatch.Stop()

        $retryAfterSeconds = 0
        $retryAfter = $response.Headers.RetryAfter
        if ($null -ne $retryAfter) {
            if ($null -ne $retryAfter.Delta) {
                $retryAfterSeconds = [int][System.Math]::Ceiling($retryAfter.Delta.TotalSeconds)
            }
            elseif ($null -ne $retryAfter.Date) {
                $retryAfterSeconds = [int][System.Math]::Ceiling(($retryAfter.Date - [System.DateTimeOffset]::UtcNow).TotalSeconds)
            }
        }

        $requestId = ''
        $values = $null
        if ($response.Headers.TryGetValues('request-id', [ref]$values)) {
            $requestId = [string]::Join(',', [string[]]$values)
        }

        return [pscustomobject]@{
            StatusCode        = [int]$response.StatusCode
            Content           = $content
            RetryAfterSeconds = [System.Math]::Max(0, $retryAfterSeconds)
            RequestId         = $requestId
            DurationMs        = $stopwatch.ElapsedMilliseconds
        }
    }
    finally {
        if ($null -ne $response) { $response.Dispose() }
        $request.Dispose()
    }
}

function Wait-TbmDelay {
    <#
    .SYNOPSIS
        Sleeps in short slices so cancellation is honoured. Returns $false if cancelled.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [ValidateRange(0, 3600000)]
        [int]$Milliseconds,

        [Parameter()]
        [hashtable]$Sync
    )

    $deadline = [datetime]::UtcNow.AddMilliseconds($Milliseconds)
    while ([datetime]::UtcNow -lt $deadline) {
        if ($null -ne $Sync -and [bool]$Sync['Cancel']) { return $false }
        $remaining = [int]($deadline - [datetime]::UtcNow).TotalMilliseconds
        [System.Threading.Thread]::Sleep([System.Math]::Max(1, [System.Math]::Min(200, $remaining)))
    }
    return -not ($null -ne $Sync -and [bool]$Sync['Cancel'])
}

#endregion

#region Graph

function Invoke-TbmGraphRequest {
    <#
    .SYNOPSIS
        Calls Microsoft Graph with the current token, retrying throttled/transient failures and refreshing on 401.
    .OUTPUTS
        The parsed JSON response (PSCustomObject), or $null for empty bodies / allowed 404s.
    #>
    [CmdletBinding()]
    [OutputType([object])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter()]
        [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')]
        [string]$Method = 'GET',

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter()]
        [object]$Body,

        [Parameter()]
        [hashtable]$Headers,

        [Parameter()]
        [ValidateRange(0, 10)]
        [int]$MaxRetries = 6,

        [Parameter()]
        [switch]$AllowNotFound,

        [Parameter()]
        [hashtable]$Sync
    )

    $uri = $Path
    if (-not $Path.StartsWith('https://', [System.StringComparison]::OrdinalIgnoreCase)) { $uri = $script:TbmGraphBaseUri + $Path }
    $logPath = $uri.Replace($script:TbmGraphBaseUri, '')

    $jsonBody = $null
    if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
        $jsonBody = ConvertTo-Json -InputObject $Body -Depth 12 -Compress
    }

    $refreshedAfter401 = $false
    $attempt = 0
    while ($true) {
        $attempt++
        $token = Get-TbmAccessToken -AuthState $AuthState
        $requestHeaders = @{ Authorization = 'Bearer ' + $token }
        if ($null -ne $Headers) { foreach ($key in $Headers.Keys) { $requestHeaders[$key] = $Headers[$key] } }

        $httpParameters = @{ Method = $Method; Uri = $uri; Headers = $requestHeaders }
        if ($null -ne $jsonBody) { $httpParameters['JsonBody'] = $jsonBody }

        $response = $null
        try {
            $response = Invoke-TbmHttpRequest @httpParameters
        }
        catch {
            # Network-level failure (DNS, reset, timeout). Treat as transient.
            if ($attempt -gt $MaxRetries) { throw }
            $delay = [int][System.Math]::Min(60000, 1000 * [System.Math]::Pow(2, $attempt))
            Write-TbmLog -Level Warn -Source 'Graph' -Message ('{0} {1} network error (attempt {2}): {3}. Retrying in {4} ms.' -f $Method, $logPath, $attempt, $_.Exception.GetBaseException().Message, $delay)
            if (-not (Wait-TbmDelay -Milliseconds $delay -Sync $Sync)) { throw [System.OperationCanceledException]::new('Cancelled.') }
            continue
        }

        Write-TbmLog -Level Info -Source 'Graph' -Message ('{0} {1} -> {2} ({3} ms)' -f $Method, $logPath, $response.StatusCode, $response.DurationMs)

        if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 300) {
            if ([string]::IsNullOrWhiteSpace($response.Content)) { return $null }
            return ConvertFrom-Json -InputObject $response.Content
        }

        if ($response.StatusCode -eq 401 -and -not $refreshedAfter401) {
            $refreshedAfter401 = $true
            Write-TbmLog -Level Warn -Source 'Graph' -Message 'Received 401; refreshing access token and retrying once.'
            [void](Get-TbmAccessToken -AuthState $AuthState -ForceRefresh)
            continue
        }

        if ($response.StatusCode -eq 404 -and $AllowNotFound) { return $null }

        $transient = $response.StatusCode -eq 429 -or $response.StatusCode -eq 502 -or $response.StatusCode -eq 503 -or $response.StatusCode -eq 504
        if ($transient -and $attempt -le $MaxRetries) {
            $delay = $response.RetryAfterSeconds * 1000
            if ($delay -le 0) { $delay = [int][System.Math]::Min(60000, 1000 * [System.Math]::Pow(2, $attempt)) }
            Write-TbmLog -Level Warn -Source 'Graph' -Message ('{0} {1} throttled/unavailable ({2}); waiting {3} ms (attempt {4}/{5}).' -f $Method, $logPath, $response.StatusCode, $delay, $attempt, $MaxRetries)
            if (-not (Wait-TbmDelay -Milliseconds $delay -Sync $Sync)) { throw [System.OperationCanceledException]::new('Cancelled.') }
            continue
        }

        $errorCode = 'Unknown'
        $errorMessage = $response.Content
        try {
            $errorObject = ConvertFrom-Json -InputObject $response.Content
            $code = Get-TbmPropertyValue -InputObject $errorObject -Name 'error.code'
            $message = Get-TbmPropertyValue -InputObject $errorObject -Name 'error.message'
            if ($null -ne $code) { $errorCode = [string]$code }
            if ($null -ne $message) { $errorMessage = [string]$message }
        }
        catch {
            Write-TbmLog -Level Debug -Source 'Graph' -Message 'Error body was not JSON.'
        }
        throw [TeamsBlastMessenger.TbmGraphException]::new($response.StatusCode, $errorCode, $errorMessage, $response.RequestId, $Method, $logPath)
    }
}

function Invoke-TbmGraphPagedRequest {
    <#
    .SYNOPSIS
        GETs a Graph collection and follows @odata.nextLink, returning all items.
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[object]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter()]
        [ValidateRange(1, 10000)]
        [int]$MaxPages = 100,

        [Parameter()]
        [hashtable]$Headers,

        [Parameter()]
        [hashtable]$Sync,

        # Optional predicate; paging stops after a page where it returns $true for any item.
        [Parameter()]
        [scriptblock]$StopWhen
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Path
    $page = 0
    while (-not [string]::IsNullOrEmpty($next) -and $page -lt $MaxPages) {
        if ($null -ne $Sync -and [bool]$Sync['Cancel']) { throw [System.OperationCanceledException]::new('Cancelled.') }
        $page++
        $response = Invoke-TbmGraphRequest -AuthState $AuthState -Method GET -Path $next -Headers $Headers -Sync $Sync
        $stop = $false
        $values = Get-TbmPropertyValue -InputObject $response -Name 'value'
        if ($null -ne $values) {
            foreach ($value in $values) {
                $items.Add($value)
                if ($null -ne $StopWhen -and (& $StopWhen $value)) { $stop = $true }
            }
        }
        if ($stop) { break }
        $next = [string](Get-TbmPropertyValue -InputObject $response -Name '@odata.nextLink')
    }
    if ($page -ge $MaxPages -and -not [string]::IsNullOrEmpty($next)) {
        Write-TbmLog -Level Warn -Source 'Graph' -Message ('Stopped paging {0} after {1} page(s); results may be incomplete.' -f $Path, $MaxPages)
    }
    return , $items
}

#endregion

#region Graph operations

function Resolve-TbmGraphUser {
    <#
    .SYNOPSIS
        Resolves an email/UPN to a directory user (id, displayName). Tries UPN lookup, then mail filter.
    .OUTPUTS
        PSCustomObject with Id, DisplayName, Upn, Mail; $null if not found.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Email,

        [Parameter()]
        [System.Collections.Generic.Dictionary[string, object]]$Cache,

        [Parameter()]
        [hashtable]$Sync
    )

    $key = $Email.Trim().ToLowerInvariant()
    if ($null -ne $Cache -and $Cache.ContainsKey($key)) { return $Cache[$key] }

    $select = '$select=id,displayName,userPrincipalName,mail'
    $user = Invoke-TbmGraphRequest -AuthState $AuthState -Path ('/users/{0}?{1}' -f [System.Uri]::EscapeDataString($key), $select) -AllowNotFound -Sync $Sync
    if ($null -eq $user) {
        $filter = [System.Uri]::EscapeDataString(("mail eq '{0}'" -f $key.Replace("'", "''")))
        $result = Invoke-TbmGraphRequest -AuthState $AuthState -Path ('/users?$filter={0}&{1}' -f $filter, $select) -Sync $Sync
        $found = Get-TbmPropertyValue -InputObject $result -Name 'value'
        if ($null -ne $found -and @($found).Count -gt 0) { $user = @($found)[0] }
    }

    $resolved = $null
    if ($null -ne $user) {
        $resolved = [pscustomobject]@{
            Id          = [string]$user.id
            DisplayName = [string](Get-TbmPropertyValue -InputObject $user -Name 'displayName')
            Upn         = [string](Get-TbmPropertyValue -InputObject $user -Name 'userPrincipalName')
            Mail        = [string](Get-TbmPropertyValue -InputObject $user -Name 'mail')
        }
    }
    if ($null -ne $Cache) { $Cache[$key] = $resolved }
    return $resolved
}

function New-TbmGraphChat {
    <#
    .SYNOPSIS
        Creates (or, for oneOnOne, returns the existing) chat with the given user ids. Returns the chat id.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$MemberUserIds,

        [Parameter()]
        [string]$Topic,

        [Parameter()]
        [hashtable]$Sync
    )

    $allIds = [System.Collections.Generic.List[string]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    [void]$seen.Add([string]$AuthState.MeId)
    $allIds.Add([string]$AuthState.MeId)
    foreach ($id in $MemberUserIds) { if ($seen.Add($id)) { $allIds.Add($id) } }

    if ($allIds.Count -lt 2) {
        # Messaging yourself: Teams exposes the personal "notes" chat with this well-known id.
        return '48:notes'
    }

    $chatType = 'group'
    if ($allIds.Count -eq 2) { $chatType = 'oneOnOne' }

    $members = [System.Collections.Generic.List[object]]::new()
    foreach ($id in $allIds) {
        $members.Add(@{
                '@odata.type'     = '#microsoft.graph.aadUserConversationMember'
                'roles'           = @('owner')
                'user@odata.bind' = ("https://graph.microsoft.com/v1.0/users('{0}')" -f $id)
            })
    }
    $body = @{ chatType = $chatType; members = $members.ToArray() }
    if ($chatType -eq 'group' -and -not [string]::IsNullOrWhiteSpace($Topic)) { $body['topic'] = $Topic.Trim() }

    $chat = Invoke-TbmGraphRequest -AuthState $AuthState -Method POST -Path '/chats' -Body $body -Sync $Sync
    return [string]$chat.id
}

function Send-TbmGraphChatMessage {
    <#
    .SYNOPSIS
        Posts an HTML message to a chat. Returns the created message (id, createdDateTime).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$AuthState,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ChatId,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$HtmlContent,

        [Parameter()]
        [hashtable]$Sync
    )

    $body = @{ body = @{ contentType = 'html'; content = $HtmlContent } }
    $path = '/chats/{0}/messages' -f [System.Uri]::EscapeDataString($ChatId)
    $message = Invoke-TbmGraphRequest -AuthState $AuthState -Method POST -Path $path -Body $body -Sync $Sync -MaxRetries 4
    return [pscustomobject]@{
        Id         = [string]$message.id
        CreatedUtc = ConvertTo-TbmUtcDateTime -Value (Get-TbmPropertyValue -InputObject $message -Name 'createdDateTime')
    }
}

#endregion
