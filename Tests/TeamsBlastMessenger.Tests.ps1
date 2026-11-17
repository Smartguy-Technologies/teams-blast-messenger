#Requires -Modules Pester
<#
    Unit tests for the non-UI logic. Run from the project folder:
        Invoke-Pester -Path .\Tests
    App data is redirected to a temp folder via TBM_APPDATA.
#>

BeforeAll {
    $script:TestRoot = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), 'tbm-tests-' + [guid]::NewGuid().ToString('N'))
    [void][System.IO.Directory]::CreateDirectory($script:TestRoot)
    [System.Environment]::SetEnvironmentVariable('TBM_APPDATA', [System.IO.Path]::Combine($script:TestRoot, 'appdata'))
    $script:ProjectRoot = [System.IO.Path]::GetDirectoryName($PSScriptRoot)
    Import-Module ([System.IO.Path]::Combine($script:ProjectRoot, 'TeamsBlastMessenger.psd1')) -Force
    InModuleScope TeamsBlastMessenger { [void](Initialize-TbmEnvironment) }
}

AfterAll {
    Remove-Module TeamsBlastMessenger -ErrorAction SilentlyContinue
    [System.Environment]::SetEnvironmentVariable('TBM_APPDATA', $null)
    if ([System.IO.Directory]::Exists($script:TestRoot)) { [System.IO.Directory]::Delete($script:TestRoot, $true) }
}

Describe 'Token template engine' {
    It 'matches tokens ignoring case, spaces and punctuation, with defaults and escapes' {
        InModuleScope TeamsBlastMessenger {
            $context = New-TbmTokenContext -ColumnNames @('Email', 'First Name', 'Due Date') -BuiltIns @{ SenderName = 'Pat Lee' }
            $row = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $row['Email'] = 'a@contoso.com'; $row['First Name'] = 'Ana'; $row['Due Date'] = ''
            $result = Expand-TbmTemplate -Template 'Hi {first name}/{First_Name}/{FIRSTNAME}! Due {Due Date|soon}. {{x}} from {SenderName} {Nope}' -Context $context -Row $row
            $result.Text | Should -Be 'Hi Ana/Ana/Ana! Due soon. {x} from Pat Lee {Nope}'
            $result.UnknownTokens | Should -Be @('Nope')
            $result.MissingTokens.Count | Should -Be 0
        }
    }

    It 'reports blank values without defaults and HTML-encodes values in HTML mode' {
        InModuleScope TeamsBlastMessenger {
            $context = New-TbmTokenContext -ColumnNames @('Name', 'Team') -BuiltIns @{}
            $row = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $row['Name'] = 'Tom <& Jerry>'; $row['Team'] = '  '
            $result = Expand-TbmTemplate -Template '<b>{Name}</b> {Team}' -Context $context -Row $row -HtmlEncodeValues
            $result.Text | Should -Be '<b>Tom &lt;&amp; Jerry&gt;</b> '
            $result.MissingTokens | Should -Be @('Team')
        }
    }

    It 'flags column tokens, unknown tokens and header collisions' {
        InModuleScope TeamsBlastMessenger {
            $context = New-TbmTokenContext -ColumnNames @('First Name', 'First_Name') -BuiltIns (Get-TbmBuiltInTokenValues)
            $context.Collisions.Count | Should -Be 1
            $check = Test-TbmTemplate -Template '{first name} {Today} {Missing}' -Context $context
            $check.UsesRowData | Should -BeTrue
            $check.UnknownTokens | Should -Be @('Missing')
        }
    }

    It 'converts text to Teams HTML and back' {
        InModuleScope TeamsBlastMessenger {
            ConvertTo-TbmGraphHtml -Message "a < b`r`nline 2" -Format Text | Should -Be 'a &lt; b<br>line 2'
            ConvertTo-TbmPlainText -Message '<p>Hello&nbsp;<b>Bob</b></p><p>x &amp; y<br/>z</p>' -Format Html | Should -Be "Hello Bob`nx & y`nz"
        }
    }
}

Describe 'Recipient files' {
    It 'reads quoted, multi-line and semicolon-delimited CSV with a UTF-8 BOM' {
        InModuleScope TeamsBlastMessenger -Parameters @{ Root = $script:TestRoot } {
            param($Root)
            $path = [System.IO.Path]::Combine($Root, 'semi.csv')
            $text = "Email;First Name;Note`r`nzoë@contoso.com;Zoë;""multi`r`nline; with semicolon""`r`n;;`r`nbad;Bad;x`r`n"
            [System.IO.File]::WriteAllText($path, $text, [System.Text.UTF8Encoding]::new($true))
            $result = Import-TbmRecipientFile -Path $path
            $table = $result.Table
            $table.Rows.Count | Should -Be 2
            $table.Rows[0]['First Name'] | Should -Be 'Zoë'
            $table.Rows[0]['Note'] | Should -Be "multi`r`nline; with semicolon"
            Find-TbmRecipientColumn -Table $table | Should -Be 'Email'
            $summary = Update-TbmRecipientValidation -Table $table -RecipientColumn 'Email'
            $summary.Invalid | Should -Be 1
            [bool]$table.Rows[1][$script:TbmIncludeColumn] | Should -BeFalse
        }
    }

    It 'flags duplicates and names blank or repeated headers' {
        InModuleScope TeamsBlastMessenger -Parameters @{ Root = $script:TestRoot } {
            param($Root)
            $path = [System.IO.Path]::Combine($Root, 'dupes.csv')
            [System.IO.File]::WriteAllText($path, "Mail,Name,Name,`na@x.com,A,A2,z`nA@X.com,B,B2,`n")
            $table = (Import-TbmRecipientFile -Path $path).Table
            (Get-TbmDataColumnNames -Table $table) | Should -Be @('Mail', 'Name', 'Name (2)', 'Column 4')
            $table.Rows[0]['Column 4'] | Should -Be 'z'
            $summary = Update-TbmRecipientValidation -Table $table -RecipientColumn 'Mail'
            $summary.Duplicates | Should -Be 1
            $table.Rows[1][$script:TbmStatusColumn] | Should -Be 'Duplicate'
        }
    }

    It 'reads XLSX natively: shared/inline strings, dates, percent, booleans, sparse cells, sheet choice' {
        InModuleScope TeamsBlastMessenger -Parameters @{ Root = $script:TestRoot } {
            param($Root)
            $path = [System.IO.Path]::Combine($Root, 'book.xlsx')
            $parts = [ordered]@{
                '[Content_Types].xml'        = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
                'xl/workbook.xml'            = '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Notes" sheetId="1" r:id="rId1"/><sheet name="People" sheetId="2" r:id="rId2"/></sheets></workbook>'
                'xl/_rels/workbook.xml.rels' = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="ws" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="ws" Target="/xl/worksheets/sheet2.xml"/></Relationships>'
                'xl/sharedStrings.xml'       = '<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Email</t></si><si><t>First Name</t></si><si><r><t>Ad</t></r><r><t>ele</t></r></si><si><t>Due</t></si><si><t>Score</t></si><si><t>Active</t></si></sst>'
                'xl/styles.xml'              = '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/></numFmts><cellXfs><xf numFmtId="0"/><xf numFmtId="164"/><xf numFmtId="9"/></cellXfs></styleSheet>'
                'xl/worksheets/sheet1.xml'   = '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>ignore</t></is></c></row></sheetData></worksheet>'
                'xl/worksheets/sheet2.xml'   = '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
                '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>3</v></c><c r="D1" t="s"><v>4</v></c><c r="E1" t="s"><v>5</v></c></row>' +
                '<row r="2"><c r="A2" t="inlineStr"><is><t>adele@contoso.com</t></is></c><c r="B2" t="s"><v>2</v></c><c r="C2" s="1"><v>46357</v></c><c r="D2" s="2"><v>0.875</v></c><c r="E2" t="b"><v>1</v></c></row>' +
                '<row r="4"><c r="A4" t="str"><v>bob@contoso.com</v></c><c r="D4"><v>42</v></c></row>' +
                '</sheetData></worksheet>'
            }
            $stream = [System.IO.File]::Create($path)
            $archive = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Create)
            foreach ($name in $parts.Keys) {
                $writer = [System.IO.StreamWriter]::new($archive.CreateEntry($name).Open())
                $writer.Write($parts[$name]); $writer.Dispose()
            }
            $archive.Dispose(); $stream.Dispose()

            $result = Import-TbmRecipientFile -Path $path -SheetName 'People'
            $result.Sheets | Should -Be @('Notes', 'People')
            $result.Reader | Should -Be 'OpenXML'
            $table = $result.Table
            $table.Rows.Count | Should -Be 2
            $table.Rows[0]['Email'] | Should -Be 'adele@contoso.com'
            $table.Rows[0]['First Name'] | Should -Be 'Adele'
            $table.Rows[0]['Due'] | Should -Be ([datetime]::FromOADate(46357).ToString('d', [System.Globalization.CultureInfo]::CurrentCulture))
            $table.Rows[0]['Score'] | Should -Be ((87.5).ToString('0.##', [System.Globalization.CultureInfo]::CurrentCulture) + '%')
            $table.Rows[0]['Active'] | Should -Be 'TRUE'
            $table.Rows[1]['Email'] | Should -Be 'bob@contoso.com'
            $table.Rows[1]['Score'] | Should -Be '42'
            $table.Rows[1]['First Name'] | Should -Be ''

            (Import-TbmRecipientFile -Path $path).SheetName | Should -Be 'Notes'
        }
    }

    It 'classifies number formats' {
        InModuleScope TeamsBlastMessenger {
            Get-TbmNumberFormatKind -FormatId 14 -FormatCode $null | Should -Be 'Date'
            Get-TbmNumberFormatKind -FormatId 22 -FormatCode $null | Should -Be 'DateTime'
            Get-TbmNumberFormatKind -FormatId 165 -FormatCode 'h:mm AM/PM' | Should -Be 'Time'
            Get-TbmNumberFormatKind -FormatId 166 -FormatCode '[$-409]mmmm d, yyyy;@' | Should -Be 'Date'
            Get-TbmNumberFormatKind -FormatId 167 -FormatCode '"Day "0' | Should -Be 'Number'
            Get-TbmNumberFormatKind -FormatId 168 -FormatCode '0.0%' | Should -Be 'Percent'
            Get-TbmNumberFormatKind -FormatId -1 -FormatCode 'General' | Should -Be 'Number'
        }
    }
}

Describe 'Deep links' {
    It 'encodes users, topic and message' {
        InModuleScope TeamsBlastMessenger {
            New-TbmDeepLink -Emails @('a@x.com', 'b+1@y.com') -Message "Hi & bye`r`nline" -Topic 'Q4 #1' |
                Should -Be 'msteams:/l/chat/0/0?users=a%40x.com,b%2B1%40y.com&topicName=Q4%20%231&message=Hi%20%26%20bye%0Aline'
            New-TbmDeepLink -Emails @('a@x.com') -Message 'x' -Topic 'ignored for 1:1' -Target Web |
                Should -Be 'https://teams.microsoft.com/l/chat/0/0?users=a%40x.com&message=x'
        }
    }
}

Describe 'Send history' {
    It 'round-trips a blast record and recovers from a journal after an interruption' {
        InModuleScope TeamsBlastMessenger {
            $blast = [TeamsBlastMessenger.TbmBlastRecord]::new()
            $blast.Mode = 'Individual'; $blast.Transport = 'GraphRest'; $blast.Template = 'Hi {Name}'; $blast.Format = 'Text'
            foreach ($email in @('a@x.com', 'b@x.com')) {
                $recipient = [TeamsBlastMessenger.TbmRecipientResult]::new()
                $recipient.Email = $email; $recipient.RenderedMessage = 'Hi ' + $email
                $blast.Recipients.Add($recipient)
            }
            Save-TbmBlastRecord -Blast $blast

            $blast.Recipients[0].Status = 'Sent'
            $blast.Recipients[0].SentUtc = [datetime]::new(2026, 11, 17, 9, 30, 15, 250, [System.DateTimeKind]::Utc)
            $blast.Recipients[0].ChatId = '19:abc@unq.gbl.spaces'
            Add-TbmBlastJournalEntry -Blast $blast -Index 0

            $path = Get-TbmBlastFilePath -Blast $blast
            $loaded = Read-TbmBlastRecord -Path $path
            $loaded.Id | Should -Be $blast.Id
            $loaded.Recipients[0].Status | Should -Be ([TeamsBlastMessenger.TbmRecipientStatus]::Sent)
            $loaded.Recipients[0].SentUtc | Should -Be $blast.Recipients[0].SentUtc
            $loaded.Recipients[0].SentUtc.Kind | Should -Be ([System.DateTimeKind]::Utc)
            $loaded.Recipients[0].ChatId | Should -Be '19:abc@unq.gbl.spaces'
            $loaded.Recipients[1].Status | Should -Be ([TeamsBlastMessenger.TbmRecipientStatus]::Pending)

            $blast.CompletedUtc = [datetime]::UtcNow
            Save-TbmBlastRecord -Blast $blast
            [System.IO.File]::Exists($path + '.journal') | Should -BeFalse
            (Get-TbmBlastHistory).Count | Should -BeGreaterOrEqual 1
        }
    }
}

Describe 'Graph client' {
    It 'retries 429 using Retry-After and returns parsed JSON' {
        InModuleScope TeamsBlastMessenger {
            $script:calls = 0
            Mock Get-TbmAccessToken { 'token' }
            Mock Wait-TbmDelay { $true }
            Mock Invoke-TbmHttpRequest {
                $script:calls++
                if ($script:calls -eq 1) { return [pscustomobject]@{ StatusCode = 429; Content = ''; RetryAfterSeconds = 3; RequestId = 'r1'; DurationMs = 1 } }
                return [pscustomobject]@{ StatusCode = 200; Content = '{"id":"42"}'; RetryAfterSeconds = 0; RequestId = 'r2'; DurationMs = 1 }
            }
            $auth = New-TbmAuthState; $auth.Mode = 'GraphRest'
            (Invoke-TbmGraphRequest -AuthState $auth -Path '/me').id | Should -Be '42'
            Should -Invoke Wait-TbmDelay -Times 1 -ParameterFilter { $Milliseconds -eq 3000 }
        }
    }

    It 'refreshes the token once on 401 and throws TbmGraphException with Graph error details' {
        InModuleScope TeamsBlastMessenger {
            Mock Get-TbmAccessToken { 'token' }
            Mock Invoke-TbmHttpRequest {
                if ($Uri.EndsWith('/me')) { return [pscustomobject]@{ StatusCode = 401; Content = ''; RetryAfterSeconds = 0; RequestId = ''; DurationMs = 1 } }
                return [pscustomobject]@{ StatusCode = 403; Content = '{"error":{"code":"Forbidden","message":"Nope"}}'; RetryAfterSeconds = 0; RequestId = 'req'; DurationMs = 1 }
            }
            $auth = New-TbmAuthState; $auth.Mode = 'GraphRest'
            { Invoke-TbmGraphRequest -AuthState $auth -Path '/me' } | Should -Throw -ExceptionType ([TeamsBlastMessenger.TbmGraphException])
            Should -Invoke Get-TbmAccessToken -Times 1 -ParameterFilter { $ForceRefresh }
            try { Invoke-TbmGraphRequest -AuthState $auth -Path '/chats' } catch { $caught = $_.Exception }
            $caught.StatusCode | Should -Be 403
            $caught.ErrorCode | Should -Be 'Forbidden'
            Get-TbmFailureMessage -Exception $caught | Should -Match 'Access denied'
        }
    }

    It 'reads the expiry of a JWT access token' {
        InModuleScope TeamsBlastMessenger {
            $payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes('{"exp":1893456000}')).TrimEnd('=').Replace('+', '-').Replace('/', '_')
            Get-TbmJwtExpiryUtc -Token ('h.' + $payload + '.s') | Should -Be ([datetime]::new(2030, 1, 1, 0, 0, 0, [System.DateTimeKind]::Utc))
        }
    }
}

Describe 'Send worker' {
    It 'sends individually, records chat/message ids and reports per-recipient failures' {
        InModuleScope TeamsBlastMessenger {
            Mock Resolve-TbmGraphUser {
                if ($Email -eq 'ghost@x.com') { return $null }
                [pscustomobject]@{ Id = 'id-' + $Email; DisplayName = 'Name ' + $Email; Upn = $Email; Mail = $Email }
            }
            Mock New-TbmGraphChat { 'chat-' + $MemberUserIds[0] }
            Mock Send-TbmGraphChatMessage { [pscustomobject]@{ Id = 'msg-' + $ChatId; CreatedUtc = [datetime]::new(2026, 11, 17, 10, 0, 0, [System.DateTimeKind]::Utc) } }

            $blast = [TeamsBlastMessenger.TbmBlastRecord]::new()
            $blast.Mode = 'Individual'; $blast.Transport = 'GraphRest'; $blast.Format = 'Text'; $blast.Template = 'x'
            $items = [System.Collections.Generic.List[hashtable]]::new()
            foreach ($email in @('a@x.com', 'ghost@x.com', 'c@x.com')) {
                $recipient = [TeamsBlastMessenger.TbmRecipientResult]::new(); $recipient.Email = $email
                $blast.Recipients.Add($recipient)
                $items.Add(@{ Index = $blast.Recipients.Count - 1; Email = $email; Message = 'Hello ' + $email })
            }
            $sync = New-TbmSyncState
            $auth = New-TbmAuthState; $auth.Mode = 'GraphRest'; $auth.MeId = 'me'
            Invoke-TbmSendWorker -Job @{ Blast = $blast; Items = $items; Settings = @{ PacingDelayMs = 0 }; AuthState = $auth; Sync = $sync }

            $blast.Recipients[0].Status | Should -Be ([TeamsBlastMessenger.TbmRecipientStatus]::Sent)
            $blast.Recipients[0].ChatId | Should -Be 'chat-id-a@x.com'
            $blast.Recipients[0].MessageId | Should -Be 'msg-chat-id-a@x.com'
            $blast.Recipients[1].Status | Should -Be ([TeamsBlastMessenger.TbmRecipientStatus]::Failed)
            $blast.Recipients[1].Error | Should -Match 'not found'
            $blast.Recipients[2].Status | Should -Be ([TeamsBlastMessenger.TbmRecipientStatus]::Sent)
            $blast.CompletedUtc | Should -Not -BeNullOrEmpty
            Should -Invoke Send-TbmGraphChatMessage -Times 1 -ParameterFilter { $HtmlContent -eq 'Hello a@x.com' }

            $events = [System.Collections.Generic.List[object]]::new(); $item = $null
            while ($sync.Events.TryDequeue([ref]$item)) { $events.Add($item) }
            $events[$events.Count - 1].Type | Should -Be 'Done'
            $events[$events.Count - 1].Sent | Should -Be 2

            $saved = Read-TbmBlastRecord -Path (Get-TbmBlastFilePath -Blast $blast)
            $saved.SentCount | Should -Be 2
            $saved.FailedCount | Should -Be 1
        }
    }

    It 'sends one group message to all resolvable members and cancels cleanly' {
        InModuleScope TeamsBlastMessenger {
            Mock Resolve-TbmGraphUser { [pscustomobject]@{ Id = 'id-' + $Email; DisplayName = $Email; Upn = $Email; Mail = $Email } }
            Mock New-TbmGraphChat { 'group-chat' }
            Mock Send-TbmGraphChatMessage { [pscustomobject]@{ Id = 'm1'; CreatedUtc = [datetime]::UtcNow } }
            $blast = [TeamsBlastMessenger.TbmBlastRecord]::new()
            $blast.Mode = 'Group'; $blast.Transport = 'GraphSdk'; $blast.Format = 'Text'; $blast.Topic = 'Team'
            $items = [System.Collections.Generic.List[hashtable]]::new()
            foreach ($email in @('a@x.com', 'b@x.com')) {
                $recipient = [TeamsBlastMessenger.TbmRecipientResult]::new(); $recipient.Email = $email
                $blast.Recipients.Add($recipient)
                $items.Add(@{ Index = $blast.Recipients.Count - 1; Email = $email; Message = 'Same for all' })
            }
            $auth = New-TbmAuthState; $auth.Mode = 'GraphSdk'; $auth.MeId = 'me'
            Invoke-TbmSendWorker -Job @{ Blast = $blast; Items = $items; Settings = @{ PacingDelayMs = 0 }; AuthState = $auth; Sync = (New-TbmSyncState) }
            Should -Invoke New-TbmGraphChat -Times 1 -ParameterFilter { $MemberUserIds.Count -eq 2 -and $Topic -eq 'Team' }
            Should -Invoke Send-TbmGraphChatMessage -Times 1
            $blast.SentCount | Should -Be 2
            $blast.Recipients[1].MessageId | Should -Be 'm1'

            $cancelled = [TeamsBlastMessenger.TbmBlastRecord]::new()
            $cancelled.Mode = 'Individual'; $cancelled.Transport = 'GraphSdk'; $cancelled.Format = 'Text'
            $recipient = [TeamsBlastMessenger.TbmRecipientResult]::new(); $recipient.Email = 'a@x.com'; $cancelled.Recipients.Add($recipient)
            $sync = New-TbmSyncState; $sync['Cancel'] = $true
            Invoke-TbmSendWorker -Job @{ Blast = $cancelled; Items = [System.Collections.Generic.List[hashtable]]@(@{ Index = 0; Email = 'a@x.com'; Message = 'x' }); Settings = @{ PacingDelayMs = 0 }; AuthState = $auth; Sync = $sync }
            $cancelled.Recipients[0].Status | Should -Be ([TeamsBlastMessenger.TbmRecipientStatus]::Cancelled)
        }
    }
}

Describe 'Interaction tracking' {
    BeforeAll {
        InModuleScope TeamsBlastMessenger {
            function script:New-TestMessage {
                param([string]$Id, [string]$From, [string]$Created, [string]$Html, [string]$Type = 'message', [object[]]$Attachments = @())
                [pscustomobject]@{
                    id = $Id; messageType = $Type; createdDateTime = $Created; deletedDateTime = $null
                    from = [pscustomobject]@{ user = [pscustomobject]@{ id = $From; displayName = 'User ' + $From } }
                    body = [pscustomobject]@{ contentType = 'html'; content = $Html }
                    attachments = $Attachments
                }
            }
        }
    }

    It 'attributes 1:1 replies after the send, flags quoted replies and computes response time' {
        InModuleScope TeamsBlastMessenger {
            $result = [TeamsBlastMessenger.TbmInteractionResult]::new()
            $result.Email = 'a@x.com'; $result.UserId = 'ua'; $result.ChatId = 'c1'
            $result.SentUtc = [datetime]::new(2026, 11, 17, 9, 0, 0, [System.DateTimeKind]::Utc)
            $results = [System.Collections.Generic.List[TeamsBlastMessenger.TbmInteractionResult]]::new(); $results.Add($result)
            $messages = [System.Collections.Generic.List[object]]::new()
            $messages.Add((New-TestMessage -Id 'm3' -From 'ua' -Created '2026-11-17T11:30:00.000Z' -Html '<p>Second</p>'))
            $messages.Add((New-TestMessage -Id 'm2' -From 'ua' -Created '2026-11-17T09:12:34.000Z' -Html '<p>Thanks &amp; noted</p>' -Attachments @([pscustomobject]@{ id = 'blast1'; contentType = 'messageReference' })))
            $messages.Add((New-TestMessage -Id 'm1' -From 'me' -Created '2026-11-17T09:05:00.000Z' -Html 'my follow-up'))
            $messages.Add((New-TestMessage -Id 'm0' -From 'ua' -Created '2026-11-17T08:59:00.000Z' -Html 'before the blast'))
            $messages.Add((New-TestMessage -Id 'sys' -From 'ua' -Created '2026-11-17T09:20:00.000Z' -Html '' -Type 'systemEventMessage'))
            $ids = [System.Collections.Generic.HashSet[string]]::new(); [void]$ids.Add('blast1')
            $reactions = [System.Collections.Generic.Dictionary[string, string]]::new(); $reactions['ua'] = 'like'

            Set-TbmInteractionReplies -Results $results -Messages $messages -MeId 'me' -BlastMessageIds $ids -Reactions $reactions

            $result.Responded | Should -BeTrue
            $result.ReplyCount | Should -Be 2
            $result.FirstReplyText | Should -Be 'Thanks & noted'
            $result.DirectReply | Should -BeTrue
            $result.ResponseTime | Should -Be ([System.TimeSpan]::new(0, 12, 34))
            $result.ResponseTimeDisplay | Should -Be '00:12:34'
            $result.Reactions | Should -Be 'like'
            $result.Replies[1].Text | Should -Be 'Second'
        }
    }

    It 'attributes group replies by sender and leaves non-responders unanswered' {
        InModuleScope TeamsBlastMessenger {
            $sent = [datetime]::new(2026, 11, 17, 9, 0, 0, [System.DateTimeKind]::Utc)
            $results = [System.Collections.Generic.List[TeamsBlastMessenger.TbmInteractionResult]]::new()
            foreach ($id in @('ua', 'ub')) {
                $result = [TeamsBlastMessenger.TbmInteractionResult]::new(); $result.UserId = $id; $result.Email = $id + '@x.com'; $result.SentUtc = $sent; $result.ChatId = 'g'
                $results.Add($result)
            }
            $messages = [System.Collections.Generic.List[object]]::new()
            $messages.Add((New-TestMessage -Id 'x' -From 'ub' -Created '2026-11-18T10:00:01.000Z' -Html 'Late reply'))
            $messages.Add((New-TestMessage -Id 'y' -From 'stranger' -Created '2026-11-17T10:00:00.000Z' -Html 'not a recipient'))
            Set-TbmInteractionReplies -Results $results -Messages $messages -MeId 'me' -BlastMessageIds ([System.Collections.Generic.HashSet[string]]::new()) -IsGroup $true
            $results[0].Responded | Should -BeFalse
            $results[1].Responded | Should -BeTrue
            $results[1].ResponseTimeDisplay | Should -Be '1d 01:00:01'
            Get-TbmInteractionSummary -Results $results | Should -Match '^1 of 2 responded'
        }
    }

    It 'finds the 1:1 chat for a deep-link blast and reads replies through Graph' {
        InModuleScope TeamsBlastMessenger {
            Mock Resolve-TbmGraphUser { [pscustomobject]@{ Id = 'ua'; DisplayName = 'Ada'; Upn = $Email; Mail = $Email } }
            Mock Invoke-TbmGraphPagedRequest {
                $list = [System.Collections.Generic.List[object]]::new()
                if ($Path.StartsWith('/me/chats')) {
                    $list.Add([pscustomobject]@{ id = 'chat-ua'; chatType = 'oneOnOne'; members = @([pscustomobject]@{ userId = 'me' }, [pscustomobject]@{ userId = 'ua' }) })
                }
                else {
                    $Path | Should -Match 'lastModifiedDateTime gt 2026-11-17T09:00:00.000Z'
                    $list.Add((New-TestMessage -Id 'r1' -From 'ua' -Created '2026-11-17T09:01:00.000Z' -Html 'On it'))
                }
                return , $list
            }
            Mock Invoke-TbmGraphRequest { $null }
            Mock Save-TbmBlastRecord { }

            $blast = [TeamsBlastMessenger.TbmBlastRecord]::new()
            $blast.Mode = 'Individual'; $blast.Transport = 'DeepLinkAuto'
            $recipient = [TeamsBlastMessenger.TbmRecipientResult]::new()
            $recipient.Email = 'ada@x.com'; $recipient.Status = 'Sent'; $recipient.SentUtc = [datetime]::new(2026, 11, 17, 9, 0, 0, [System.DateTimeKind]::Utc)
            $blast.Recipients.Add($recipient)
            $auth = New-TbmAuthState; $auth.Mode = 'GraphRest'; $auth.MeId = 'me'
            $sync = New-TbmSyncState
            Invoke-TbmInteractionWorker -Job @{ Blast = $blast; AuthState = $auth; Settings = @{ InteractionMaxChatPages = 2 }; Sync = $sync }

            $blast.Interactions.Count | Should -Be 1
            $blast.Interactions[0].ChatId | Should -Be 'chat-ua'
            $blast.Interactions[0].FirstReplyText | Should -Be 'On it'
            $blast.Interactions[0].ResponseTimeDisplay | Should -Be '00:01:00'
            $blast.LastInteractionCheckUtc | Should -Not -BeNullOrEmpty
        }
    }
}

Describe 'Exports' {
    It 'writes RFC 4180 CSV with interaction details' {
        InModuleScope TeamsBlastMessenger -Parameters @{ Root = $script:TestRoot } {
            param($Root)
            $result = [TeamsBlastMessenger.TbmInteractionResult]::new()
            $result.Email = 'a@x.com'; $result.Responded = $true; $result.ReplyCount = 1
            $result.SentUtc = [datetime]::new(2026, 1, 1, 9, 0, 0, [System.DateTimeKind]::Utc)
            $result.FirstReplyUtc = [datetime]::new(2026, 1, 1, 9, 30, 0, [System.DateTimeKind]::Utc)
            $result.FirstReplyText = 'Yes, "great", thanks'
            $list = [System.Collections.Generic.List[TeamsBlastMessenger.TbmInteractionResult]]::new(); $list.Add($result)
            $path = [System.IO.Path]::Combine($Root, 'interactions.csv')
            Export-TbmInteractionCsv -Interactions $list -Path $path
            $rows = Import-Csv -LiteralPath $path
            $rows[0].'Response Time (minutes)' | Should -Be '30'
            $rows[0].'First Reply' | Should -Be 'Yes, "great", thanks'
        }
    }
}

Describe 'XAML and UI wiring' {
    It 'every XAML file is well-formed and every control referenced in code exists' {
        $xamlFolder = [System.IO.Path]::Combine($script:ProjectRoot, 'Xaml')
        $names = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($file in [System.IO.Directory]::GetFiles($xamlFolder, '*.xaml')) {
            $document = [System.Xml.XmlDocument]::new()
            { $document.LoadXml([System.IO.File]::ReadAllText($file)) } | Should -Not -Throw
            foreach ($node in $document.SelectNodes('//*[@Name]')) { [void]$names.Add($node.GetAttribute('Name')) }
        }
        $code = ''
        foreach ($file in [System.IO.Directory]::GetFiles([System.IO.Path]::Combine($script:ProjectRoot, 'Private'), 'Ui*.ps1')) { $code += [System.IO.File]::ReadAllText($file) }
        $missing = [System.Collections.Generic.List[string]]::new()
        foreach ($match in [System.Text.RegularExpressions.Regex]::Matches($code, '\b(?:C|Controls|c)\.((?:Btn|Txt|Cmb|Grid|Lst|Rad|Prg|Ell|Chk)[A-Za-z]+)')) {
            $name = $match.Groups[1].Value
            if (-not $names.Contains($name) -and -not $missing.Contains($name)) { $missing.Add($name) }
        }
        foreach ($match in [System.Text.RegularExpressions.Regex]::Matches($code, "'((?:Btn|Txt|Cmb|Rad)[A-Za-z]+)'")) {
            $name = $match.Groups[1].Value
            if (-not $names.Contains($name) -and -not $missing.Contains($name)) { $missing.Add($name) }
        }
        $missing | Should -BeNullOrEmpty
    }
}
