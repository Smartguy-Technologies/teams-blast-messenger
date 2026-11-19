#Requires -Modules Pester
<#
    Unit tests for the non-UI logic. Run from the project folder:
        Invoke-Pester -Path .\Tests
    App data is redirected to a temp folder via TBM_APPDATA.
#>

BeforeAll {
    $script:TestRoot = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), 'tbm-tests-' + [guid]::NewGuid().ToString('N'))
    $null = [System.IO.Directory]::CreateDirectory($script:TestRoot)
    [System.Environment]::SetEnvironmentVariable('TBM_APPDATA', [System.IO.Path]::Combine($script:TestRoot, 'appdata'))
    $script:ProjectRoot = [System.IO.Path]::GetDirectoryName($PSScriptRoot)
    Import-Module ([System.IO.Path]::Combine($script:ProjectRoot, 'TeamsBlastMessenger.psd1')) -Force
    InModuleScope TeamsBlastMessenger { $null = Initialize-TbmEnvironment }
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
            $zoe = 'Zo' + [char]0x00EB
            $text = "Email;First Name;Note`r`n$($zoe.ToLowerInvariant())@contoso.com;$zoe;""multi`r`nline; with semicolon""`r`n;;`r`nbad;Bad;x`r`n"
            [System.IO.File]::WriteAllText($path, $text, [System.Text.UTF8Encoding]::new($true))
            $result = Import-TbmRecipientFile -Path $path
            $table = $result.Table
            $table.Rows.Count | Should -Be 2
            $table.Rows[0]['First Name'] | Should -Be $zoe
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

Describe 'Manually entered recipients' {
    It 'parses emails, display-name forms, Excel cells, multiple per line and flags problems' {
        InModuleScope TeamsBlastMessenger {
            $text = @(
                'ana@contoso.com',
                'Bob Stone <bob@contoso.com>, "Lee, Cara" <cara@contoso.com>',
                "Name`tEmail",
                "Dan Fox`tdan@contoso.com",
                'eve@contoso.com (Eve Moss); fay@contoso',
                'BOB@contoso.com',
                'Gil Hart'
            ) -join "`r`n"
            $entries = ConvertFrom-TbmManualRecipientText -Text $text
            $entries.Count | Should -Be 8
            $entries[0].Email | Should -Be 'ana@contoso.com'; $entries[0].Status | Should -Be 'OK'; $entries[0].Name | Should -Be ''
            $entries[1].Name | Should -Be 'Bob Stone'; $entries[1].FirstName | Should -Be 'Bob'; $entries[1].LastName | Should -Be 'Stone'
            $entries[2].Email | Should -Be 'cara@contoso.com'; $entries[2].Name | Should -Be 'Cara Lee'; $entries[2].LastName | Should -Be 'Lee'
            $entries[3].Name | Should -Be 'Dan Fox'; $entries[3].Email | Should -Be 'dan@contoso.com'
            $entries[4].Name | Should -Be 'Eve Moss'
            $entries[5].Status | Should -Be 'Invalid email'
            $entries[6].Status | Should -Be 'Duplicate'
            $entries[7].Status | Should -Be 'Needs email'; $entries[7].FirstName | Should -Be 'Gil'
        }
    }

    It 'looks up name-only entries: one match fills the email, several are ambiguous, none is not found' {
        InModuleScope TeamsBlastMessenger {
            Mock Invoke-TbmGraphRequest {
                if ($Path -match 'Gil%20Hart') { return [pscustomobject]@{ value = @([pscustomobject]@{ displayName = 'Gil Hart'; mail = 'gil@contoso.com'; userPrincipalName = 'gil@contoso.com' }) } }
                if ($Path -match 'filter' -and $Path -match 'Sam') { return [pscustomobject]@{ value = @() } }
                if ($Path -match 'search' -and $Path -match 'Sam') { return [pscustomobject]@{ value = @([pscustomobject]@{ displayName = 'Sam A'; mail = 'sa@contoso.com' }, [pscustomobject]@{ displayName = 'Sam B'; mail = $null; userPrincipalName = 'sb@contoso.com' }) } }
                return [pscustomobject]@{ value = @() }
            }
            $entries = ConvertFrom-TbmManualRecipientText -Text "Gil Hart`nSam`nNobody Here"
            $auth = New-TbmAuthState; $auth.Mode = 'GraphRest'; $auth.MeId = 'me'
            Resolve-TbmManualEntryNames -AuthState $auth -Entries $entries
            $entries[0].Status | Should -Be 'OK'; $entries[0].Email | Should -Be 'gil@contoso.com'
            $entries[1].Status | Should -Be 'Ambiguous'; $entries[1].Detail | Should -Match 'sb@contoso.com'
            $entries[2].Status | Should -Be 'Not found'
            Should -Invoke Invoke-TbmGraphRequest -ParameterFilter { $Headers -and $Headers.ConsistencyLevel -eq 'eventual' }
        }
    }

    It 'builds a new list, or appends to an imported one filling matching name columns' {
        InModuleScope TeamsBlastMessenger -Parameters @{ Root = $script:TestRoot } {
            param($Root)
            $entries = ConvertFrom-TbmManualRecipientText -Text "Jane Doe <jane@x.com>`nbad@`nkim@x.com"
            $new = Add-TbmManualEntriesToTable -Entries $entries
            $new.Added | Should -Be 2
            $new.RecipientColumn | Should -Be 'Email'
            (Get-TbmDataColumnNames -Table $new.Table) | Should -Be @('Email', 'Name', 'First Name', 'Last Name')
            $new.Table.Rows[0]['First Name'] | Should -Be 'Jane'
            [bool]$new.Table.Rows[1][$script:TbmIncludeColumn] | Should -BeTrue

            $path = [System.IO.Path]::Combine($Root, 'base.csv')
            [System.IO.File]::WriteAllText($path, "Work Email,Given Name,Surname,Dept`nold@x.com,Old,Timer,IT`n")
            $table = (Import-TbmRecipientFile -Path $path).Table
            $appended = Add-TbmManualEntriesToTable -Entries $entries -Table $table -RecipientColumn 'Work Email'
            $appended.Table.Rows.Count | Should -Be 3
            $appended.Table.Rows[1]['Work Email'] | Should -Be 'jane@x.com'
            $appended.Table.Rows[1]['Given Name'] | Should -Be 'Jane'
            $appended.Table.Rows[1]['Surname'] | Should -Be 'Doe'
            $appended.Table.Rows[1]['Dept'] | Should -Be ''
        }
    }
}

Describe 'Settings and organization policy' {
    AfterEach {
        InModuleScope TeamsBlastMessenger {
            $script:TbmPolicy = $null
            $path = Get-TbmAppDataPath -Name 'Settings'
            if ([System.IO.File]::Exists($path)) { [System.IO.File]::Delete($path) }
        }
    }

    It 'validates and normalises setting values' {
        InModuleScope TeamsBlastMessenger {
            (Test-TbmSettingValue -Name 'Transport' -Value 'graph').Value | Should -Be 'Graph'
            (Test-TbmSettingValue -Name 'Transport' -Value 'Teams').Valid | Should -BeFalse
            (Test-TbmSettingValue -Name 'PacingDelayMs' -Value ([int]2000)).Value | Should -Be 2000
            (Test-TbmSettingValue -Name 'PacingDelayMs' -Value ' 250 ').Value | Should -Be 250
            (Test-TbmSettingValue -Name 'PacingDelayMs' -Value 60001).Error | Should -Match 'from 0 to 60000'
            (Test-TbmSettingValue -Name 'PacingDelayMs' -Value '1.5').Valid | Should -BeFalse
            (Test-TbmSettingValue -Name 'ClientId' -Value 'not-a-guid').Valid | Should -BeFalse
            (Test-TbmSettingValue -Name 'ClientId' -Value '').Valid | Should -BeTrue
            (Test-TbmSettingValue -Name 'TenantId' -Value '').Value | Should -Be 'organizations'
            (Test-TbmSettingValue -Name 'TenantId' -Value 'contoso.com/x').Valid | Should -BeFalse
            (Test-TbmSettingValue -Name 'TestRecipient' -Value 'nope').Valid | Should -BeFalse
            (Test-TbmSettingValue -Name 'Bogus' -Value 1).Valid | Should -BeFalse
        }
    }

    It 'applies built-in < HKCU Defaults < HKLM Defaults < user < HKCU Enforced < HKLM Enforced, ignoring bad values' {
        InModuleScope TeamsBlastMessenger {
            Mock Read-TbmRegistryPolicy {
                switch ('{0}/{1}' -f $Hive, $Section) {
                    'CurrentUser/Defaults' { return @{ PacingDelayMs = 100; LogLevel = 'Warn'; DeepLinkTarget = 'Web' } }
                    'LocalMachine/Defaults' { return @{ PacingDelayMs = [int]200; clientid = '11111111-2222-3333-4444-555555555555' } }
                    'CurrentUser/Enforced' { return @{ Transport = 'Graph'; MessageFormat = 'Html' } }
                    'LocalMachine/Enforced' { return @{ Transport = 'DeepLinkManual'; LogRetentionDays = 0; Bogus = 'x'; LastFolder = 'C:\x' } }
                }
                return $null
            }
            Write-TbmTextFileAtomic -Path (Get-TbmAppDataPath -Name 'Settings') -Content '{ "PacingDelayMs": 300, "Transport": "Auto", "DeepLinkTarget": "Desktop", "LogLevel": "nonsense" }'

            $settings = Get-TbmSettings
            $settings.PacingDelayMs | Should -Be 300                 # user beats both Defaults
            $settings.DeepLinkTarget | Should -Be 'Desktop'          # user beats HKCU Defaults
            $settings.LogLevel | Should -Be 'Warn'                   # invalid user value ignored -> HKCU Default
            $settings.ClientId | Should -Be '11111111-2222-3333-4444-555555555555'   # case-insensitive value name
            $settings.Transport | Should -Be 'DeepLinkManual'        # HKLM Enforced beats HKCU Enforced and user
            $settings.MessageFormat | Should -Be 'Html'              # HKCU Enforced
            $settings.LogRetentionDays | Should -Be 30               # invalid enforced value ignored
            $settings.LastFolder | Should -Be ''                     # not policy-controllable
            Test-TbmSettingEnforced -Name 'Transport' | Should -BeTrue
            Test-TbmSettingEnforced -Name 'PacingDelayMs' | Should -BeFalse
            Get-TbmPolicySource -Name 'Transport' | Should -Be 'HKLM\SOFTWARE\Policies\TeamsBlastMessenger\Enforced'
            Get-TbmPolicySource -Name 'LogLevel' | Should -Be 'HKCU\SOFTWARE\Policies\TeamsBlastMessenger\Defaults'
        }
    }

    It 'saves only personal changes, so later policy default changes still reach the user' {
        InModuleScope TeamsBlastMessenger {
            $script:orgPacing = 200
            Mock Read-TbmRegistryPolicy {
                if ($Hive -eq 'LocalMachine' -and $Section -eq 'Defaults') { return @{ PacingDelayMs = $script:orgPacing } }
                if ($Hive -eq 'LocalMachine' -and $Section -eq 'Enforced') { return @{ Transport = 'Graph' } }
                return $null
            }
            $settings = Get-TbmSettings
            $settings.PacingDelayMs | Should -Be 200
            $settings.LogLevel = 'Debug'
            $settings.Transport = 'DeepLinkAuto'
            Save-TbmSettings -Settings $settings

            $saved = ConvertFrom-Json -InputObject ([System.IO.File]::ReadAllText((Get-TbmAppDataPath -Name 'Settings')))
            ($saved.PSObject.Properties.Name) | Should -Be @('LogLevel')

            $script:orgPacing = 750
            $reloaded = Get-TbmSettings
            $reloaded.PacingDelayMs | Should -Be 750
            $reloaded.LogLevel | Should -Be 'Debug'
            $reloaded.Transport | Should -Be 'Graph'
        }
    }

    It 'ships ADMX/ADML templates and an example .reg that match the settings schema' {
        InModuleScope TeamsBlastMessenger -Parameters @{ Root = $script:ProjectRoot } {
            param($Root)
            $policyFolder = [System.IO.Path]::Combine($Root, 'Policy')
            $admx = [System.Xml.XmlDocument]::new(); $admx.Load([System.IO.Path]::Combine($policyFolder, 'TeamsBlastMessenger.admx'))
            $adml = [System.Xml.XmlDocument]::new(); $adml.Load([System.IO.Path]::Combine($policyFolder, 'en-US', 'TeamsBlastMessenger.adml'))
            $ns = [System.Xml.XmlNamespaceManager]::new($admx.NameTable); $ns.AddNamespace('p', 'http://schemas.microsoft.com/GroupPolicy/2006/07/PolicyDefinitions')
            $nsl = [System.Xml.XmlNamespaceManager]::new($adml.NameTable); $nsl.AddNamespace('p', 'http://schemas.microsoft.com/GroupPolicy/2006/07/PolicyDefinitions')

            $strings = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($node in $adml.SelectNodes('//p:stringTable/p:string', $nsl)) { $null = $strings.Add($node.GetAttribute('id')) }
            $presentations = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($node in $adml.SelectNodes('//p:presentationTable/p:presentation', $nsl)) { $null = $presentations.Add($node.GetAttribute('id')) }
            foreach ($match in [System.Text.RegularExpressions.Regex]::Matches($admx.OuterXml, '\$\((string|presentation)\.([A-Za-z0-9_]+)\)')) {
                if ($match.Groups[1].Value -eq 'string') { $strings.Contains($match.Groups[2].Value) | Should -BeTrue -Because $match.Value }
                else { $presentations.Contains($match.Groups[2].Value) | Should -BeTrue -Because $match.Value }
            }

            $schema = Get-TbmSettingSchema
            $expected = [System.Collections.Generic.List[string]]::new()
            foreach ($name in $schema.Keys) { if ($schema[$name].Policy) { $expected.Add($name) } }
            foreach ($section in @('Defaults', 'Enforced')) {
                $found = [System.Collections.Generic.List[string]]::new()
                foreach ($policy in $admx.SelectNodes(('//p:policy[@key="SOFTWARE\Policies\TeamsBlastMessenger\{0}"]' -f $section), $ns)) {
                    foreach ($element in $policy.SelectNodes('p:elements/*', $ns)) {
                        $name = $element.GetAttribute('valueName')
                        $found.Add($name)
                        if ($element.LocalName -eq 'enum') {
                            foreach ($value in $element.SelectNodes('.//p:string', $ns)) { (Test-TbmSettingValue -Name $name -Value $value.InnerText).Valid | Should -BeTrue -Because "$name=$($value.InnerText)" }
                        }
                        if ($element.LocalName -eq 'decimal') {
                            [int]$element.GetAttribute('minValue') | Should -Be $schema[$name].Min
                            [int]$element.GetAttribute('maxValue') | Should -Be $schema[$name].Max
                        }
                    }
                }
                $found.Sort(); $expected.Sort()
                $found | Should -Be $expected
            }

            $regPath = [System.IO.Path]::Combine($policyFolder, 'Example-Policy.reg')
            [System.IO.File]::ReadAllLines($regPath)[0] | Should -Be 'Windows Registry Editor Version 5.00'
            foreach ($line in [System.IO.File]::ReadAllLines($regPath)) {
                $match = [System.Text.RegularExpressions.Regex]::Match($line, '^"([^"]+)"=(?:"([^"]*)"|dword:([0-9a-fA-F]{8}))$')
                if (-not $match.Success) { continue }
                $value = $match.Groups[2].Value
                if ($match.Groups[3].Success) { $value = [System.Convert]::ToInt32($match.Groups[3].Value, 16) }
                (Test-TbmSettingValue -Name (Resolve-TbmSettingName -Name $match.Groups[1].Value) -Value $value).Valid | Should -BeTrue -Because $line
            }
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

Describe 'Graph batching' {
    It 'sends up to 20 requests per batch call and retries only throttled items after Retry-After' {
        InModuleScope TeamsBlastMessenger {
            $script:batchCalls = [System.Collections.Generic.List[object]]::new()
            Mock Wait-TbmDelay { $true }
            Mock Invoke-TbmGraphRequest {
                $Path | Should -Be '/$batch'
                $script:batchCalls.Add($Body.requests)
                $responses = [System.Collections.Generic.List[object]]::new()
                foreach ($entry in $Body.requests) {
                    $status = 200
                    if ($entry.id -eq '1') { $status = 404 }
                    if ($entry.id -eq '2' -and $script:batchCalls.Count -eq 1) { $status = 429 }
                    $responses.Add([pscustomobject]@{ id = $entry.id; status = $status; headers = [pscustomobject]@{ 'Retry-After' = '3' }; body = [pscustomobject]@{ url = $entry.url } })
                }
                return [pscustomobject]@{ responses = $responses.ToArray() }
            }
            $requests = [System.Collections.Generic.List[hashtable]]::new()
            for ($i = 0; $i -lt 25; $i++) { $requests.Add(@{ Id = [string]$i; Path = ('https://graph.microsoft.com/v1.0/users/u{0}?$filter=a eq b' -f $i) }) }

            $answers = Invoke-TbmGraphBatch -AuthState (New-TbmAuthState) -Requests $requests

            $answers.Count | Should -Be 25
            $script:batchCalls.Count | Should -Be 3
            $script:batchCalls[0].Count | Should -Be 20
            $script:batchCalls[1].Count | Should -Be 5
            $script:batchCalls[2].Count | Should -Be 1
            $script:batchCalls[2][0].id | Should -Be '2'
            $script:batchCalls[0][0].url | Should -Be '/users/u0?$filter=a%20eq%20b'
            $answers['1'].Status | Should -Be 404
            $answers['2'].Status | Should -Be 200
            Should -Invoke Wait-TbmDelay -Times 1 -Exactly -ParameterFilter { $Milliseconds -eq 3000 }
        }
    }

    It 'resolves users in bulk: UPN first, then mail, caching misses and leaving errors to the per-user path' {
        InModuleScope TeamsBlastMessenger {
            Mock Invoke-TbmGraphBatch {
                $answers = [System.Collections.Generic.Dictionary[string, object]]::new()
                foreach ($request in $Requests) {
                    $status = 200; $body = $null
                    switch -Regex ($request.Path) {
                        '^/users/a%40x\.com' { $body = [pscustomobject]@{ id = 'ua'; displayName = 'Ada'; userPrincipalName = 'a@x.com'; mail = 'a@x.com' } }
                        '^/users/(b|n)%40x\.com' { $status = 404 }
                        '^/users/c%40x\.com' { $status = 500; $body = [pscustomobject]@{ error = [pscustomobject]@{ code = 'boom'; message = 'Server error' } } }
                        'mail%20eq%20%27b%40x\.com%27' { $body = [pscustomobject]@{ value = @([pscustomobject]@{ id = 'ub'; displayName = 'Bo'; userPrincipalName = 'bo@corp'; mail = 'b@x.com' }) } }
                        'mail%20eq%20%27n%40x\.com%27' { $body = [pscustomobject]@{ value = @() } }
                    }
                    $answers[$request.Id] = [pscustomobject]@{ Id = $request.Id; Status = $status; Body = $body; Headers = $null }
                }
                return , $answers
            }
            $cache = [System.Collections.Generic.Dictionary[string, object]]::new([System.StringComparer]::OrdinalIgnoreCase)
            Resolve-TbmGraphUsers -AuthState (New-TbmAuthState) -Emails @('A@x.com', 'a@x.com ', 'b@x.com', 'c@x.com', 'n@x.com', '') -Cache $cache

            Should -Invoke Invoke-TbmGraphBatch -Times 2 -Exactly
            Should -Invoke Invoke-TbmGraphBatch -Times 1 -Exactly -ParameterFilter { $Requests.Count -eq 4 }
            $cache['a@x.com'].Id | Should -Be 'ua'
            $cache['b@x.com'].Id | Should -Be 'ub'
            $cache.ContainsKey('n@x.com') | Should -BeTrue
            $cache['n@x.com'] | Should -BeNullOrEmpty
            $cache.ContainsKey('c@x.com') | Should -BeFalse
        }
    }

    It 'counts Graph request time toward the pacing delay, but always pauses fully between deep links' {
        InModuleScope TeamsBlastMessenger {
            Mock Resolve-TbmGraphUsers { }
            Mock Send-TbmBlastItem { [System.Threading.Thread]::Sleep(300) }
            Mock Set-TbmRecipientOutcome { }
            Mock Wait-TbmDelay { $true }
            $items = [System.Collections.Generic.List[hashtable]]::new()
            $blast = [TeamsBlastMessenger.TbmBlastRecord]::new()
            foreach ($email in @('a@x.com', 'b@x.com')) {
                $recipient = [TeamsBlastMessenger.TbmRecipientResult]::new(); $recipient.Email = $email; $blast.Recipients.Add($recipient)
                $items.Add(@{ Index = $blast.Recipients.Count - 1; Email = $email; Message = 'x' })
            }

            $blast.Transport = 'GraphRest'
            Invoke-TbmIndividualSend -Job @{ Blast = $blast; Items = $items; Settings = @{ PacingDelayMs = 1000 }; AuthState = (New-TbmAuthState); Sync = (New-TbmSyncState) }
            Should -Invoke Wait-TbmDelay -Times 1 -Exactly -ParameterFilter { $Milliseconds -ge 0 -and $Milliseconds -le 750 }

            $blast.Transport = 'DeepLinkAuto'
            Invoke-TbmIndividualSend -Job @{ Blast = $blast; Items = $items; Settings = @{ PacingDelayMs = 1000 }; AuthState = (New-TbmAuthState); Sync = (New-TbmSyncState) }
            Should -Invoke Wait-TbmDelay -Times 1 -Exactly -ParameterFilter { $Milliseconds -eq 1500 }
        }
    }

    It 'probes the server-side message filter once, then pages newest-first for every chat after a 400' {
        InModuleScope TeamsBlastMessenger {
            $script:filteredRequests = 0
            $script:roundSizes = [System.Collections.Generic.List[int]]::new()
            Mock Invoke-TbmGraphBatch {
                $script:roundSizes.Add($Requests.Count)
                $answers = [System.Collections.Generic.Dictionary[string, object]]::new()
                foreach ($request in $Requests) {
                    if ($request.Path.Contains('$filter')) {
                        $script:filteredRequests++
                        $answers[$request.Id] = [pscustomobject]@{ Id = $request.Id; Status = 400; Body = $null; Headers = $null }
                        continue
                    }
                    if ($request.Path.StartsWith('/chats/c3/')) {
                        $answers[$request.Id] = [pscustomobject]@{ Id = $request.Id; Status = 403; Body = [pscustomobject]@{ error = [pscustomobject]@{ code = 'Forbidden'; message = 'No access' } }; Headers = $null }
                        continue
                    }
                    $value = @(
                        [pscustomobject]@{ id = 'new'; createdDateTime = '2026-11-17T10:00:00Z' },
                        [pscustomobject]@{ id = 'old'; createdDateTime = '2026-11-17T08:00:00Z' }
                    )
                    $answers[$request.Id] = [pscustomobject]@{ Id = $request.Id; Status = 200; Body = [pscustomobject]@{ value = $value; '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/next' }; Headers = $null }
                }
                return , $answers
            }
            $since = [datetime]::new(2026, 11, 17, 9, 0, 0, [System.DateTimeKind]::Utc)
            $chats = [System.Collections.Generic.List[hashtable]]::new()
            foreach ($id in @('c1', 'c2', 'c3')) { $chats.Add(@{ ChatId = $id; SinceUtc = $since }) }

            $read = Get-TbmChatMessagesBatch -AuthState (New-TbmAuthState) -Chats $chats

            $script:filteredRequests | Should -Be 1
            $script:roundSizes[0] | Should -Be 1
            $read['c1'].Messages.Count | Should -Be 1
            $read['c1'].Messages[0].id | Should -Be 'new'
            $read['c2'].Messages.Count | Should -Be 1
            $read['c1'].Error | Should -BeNullOrEmpty
            $read['c3'].Error | Should -Match 'Forbidden'
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
            Mock Resolve-TbmGraphUsers { }
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
            Mock Resolve-TbmGraphUsers { }
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
            $ids = [System.Collections.Generic.HashSet[string]]::new(); $null = $ids.Add('blast1')
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
                $Path | Should -Match '^/me/chats'
                $list.Add([pscustomobject]@{ id = 'chat-ua'; chatType = 'oneOnOne'; members = @([pscustomobject]@{ userId = 'me' }, [pscustomobject]@{ userId = 'ua' }) })
                return , $list
            }
            Mock Invoke-TbmGraphBatch {
                $answers = [System.Collections.Generic.Dictionary[string, object]]::new()
                foreach ($request in $Requests) {
                    $request.Path | Should -Match '^/chats/chat-ua/messages\?.*lastModifiedDateTime gt 2026-11-17T09:00:00.000Z'
                    $body = [pscustomobject]@{ value = @((New-TestMessage -Id 'r1' -From 'ua' -Created '2026-11-17T09:01:00.000Z' -Html 'On it')) }
                    $answers[$request.Id] = [pscustomobject]@{ Id = $request.Id; Status = 200; Body = $body; Headers = $null }
                }
                return , $answers
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
            foreach ($node in $document.SelectNodes('//*[@Name]')) { $null = $names.Add($node.GetAttribute('Name')) }
        }
        $codeBuilder = [System.Text.StringBuilder]::new()
        foreach ($file in [System.IO.Directory]::GetFiles([System.IO.Path]::Combine($script:ProjectRoot, 'Private'), 'Ui*.ps1')) { $null = $codeBuilder.Append([System.IO.File]::ReadAllText($file)) }
        $code = $codeBuilder.ToString()
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

Describe 'Performance helpers' {
    It 'lists history from current summaries and falls back to the full record when a summary is stale or missing' {
        InModuleScope TeamsBlastMessenger {
            $blast = [TeamsBlastMessenger.TbmBlastRecord]::new()
            $blast.Mode = 'Individual'; $blast.Transport = 'GraphRest'; $blast.Template = 'Summary test'; $blast.Format = 'Text'
            foreach ($status in @('Sent', 'Sent', 'Failed')) {
                $recipient = [TeamsBlastMessenger.TbmRecipientResult]::new(); $recipient.Email = 'x@x.com'; $recipient.Status = $status
                $blast.Recipients.Add($recipient)
            }
            Save-TbmBlastRecord -Blast $blast
            $path = Get-TbmBlastFilePath -Blast $blast
            [System.IO.File]::Exists($path + '.summary') | Should -BeTrue

            $find = {
                foreach ($record in (Get-TbmBlastHistory)) { if ($record.Id -eq $blast.Id) { return $record } }
                return $null
            }
            $listed = & $find
            $listed.IsSummaryOnly | Should -BeTrue
            $listed.SentCount | Should -Be 2
            $listed.FailedCount | Should -Be 1
            $listed.RecipientCount | Should -Be 3
            $listed.Recipients.Count | Should -Be 0
            $listed.Template | Should -Be 'Summary test'
            $null = Expand-TbmBlastRecord -Blast $listed
            $listed.IsSummaryOnly | Should -BeFalse
            $listed.Recipients.Count | Should -Be 3
            $listed.SentCount | Should -Be 2

            # A journal (send in progress or interrupted) always means the full record.
            $blast.Recipients[2].Status = 'Sent'
            Add-TbmBlastJournalEntry -Blast $blast -Index 2
            $listed = & $find
            $listed.IsSummaryOnly | Should -BeFalse
            $listed.SentCount | Should -Be 3
            [System.IO.File]::Delete($path + '.journal')

            # The .json changed after the summary was written: summary ignored.
            [System.IO.File]::AppendAllText($path, ' ')
            (& $find).IsSummaryOnly | Should -BeFalse

            # Missing summary (older history): full read, then the summary is backfilled.
            [System.IO.File]::Delete($path + '.summary')
            (& $find).IsSummaryOnly | Should -BeFalse
            [System.IO.File]::Exists($path + '.summary') | Should -BeTrue
            (& $find).IsSummaryOnly | Should -BeTrue

            Remove-TbmBlastRecord -Blast $blast -Confirm:$false
            [System.IO.File]::Exists($path + '.summary') | Should -BeFalse
        }
    }

    It 'writes validation cells only when they change' {
        InModuleScope TeamsBlastMessenger {
            $table = [System.Data.DataTable]::new()
            $null = $table.Columns.Add('Email', [string])
            $null = $table.Columns.Add($script:TbmIncludeColumn, [bool])
            $null = $table.Columns.Add($script:TbmStatusColumn, [string])
            foreach ($email in @('a@x.com', 'bad', 'A@x.com', 'c@x.com')) { $null = $table.Rows.Add($email, $true, '') }
            $script:changes = 0
            $table.add_ColumnChanged({ $script:changes++ })

            $summary = Update-TbmRecipientValidation -Table $table -RecipientColumn 'Email'
            $summary.Invalid | Should -Be 1
            $summary.Duplicates | Should -Be 1
            $script:changes | Should -Be 4
            $table.Rows[1][$script:TbmStatusColumn] | Should -Be 'Invalid email'
            $table.Rows[2][$script:TbmIncludeColumn] | Should -BeFalse

            $script:changes = 0
            $null = Update-TbmRecipientValidation -Table $table -RecipientColumn 'Email'
            $script:changes | Should -Be 0

            $table.Rows[1]['Email'] = 'b@x.com'
            $script:changes = 0
            $null = Update-TbmRecipientValidation -Table $table -RecipientColumn 'Email'
            $script:changes | Should -Be 2
            $table.Rows[1][$script:TbmStatusColumn] | Should -Be ''
            $table.Rows[1][$script:TbmIncludeColumn] | Should -BeTrue
        }
    }

    It 'keeps the log file open between entries and survives closing and reopening it' {
        $marker = [guid]::NewGuid().ToString('N')
        [TeamsBlastMessenger.TbmLogger]::Info('Test', 'first ' + $marker)
        [TeamsBlastMessenger.TbmLogger]::Close()
        [TeamsBlastMessenger.TbmLogger]::Info('Test', 'second ' + $marker)
        $stream = [System.IO.FileStream]::new([TeamsBlastMessenger.TbmLogger]::CurrentLogPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try { $text = [System.IO.StreamReader]::new($stream).ReadToEnd() } finally { $stream.Dispose() }
        $text | Should -Match ('first ' + $marker)
        $text | Should -Match ('second ' + $marker)
        $null = [TeamsBlastMessenger.TbmLogger]::DrainPending()
    }

    It 'sorts with the compiled comparers and reads JSON fields and dates without PowerShell calls' {
        $records = [System.Collections.Generic.List[TeamsBlastMessenger.TbmBlastRecord]]::new()
        foreach ($day in @(3, 1, 2)) { $record = [TeamsBlastMessenger.TbmBlastRecord]::new(); $record.CreatedUtc = [datetime]::new(2026, 11, $day, 0, 0, 0, [System.DateTimeKind]::Utc); $records.Add($record) }
        $records.Sort([TeamsBlastMessenger.TbmBlastRecordByCreatedDesc]::new())
        $records[0].CreatedUtc.Day | Should -Be 3
        $records[2].CreatedUtc.Day | Should -Be 1

        $replies = [System.Collections.Generic.List[TeamsBlastMessenger.TbmReplyMessage]]::new()
        foreach ($minute in @(30, 10, 20)) { $reply = [TeamsBlastMessenger.TbmReplyMessage]::new(); $reply.CreatedUtc = [datetime]::new(2026, 11, 17, 9, $minute, 0); $replies.Add($reply) }
        $replies.Sort([TeamsBlastMessenger.TbmReplyByCreated]::new())
        $replies[0].CreatedUtc.Minute | Should -Be 10

        $message = ConvertFrom-Json -InputObject '{ "from": { "user": { "id": "u1" } }, "deletedDateTime": null, "attachments": [] }'
        [TeamsBlastMessenger.TbmObject]::GetValue($message, 'from.user.id') | Should -Be 'u1'
        [TeamsBlastMessenger.TbmObject]::GetValue($message, 'from.application.id') | Should -BeNullOrEmpty
        [TeamsBlastMessenger.TbmObject]::GetValue($message, 'deletedDateTime') | Should -BeNullOrEmpty
        [TeamsBlastMessenger.TbmObject]::GetValue(@{ a = @{ b = 5 } }, 'a.b') | Should -Be 5
        [TeamsBlastMessenger.TbmObject]::GetValue($null, 'a') | Should -BeNullOrEmpty

        [TeamsBlastMessenger.TbmObject]::ToUtc('2026-11-17T09:30:00+01:00') | Should -Be ([datetime]::new(2026, 11, 17, 8, 30, 0, [System.DateTimeKind]::Utc))
        ([TeamsBlastMessenger.TbmObject]::ToUtc([datetime]::new(2026, 1, 2, 3, 4, 5))).Kind | Should -Be ([System.DateTimeKind]::Utc)
        [TeamsBlastMessenger.TbmObject]::ToUtc(' ') | Should -BeNullOrEmpty
        [TeamsBlastMessenger.TbmText]::IsEmailAddress(' ada@contoso.com ') | Should -BeTrue
        [TeamsBlastMessenger.TbmText]::IsEmailAddress('Ada <ada@contoso.com>') | Should -BeFalse
        [TeamsBlastMessenger.TbmText]::IsEmailAddress('ada@localhost') | Should -BeFalse
        [TeamsBlastMessenger.TbmText]::Truncate('abcdefghij', 6) | Should -Be 'abc...'
    }
}

Describe 'Coding standards' {
    BeforeAll {
        $script:SourceFiles = [System.Collections.Generic.List[string]]::new()
        $script:ScriptFiles = [System.Collections.Generic.List[string]]::new()
        $skip = [System.IO.Path]::DirectorySeparatorChar + '.git' + [System.IO.Path]::DirectorySeparatorChar
        foreach ($file in [System.IO.Directory]::GetFiles($script:ProjectRoot, '*', [System.IO.SearchOption]::AllDirectories)) {
            if ($file.Contains($skip)) { continue }
            $script:SourceFiles.Add($file)
            $extension = [System.IO.Path]::GetExtension($file).ToLowerInvariant()
            # Benchmarks.ps1 measures the slow idioms on purpose; it is still checked for ASCII.
            if ([System.IO.Path]::GetFileName($file) -eq 'Benchmarks.ps1') { continue }
            if ($extension -eq '.ps1' -or $extension -eq '.psm1' -or $extension -eq '.psd1') { $script:ScriptFiles.Add($file) }
        }

        function Get-TbmStandardsViolation {
            param([string]$Pattern)
            $found = [System.Collections.Generic.List[string]]::new()
            $regex = [System.Text.RegularExpressions.Regex]::new($Pattern)
            foreach ($file in $script:ScriptFiles) {
                $lines = [System.IO.File]::ReadAllLines($file)
                for ($i = 0; $i -lt $lines.Length; $i++) {
                    $trimmed = $lines[$i].TrimStart()
                    if ($trimmed.StartsWith('#')) { continue }
                    if ($regex.IsMatch($lines[$i])) { $found.Add(('{0}:{1}: {2}' -f [System.IO.Path]::GetFileName($file), ($i + 1), $trimmed)) }
                }
            }
            return , $found
        }
    }

    It 'keeps every source file ASCII (an optional UTF-8 BOM is allowed)' {
        $offenders = [System.Collections.Generic.List[string]]::new()
        foreach ($file in $script:SourceFiles) {
            $bytes = [System.IO.File]::ReadAllBytes($file)
            $start = 0
            if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { $start = 3 }
            for ($i = $start; $i -lt $bytes.Length; $i++) {
                if ($bytes[$i] -gt 0x7F) { $offenders.Add(('{0} (byte {1})' -f $file, $i)); break }
            }
        }
        $offenders | Should -BeNullOrEmpty
    }

    It 'discards output by assigning to $null, never with a void cast' {
        Get-TbmStandardsViolation -Pattern '\[void\]' | Should -BeNullOrEmpty
    }

    It 'does not build strings or arrays with the add-assign operator' {
        Get-TbmStandardsViolation -Pattern '\+=\s*[''"@(\[]' | Should -BeNullOrEmpty
    }

    It 'does not pipe into Out-Null, ForEach-Object, Where-Object, Select-Object or Sort-Object' {
        Get-TbmStandardsViolation -Pattern '\|\s*(?:Out-Null|ForEach-Object|Where-Object|Select-Object|Sort-Object|%|\?)(?:\s|$)' | Should -BeNullOrEmpty
    }
}
