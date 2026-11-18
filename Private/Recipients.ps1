#region Recipient spreadsheet loading
<#
    Every source (CSV, XLSX via native OpenXML reader, ImportExcel module or Excel COM) is normalised to a
    List[string[]] of raw rows, then converted to a System.Data.DataTable whose first two columns are the
    tool's own Include / Status columns.
#>

$script:TbmIncludeColumn = 'TBM_Include'
$script:TbmStatusColumn = 'TBM_Status'
$script:TbmSpreadsheetNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:TbmOfficeRelNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:TbmPackageRelNs = 'http://schemas.openxmlformats.org/package/2006/relationships'

function Import-TbmRecipientFile {
    <#
    .SYNOPSIS
        Loads a CSV/TSV/TXT/XLSX/XLSM/XLS file into a recipient DataTable.
    .OUTPUTS
        PSCustomObject: Table (DataTable), Sheets (string[]), SheetName (string), Reader (string).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter()]
        [string]$SheetName
    )

    if (-not [System.IO.File]::Exists($Path)) { throw [System.IO.FileNotFoundException]::new('File not found.', $Path) }
    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    $sheets = [string[]]@()
    $reader = ''
    $rows = $null

    switch ($extension) {
        { $_ -in @('.csv', '.txt', '.tsv') } {
            $rows = Read-TbmDelimitedFile -Path $Path
            $reader = 'Delimited text'
        }
        { $_ -in @('.xlsx', '.xlsm') } {
            try {
                $workbook = Read-TbmXlsxNative -Path $Path -SheetName $SheetName
                $reader = 'OpenXML'
            }
            catch {
                Write-TbmLog -Level Warn -Source 'Recipients' -Message ('Native XLSX reader failed ({0}); trying fallbacks.' -f $_.Exception.Message)
                $workbook = Read-TbmWorkbookFallback -Path $Path -SheetName $SheetName
                $reader = $workbook.Reader
            }
            $rows = $workbook.Rows
            $sheets = $workbook.Sheets
            $SheetName = $workbook.SheetName
        }
        '.xls' {
            $workbook = Read-TbmWorkbookFallback -Path $Path -SheetName $SheetName
            $rows = $workbook.Rows
            $sheets = $workbook.Sheets
            $SheetName = $workbook.SheetName
            $reader = $workbook.Reader
        }
        default { throw [System.NotSupportedException]::new(('Unsupported file type "{0}". Use CSV or XLSX.' -f $extension)) }
    }

    $table = ConvertTo-TbmRecipientTable -Rows $rows
    Write-TbmLog -Level Info -Source 'Recipients' -Message ('Loaded {0} row(s), {1} column(s) from {2} [{3}{4}].' -f $table.Rows.Count, ($table.Columns.Count - 2), $Path, $reader, $(if ($SheetName) { ', sheet ' + $SheetName } else { '' }))
    return [pscustomobject]@{
        Table     = $table
        Sheets    = $sheets
        SheetName = $SheetName
        Reader    = $reader
    }
}

function ConvertTo-TbmRecipientTable {
    <#
    .SYNOPSIS
        Converts raw rows (first non-blank row = header) to a DataTable with Include/Status columns.
    #>
    [CmdletBinding()]
    [OutputType([System.Data.DataTable])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string[]]]$Rows
    )

    $table = [System.Data.DataTable]::new('Recipients')
    $includeColumn = $table.Columns.Add($script:TbmIncludeColumn, [bool])
    $includeColumn.DefaultValue = $true
    $statusColumn = $table.Columns.Add($script:TbmStatusColumn, [string])
    $statusColumn.DefaultValue = ''

    $headerIndex = -1
    for ($i = 0; $i -lt $Rows.Count; $i++) {
        if (-not (Test-TbmBlankRow -Values $Rows[$i])) { $headerIndex = $i; break }
    }
    if ($headerIndex -lt 0) { return , $table }

    # Width = last column holding any value in the header OR the data, so a data column with a blank header is kept.
    $width = 0
    for ($i = $headerIndex; $i -lt $Rows.Count; $i++) {
        $values = $Rows[$i]
        if ($null -eq $values) { continue }
        for ($c = $values.Length - 1; $c -ge $width; $c--) {
            if (-not [string]::IsNullOrWhiteSpace($values[$c])) { $width = $c + 1; break }
        }
    }
    $headers = Get-TbmUniqueHeaderNames -Header $Rows[$headerIndex] -Width $width
    foreach ($header in $headers) { $null = $table.Columns.Add($header, [string]) }

    $table.BeginLoadData()
    try {
        for ($i = $headerIndex + 1; $i -lt $Rows.Count; $i++) {
            $values = $Rows[$i]
            if (Test-TbmBlankRow -Values $values) { continue }
            $row = $table.NewRow()
            $row[$script:TbmIncludeColumn] = $true
            $row[$script:TbmStatusColumn] = ''
            for ($c = 0; $c -lt $headers.Count; $c++) {
                $value = ''
                if ($c -lt $values.Length -and $null -ne $values[$c]) { $value = $values[$c].Trim() }
                $row[$c + 2] = $value
            }
            $table.Rows.Add($row)
        }
    }
    finally {
        $table.EndLoadData()
    }
    return , $table
}

function Test-TbmBlankRow {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter()]
        [AllowNull()]
        [AllowEmptyCollection()]
        [string[]]$Values
    )
    if ($null -eq $Values) { return $true }
    foreach ($value in $Values) { if (-not [string]::IsNullOrWhiteSpace($value)) { return $false } }
    return $true
}

function Get-TbmUniqueHeaderNames {
    <#
    .SYNOPSIS
        Trims header names, names blank headers "Column N" and de-duplicates ("Name", "Name (2)").
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[string]])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [AllowEmptyString()]
        [AllowNull()]
        [string[]]$Header,

        [Parameter(Mandatory)]
        [ValidateRange(0, 16384)]
        [int]$Width
    )

    $names = [System.Collections.Generic.List[string]]::new()
    $used = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $null = $used.Add($script:TbmIncludeColumn)
    $null = $used.Add($script:TbmStatusColumn)
    for ($i = 0; $i -lt $Width; $i++) {
        $name = ''
        if ($i -lt $Header.Length -and $null -ne $Header[$i]) { $name = $Header[$i].Trim() }
        if ([string]::IsNullOrEmpty($name)) { $name = 'Column {0}' -f ($i + 1) }
        $candidate = $name
        $suffix = 2
        while (-not $used.Add($candidate)) {
            $candidate = '{0} ({1})' -f $name, $suffix
            $suffix++
        }
        $names.Add($candidate)
    }
    return , $names
}

function Get-TbmDataColumnNames {
    <#
    .SYNOPSIS
        Returns the spreadsheet's own column names (excludes the tool's Include/Status columns).
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Data.DataTable]$Table
    )

    $names = [System.Collections.Generic.List[string]]::new()
    foreach ($column in $Table.Columns) {
        if ($column.ColumnName -eq $script:TbmIncludeColumn -or $column.ColumnName -eq $script:TbmStatusColumn) { continue }
        $names.Add($column.ColumnName)
    }
    return , $names.ToArray()
}

function Find-TbmRecipientColumn {
    <#
    .SYNOPSIS
        Guesses which column holds recipient email/UPN: known header names first, then the column with most '@' values.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Data.DataTable]$Table
    )

    $columns = Get-TbmDataColumnNames -Table $Table
    $preferred = @('email', 'emailaddress', 'mail', 'upn', 'userprincipalname', 'workemail', 'teamsemail', 'teamsaddress', 'user', 'username', 'login')
    foreach ($candidate in $preferred) {
        foreach ($column in $columns) {
            if ((ConvertTo-TbmTokenKey -Name $column) -eq $candidate) { return $column }
        }
    }

    $bestColumn = $null
    $bestCount = 0
    $sampleSize = [System.Math]::Min(50, $Table.Rows.Count)
    foreach ($column in $columns) {
        $count = 0
        for ($i = 0; $i -lt $sampleSize; $i++) {
            if (([string]$Table.Rows[$i][$column]).Contains('@')) { $count++ }
        }
        if ($count -gt $bestCount) { $bestCount = $count; $bestColumn = $column }
    }
    return $bestColumn
}

function Test-TbmEmailAddress {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter()]
        [AllowEmptyString()]
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $trimmed = $Value.Trim()
    if ($trimmed.Contains(' ') -or $trimmed.Contains(',') -or $trimmed.Contains(';')) { return $false }
    try {
        $address = [System.Net.Mail.MailAddress]::new($trimmed)
        return [string]::Equals($address.Address, $trimmed, [System.StringComparison]::OrdinalIgnoreCase) -and $address.Host.Contains('.')
    }
    catch {
        return $false
    }
}

function Update-TbmRecipientValidation {
    <#
    .SYNOPSIS
        Flags invalid and duplicate recipients (unticking them) and returns counts.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Data.DataTable]$Table,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$RecipientColumn
    )

    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $invalid = 0
    $duplicates = 0
    foreach ($row in $Table.Rows) {
        $email = ([string]$row[$RecipientColumn]).Trim()
        $status = [string]$row[$script:TbmStatusColumn]
        $isValidationStatus = [string]::IsNullOrEmpty($status) -or $status -eq 'Invalid email' -or $status -eq 'Duplicate'
        if (-not (Test-TbmEmailAddress -Value $email)) {
            $invalid++
            $row[$script:TbmIncludeColumn] = $false
            if ($isValidationStatus) { $row[$script:TbmStatusColumn] = 'Invalid email' }
            continue
        }
        if (-not $seen.Add($email)) {
            $duplicates++
            $row[$script:TbmIncludeColumn] = $false
            if ($isValidationStatus) { $row[$script:TbmStatusColumn] = 'Duplicate' }
            continue
        }
        if ($status -eq 'Invalid email' -or $status -eq 'Duplicate') {
            $row[$script:TbmStatusColumn] = ''
            $row[$script:TbmIncludeColumn] = $true
        }
    }
    return [pscustomobject]@{
        Total      = $Table.Rows.Count
        Valid      = $Table.Rows.Count - $invalid - $duplicates
        Invalid    = $invalid
        Duplicates = $duplicates
    }
}

#endregion

#region Delimited text

function Get-TbmTextFileEncoding {
    <#
    .SYNOPSIS
        Detects BOM; otherwise UTF-8 if the bytes are valid UTF-8, else the system ANSI code page (Excel "CSV" default).
    #>
    [CmdletBinding()]
    [OutputType([System.Text.Encoding])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [byte[]]$Bytes
    )

    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) { return [System.Text.UTF8Encoding]::new($true) }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) { return [System.Text.Encoding]::Unicode }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFE -and $Bytes[1] -eq 0xFF) { return [System.Text.Encoding]::BigEndianUnicode }
    try {
        $strict = [System.Text.UTF8Encoding]::new($false, $true)
        $null = $strict.GetString($Bytes)
        return [System.Text.UTF8Encoding]::new($false)
    }
    catch {
        return [System.Text.Encoding]::Default
    }
}

function Read-TbmDelimitedFile {
    <#
    .SYNOPSIS
        Reads CSV/TSV with RFC 4180 quoting (multi-line fields supported) via TextFieldParser.
        Auto-detects the delimiter (comma, semicolon or tab) from the header line. Works while Excel has the file open.
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[string[]]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    $stream = [System.IO.FileStream]::new($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $memory = [System.IO.MemoryStream]::new()
        $stream.CopyTo($memory)
        $bytes = $memory.ToArray()
    }
    finally {
        $stream.Dispose()
    }

    $encoding = Get-TbmTextFileEncoding -Bytes $bytes
    $text = $encoding.GetString($bytes)
    if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }

    $delimiter = Get-TbmDelimiter -Text $text
    $rows = [System.Collections.Generic.List[string[]]]::new()
    $reader = [System.IO.StringReader]::new($text)
    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($reader)
    try {
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters([string[]]@($delimiter))
        $parser.HasFieldsEnclosedInQuotes = $true
        $parser.TrimWhiteSpace = $false
        while (-not $parser.EndOfData) {
            $fields = $parser.ReadFields()
            if ($null -ne $fields) { $rows.Add([string[]]$fields) }
        }
    }
    finally {
        $parser.Dispose()
        $reader.Dispose()
    }
    return , $rows
}

function Get-TbmDelimiter {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Text
    )

    $end = $Text.IndexOf("`n")
    if ($end -lt 0) { $end = $Text.Length }
    $line = $Text.Substring(0, $end)
    $counts = @{ ',' = 0; ';' = 0; "`t" = 0 }
    $inQuotes = $false
    foreach ($character in $line.ToCharArray()) {
        if ($character -eq '"') { $inQuotes = -not $inQuotes; continue }
        if (-not $inQuotes -and $counts.ContainsKey([string]$character)) { $counts[[string]$character]++ }
    }
    $best = ','
    foreach ($candidate in @(';', "`t")) {
        if ($counts[$candidate] -gt $counts[$best]) { $best = $candidate }
    }
    return $best
}

#endregion

#region XLSX (native OpenXML reader)

function Get-TbmZipXml {
    [CmdletBinding()]
    [OutputType([System.Xml.XmlDocument])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.IO.Compression.ZipArchive]$Archive,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$EntryName
    )

    $entry = $Archive.GetEntry($EntryName)
    if ($null -eq $entry) {
        foreach ($candidate in $Archive.Entries) {
            if ([string]::Equals($candidate.FullName, $EntryName, [System.StringComparison]::OrdinalIgnoreCase)) { $entry = $candidate; break }
        }
    }
    if ($null -eq $entry) { return $null }

    $document = [System.Xml.XmlDocument]::new()
    $document.XmlResolver = $null
    $entryStream = $entry.Open()
    try {
        $settings = [System.Xml.XmlReaderSettings]::new()
        $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
        $xmlReader = [System.Xml.XmlReader]::Create($entryStream, $settings)
        try { $document.Load($xmlReader) } finally { $xmlReader.Dispose() }
    }
    finally {
        $entryStream.Dispose()
    }
    return $document
}

function Get-TbmNumberFormatKind {
    <#
    .SYNOPSIS
        Classifies an Excel number format as Date, Time, DateTime, Percent or Number.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [int]$FormatId,

        [Parameter()]
        [AllowEmptyString()]
        [AllowNull()]
        [string]$FormatCode
    )

    if ($FormatId -in @(14, 15, 16, 17) -or ($FormatId -ge 27 -and $FormatId -le 31) -or ($FormatId -ge 34 -and $FormatId -le 36) -or ($FormatId -ge 50 -and $FormatId -le 58)) { return 'Date' }
    if ($FormatId -eq 22) { return 'DateTime' }
    if (($FormatId -ge 18 -and $FormatId -le 21) -or ($FormatId -ge 45 -and $FormatId -le 47) -or $FormatId -in @(32, 33)) { return 'Time' }
    if ($FormatId -in @(9, 10)) { return 'Percent' }
    if ([string]::IsNullOrWhiteSpace($FormatCode) -or $FormatCode -eq 'General' -or $FormatCode -eq '@') { return 'Number' }

    # Strip literals ("text", \x, _x, *x) and [colour]/[locale] sections before looking for date/time letters.
    $code = [System.Text.RegularExpressions.Regex]::Replace($FormatCode.Split(';')[0], '"[^"]*"|\\.|_.|\*.|\[(?![hms]\])[^\]]*\]', '')
    $code = $code.ToLowerInvariant()
    $hasTime = $code.Contains('h') -or $code.Contains('s')
    $hasDate = $code.Contains('y') -or $code.Contains('d') -or ($code.Contains('m') -and -not $hasTime)
    if ($hasDate -and $hasTime) { return 'DateTime' }
    if ($hasDate) { return 'Date' }
    if ($hasTime) { return 'Time' }
    if ($code.Contains('%')) { return 'Percent' }
    return 'Number'
}

function ConvertTo-TbmCellText {
    <#
    .SYNOPSIS
        Formats a numeric cell value according to its number format kind.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [double]$Value,

        [Parameter(Mandatory)]
        [ValidateSet('Date', 'Time', 'DateTime', 'Percent', 'Number')]
        [string]$Kind,

        [Parameter()]
        [bool]$Date1904 = $false
    )

    $culture = [System.Globalization.CultureInfo]::CurrentCulture
    $serial = $Value
    if ($Date1904 -and $Kind -ne 'Number' -and $Kind -ne 'Percent') { $serial = $Value + 1462 }
    switch ($Kind) {
        'Date' {
            $date = [datetime]::FromOADate($serial)
            if ($date.TimeOfDay.Ticks -ne 0) { return $date.ToString('g', $culture) }
            return $date.ToString('d', $culture)
        }
        'DateTime' { return [datetime]::FromOADate($serial).ToString('g', $culture) }
        'Time' { return [datetime]::FromOADate($serial).ToString('t', $culture) }
        'Percent' { return ($Value * 100).ToString('0.##', $culture) + '%' }
        default { return $Value.ToString('G15', $culture) }
    }
}

function ConvertFrom-TbmCellReference {
    <#
    .SYNOPSIS
        Converts the column letters of an A1 reference to a zero-based column index.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Reference
    )

    $index = 0
    foreach ($character in $Reference.ToCharArray()) {
        if ($character -ge 'A' -and $character -le 'Z') { $index = ($index * 26) + ([int]$character - 64) }
        elseif ($character -ge 'a' -and $character -le 'z') { $index = ($index * 26) + ([int]$character - 96) }
        else { break }
    }
    return $index - 1
}

function Read-TbmXlsxNative {
    <#
    .SYNOPSIS
        Reads a worksheet from an .xlsx/.xlsm file using System.IO.Compression + XML (no Excel required).
    .OUTPUTS
        PSCustomObject: Sheets (string[]), SheetName, Rows (List[string[]]), Reader.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter()]
        [string]$SheetName
    )

    $stream = [System.IO.FileStream]::new($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Read)
        $workbook = Get-TbmZipXml -Archive $archive -EntryName 'xl/workbook.xml'
        if ($null -eq $workbook) { throw [System.IO.InvalidDataException]::new('xl/workbook.xml not found; not an Excel workbook.') }

        $ns = [System.Xml.XmlNamespaceManager]::new($workbook.NameTable)
        $ns.AddNamespace('m', $script:TbmSpreadsheetNs)

        $date1904 = $false
        $workbookPr = $workbook.SelectSingleNode('/m:workbook/m:workbookPr', $ns)
        if ($null -ne $workbookPr) {
            $flag = $workbookPr.GetAttribute('date1904')
            $date1904 = $flag -eq '1' -or $flag -eq 'true'
        }

        # Sheet name -> relationship id -> part path
        $relationships = Get-TbmZipXml -Archive $archive -EntryName 'xl/_rels/workbook.xml.rels'
        $targets = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::Ordinal)
        if ($null -ne $relationships) {
            foreach ($relationship in $relationships.DocumentElement.ChildNodes) {
                if ($relationship.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                $target = $relationship.GetAttribute('Target').Replace('\', '/')
                if ($target.StartsWith('/')) { $target = $target.TrimStart('/') } else { $target = 'xl/' + $target }
                $targets[$relationship.GetAttribute('Id')] = $target
            }
        }

        $sheetNames = [System.Collections.Generic.List[string]]::new()
        $sheetParts = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($sheet in $workbook.SelectNodes('/m:workbook/m:sheets/m:sheet', $ns)) {
            $name = $sheet.GetAttribute('name')
            $relationshipId = $sheet.GetAttribute('id', $script:TbmOfficeRelNs)
            $part = $null
            if ($targets.TryGetValue($relationshipId, [ref]$part)) {
                $sheetNames.Add($name)
                $sheetParts[$name] = $part
            }
        }
        if ($sheetNames.Count -eq 0) { throw [System.IO.InvalidDataException]::new('The workbook contains no worksheets.') }
        if ([string]::IsNullOrWhiteSpace($SheetName) -or -not $sheetParts.ContainsKey($SheetName)) { $SheetName = $sheetNames[0] }

        $sharedStrings = Read-TbmXlsxSharedStrings -Archive $archive
        $styleKinds = Read-TbmXlsxStyleKinds -Archive $archive

        $sheetXml = Get-TbmZipXml -Archive $archive -EntryName $sheetParts[$SheetName]
        if ($null -eq $sheetXml) { throw [System.IO.InvalidDataException]::new(('Worksheet part for "{0}" not found.' -f $SheetName)) }
        $sheetNs = [System.Xml.XmlNamespaceManager]::new($sheetXml.NameTable)
        $sheetNs.AddNamespace('m', $script:TbmSpreadsheetNs)

        $rows = [System.Collections.Generic.List[string[]]]::new()
        $cells = [System.Collections.Generic.Dictionary[int, string]]::new()
        foreach ($rowNode in $sheetXml.SelectNodes('/m:worksheet/m:sheetData/m:row', $sheetNs)) {
            $cells.Clear()
            $maxColumn = -1
            $position = 0
            foreach ($cell in $rowNode.ChildNodes) {
                if ($cell.LocalName -ne 'c') { continue }
                $reference = $cell.GetAttribute('r')
                $column = $position
                if (-not [string]::IsNullOrEmpty($reference)) { $column = ConvertFrom-TbmCellReference -Reference $reference }
                $position = $column + 1

                $text = Get-TbmXlsxCellText -Cell $cell -SharedStrings $sharedStrings -StyleKinds $styleKinds -Date1904 $date1904
                if ($null -ne $text) {
                    $cells[$column] = $text
                    if ($column -gt $maxColumn) { $maxColumn = $column }
                }
            }
            $values = [string[]]::new($maxColumn + 1)
            foreach ($pair in $cells.GetEnumerator()) { $values[$pair.Key] = $pair.Value }
            $rows.Add($values)
        }

        return [pscustomobject]@{
            Sheets    = $sheetNames.ToArray()
            SheetName = $SheetName
            Rows      = $rows
            Reader    = 'OpenXML'
        }
    }
    finally {
        if ($null -ne $archive) { $archive.Dispose() }
        $stream.Dispose()
    }
}

function Read-TbmXlsxSharedStrings {
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[string]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.IO.Compression.ZipArchive]$Archive
    )

    $strings = [System.Collections.Generic.List[string]]::new()
    $document = Get-TbmZipXml -Archive $Archive -EntryName 'xl/sharedStrings.xml'
    if ($null -eq $document) { return , $strings }
    foreach ($item in $document.DocumentElement.ChildNodes) {
        if ($item.LocalName -ne 'si') { continue }
        $strings.Add((Get-TbmXlsxRichText -Node $item))
    }
    return , $strings
}

function Get-TbmXlsxRichText {
    <#
    .SYNOPSIS
        Concatenates <t> runs of a shared/inline string, ignoring phonetic (<rPh>) runs.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Xml.XmlNode]$Node
    )

    $builder = [System.Text.StringBuilder]::new()
    foreach ($child in $Node.ChildNodes) {
        if ($child.LocalName -eq 't') { $null = $builder.Append($child.InnerText) }
        elseif ($child.LocalName -eq 'r') {
            foreach ($run in $child.ChildNodes) { if ($run.LocalName -eq 't') { $null = $builder.Append($run.InnerText) } }
        }
    }
    return $builder.ToString()
}

function Read-TbmXlsxStyleKinds {
    <#
    .SYNOPSIS
        Maps each cell style index (cellXfs) to a number format kind.
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[string]])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.IO.Compression.ZipArchive]$Archive
    )

    $kinds = [System.Collections.Generic.List[string]]::new()
    $document = Get-TbmZipXml -Archive $Archive -EntryName 'xl/styles.xml'
    if ($null -eq $document) { return , $kinds }
    $ns = [System.Xml.XmlNamespaceManager]::new($document.NameTable)
    $ns.AddNamespace('m', $script:TbmSpreadsheetNs)

    $customFormats = [System.Collections.Generic.Dictionary[int, string]]::new()
    foreach ($format in $document.SelectNodes('/m:styleSheet/m:numFmts/m:numFmt', $ns)) {
        $customFormats[[int]$format.GetAttribute('numFmtId')] = $format.GetAttribute('formatCode')
    }
    foreach ($xf in $document.SelectNodes('/m:styleSheet/m:cellXfs/m:xf', $ns)) {
        $formatId = 0
        $null = [int]::TryParse($xf.GetAttribute('numFmtId'), [ref]$formatId)
        $code = $null
        $null = $customFormats.TryGetValue($formatId, [ref]$code)
        $kinds.Add((Get-TbmNumberFormatKind -FormatId $formatId -FormatCode $code))
    }
    return , $kinds
}

function Get-TbmXlsxCellText {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Xml.XmlElement]$Cell,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string]]$SharedStrings,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string]]$StyleKinds,

        [Parameter()]
        [bool]$Date1904 = $false
    )

    $type = $Cell.GetAttribute('t')
    $rawValue = $null
    $inline = $null
    foreach ($child in $Cell.ChildNodes) {
        if ($child.LocalName -eq 'v') { $rawValue = $child.InnerText }
        elseif ($child.LocalName -eq 'is') { $inline = $child }
    }

    switch ($type) {
        's' {
            $index = 0
            if ([int]::TryParse($rawValue, [ref]$index) -and $index -ge 0 -and $index -lt $SharedStrings.Count) { return $SharedStrings[$index] }
            return $null
        }
        'inlineStr' { if ($null -ne $inline) { return Get-TbmXlsxRichText -Node $inline } return $null }
        'str' { return $rawValue }
        'b' { if ($rawValue -eq '1') { return 'TRUE' } return 'FALSE' }
        'e' { return $rawValue }
    }

    if ([string]::IsNullOrEmpty($rawValue)) { return $null }
    $number = 0.0
    if (-not [double]::TryParse($rawValue, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $rawValue }

    $kind = 'Number'
    $styleIndex = 0
    $styleAttribute = $Cell.GetAttribute('s')
    if (-not [string]::IsNullOrEmpty($styleAttribute) -and [int]::TryParse($styleAttribute, [ref]$styleIndex) -and $styleIndex -lt $StyleKinds.Count) {
        $kind = $StyleKinds[$styleIndex]
    }
    try {
        return ConvertTo-TbmCellText -Value $number -Kind $kind -Date1904 $Date1904
    }
    catch {
        return $number.ToString('G15', [System.Globalization.CultureInfo]::CurrentCulture)
    }
}

#endregion

#region Workbook fallbacks (ImportExcel module, Excel COM)

function Read-TbmWorkbookFallback {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter()]
        [string]$SheetName
    )

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    $importExcel = Get-Module -ListAvailable -Name 'ImportExcel'
    if ($extension -ne '.xls' -and $null -ne $importExcel -and @($importExcel).Count -gt 0) {
        try {
            return Read-TbmWorkbookImportExcel -Path $Path -SheetName $SheetName
        }
        catch {
            Write-TbmLog -Level Warn -Source 'Recipients' -Message ('ImportExcel failed ({0}); trying Excel COM.' -f $_.Exception.Message)
        }
    }
    if (Test-TbmIsWindows) {
        return Read-TbmWorkbookCom -Path $Path -SheetName $SheetName
    }
    throw [System.NotSupportedException]::new('This workbook could not be read. Save it as CSV (UTF-8) and load that instead.')
}

function Read-TbmWorkbookImportExcel {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter()]
        [string]$SheetName
    )

    Import-Module -Name 'ImportExcel' -ErrorAction Stop
    $sheetNames = [System.Collections.Generic.List[string]]::new()
    foreach ($info in (Get-ExcelSheetInfo -Path $Path)) { $sheetNames.Add([string]$info.Name) }
    if ($sheetNames.Count -eq 0) { throw [System.IO.InvalidDataException]::new('The workbook contains no worksheets.') }
    if ([string]::IsNullOrWhiteSpace($SheetName) -or -not $sheetNames.Contains($SheetName)) { $SheetName = $sheetNames[0] }

    $culture = [System.Globalization.CultureInfo]::CurrentCulture
    $rows = [System.Collections.Generic.List[string[]]]::new()
    foreach ($record in (Import-Excel -Path $Path -WorksheetName $SheetName -NoHeader)) {
        $values = [System.Collections.Generic.List[string]]::new()
        foreach ($property in $record.PSObject.Properties) {
            $value = $property.Value
            if ($null -eq $value) { $values.Add('') }
            elseif ($value -is [datetime]) {
                if ($value.TimeOfDay.Ticks -eq 0) { $values.Add($value.ToString('d', $culture)) } else { $values.Add($value.ToString('g', $culture)) }
            }
            elseif ($value -is [double]) { $values.Add($value.ToString('G15', $culture)) }
            else { $values.Add([string]$value) }
        }
        $rows.Add($values.ToArray())
    }
    return [pscustomobject]@{ Sheets = $sheetNames.ToArray(); SheetName = $SheetName; Rows = $rows; Reader = 'ImportExcel' }
}

function Read-TbmWorkbookCom {
    <#
    .SYNOPSIS
        Reads a worksheet through Excel automation (handles .xls). All COM objects are released in finally.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter()]
        [string]$SheetName
    )

    $excelType = [System.Type]::GetTypeFromProgID('Excel.Application')
    if ($null -eq $excelType) { throw [System.NotSupportedException]::new('Microsoft Excel is not installed. Save the file as CSV (UTF-8) instead.') }

    $excel = $null; $workbooks = $null; $workbook = $null; $worksheets = $null; $worksheet = $null; $usedRange = $null
    $comObjects = [System.Collections.Generic.List[object]]::new()
    try {
        $excel = [System.Activator]::CreateInstance($excelType)
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbooks = $excel.Workbooks
        $workbook = $workbooks.Open($Path, 0, $true)
        $worksheets = $workbook.Worksheets

        $sheetNames = [System.Collections.Generic.List[string]]::new()
        for ($i = 1; $i -le $worksheets.Count; $i++) {
            $sheet = $worksheets.Item($i)
            $sheetNames.Add([string]$sheet.Name)
            $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)
        }
        if ([string]::IsNullOrWhiteSpace($SheetName) -or -not $sheetNames.Contains($SheetName)) { $SheetName = $sheetNames[0] }

        $worksheet = $worksheets.Item($SheetName)
        $usedRange = $worksheet.UsedRange
        $rowCount = [int]$usedRange.Rows.Count
        $columnCount = [int]$usedRange.Columns.Count
        $data = $usedRange.Value2

        # Number format per column, taken from the first data row (row 2 of the used range).
        $kinds = [string[]]::new($columnCount)
        for ($c = 1; $c -le $columnCount; $c++) {
            $kinds[$c - 1] = 'Number'
            if ($rowCount -ge 2) {
                $cell = $usedRange.Cells.Item(2, $c)
                $kinds[$c - 1] = Get-TbmNumberFormatKind -FormatId -1 -FormatCode ([string]$cell.NumberFormat)
                $comObjects.Add($cell)
            }
        }

        $rows = [System.Collections.Generic.List[string[]]]::new()
        for ($r = 1; $r -le $rowCount; $r++) {
            $values = [string[]]::new($columnCount)
            for ($c = 1; $c -le $columnCount; $c++) {
                $value = $null
                if ($rowCount -eq 1 -and $columnCount -eq 1) { $value = $data } else { $value = $data[$r, $c] }
                if ($null -eq $value) { $values[$c - 1] = ''; continue }
                if ($value -is [double]) {
                    $kind = 'Number'
                    if ($r -gt 1) { $kind = $kinds[$c - 1] }
                    $values[$c - 1] = ConvertTo-TbmCellText -Value $value -Kind $kind
                }
                elseif ($value -is [bool]) { $values[$c - 1] = ([string]$value).ToUpperInvariant() }
                else { $values[$c - 1] = [string]$value }
            }
            $rows.Add($values)
        }
        return [pscustomobject]@{ Sheets = $sheetNames.ToArray(); SheetName = $SheetName; Rows = $rows; Reader = 'Excel COM' }
    }
    finally {
        if ($null -ne $workbook) { try { $workbook.Close($false) } catch { Write-TbmLog -Level Debug -Source 'Recipients' -Message 'Workbook close failed.' } }
        if ($null -ne $excel) { try { $excel.Quit() } catch { Write-TbmLog -Level Debug -Source 'Recipients' -Message 'Excel quit failed.' } }
        foreach ($comObject in $comObjects) { $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) }
        foreach ($comObject in @($usedRange, $worksheet, $worksheets, $workbook, $workbooks, $excel)) {
            if ($null -ne $comObject) { $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) }
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

#endregion
