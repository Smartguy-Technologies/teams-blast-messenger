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

    return [TeamsBlastMessenger.TbmText]::IsEmailAddress($Value)
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
    $includeColumn = $Table.Columns[$script:TbmIncludeColumn]
    $statusColumn = $Table.Columns[$script:TbmStatusColumn]
    $emailColumn = $Table.Columns[$RecipientColumn]
    # Suspend view notifications while rows change (one refresh at the end), and only write cells
    # whose value actually changes, so a bound grid is not redrawn row by row.
    $Table.BeginLoadData()
    try {
        foreach ($row in $Table.Rows) {
            $email = ([string]$row[$emailColumn]).Trim()
            $status = [string]$row[$statusColumn]
            $isValidationStatus = [string]::IsNullOrEmpty($status) -or $status -eq 'Invalid email' -or $status -eq 'Duplicate'
            $newStatus = $null
            if (-not [TeamsBlastMessenger.TbmText]::IsEmailAddress($email)) {
                $invalid++
                $newStatus = 'Invalid email'
            }
            elseif (-not $seen.Add($email)) {
                $duplicates++
                $newStatus = 'Duplicate'
            }

            if ($null -ne $newStatus) {
                if ([bool]$row[$includeColumn]) { $row[$includeColumn] = $false }
                if ($isValidationStatus -and $status -ne $newStatus) { $row[$statusColumn] = $newStatus }
                continue
            }
            if ($status -eq 'Invalid email' -or $status -eq 'Duplicate') {
                $row[$statusColumn] = ''
                if (-not [bool]$row[$includeColumn]) { $row[$includeColumn] = $true }
            }
        }
    }
    finally {
        $Table.EndLoadData()
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

function Get-TbmZipEntry {
    <#
    .SYNOPSIS
        Finds a package part by name (exact first, then case-insensitive).
    #>
    [CmdletBinding()]
    [OutputType([System.IO.Compression.ZipArchiveEntry])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.IO.Compression.ZipArchive]$Archive,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$EntryName
    )

    $entry = $Archive.GetEntry($EntryName)
    if ($null -ne $entry) { return $entry }
    foreach ($candidate in $Archive.Entries) {
        if ([string]::Equals($candidate.FullName, $EntryName, [System.StringComparison]::OrdinalIgnoreCase)) { return $candidate }
    }
    return $null
}

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

    $entry = Get-TbmZipEntry -Archive $Archive -EntryName $EntryName
    if ($null -eq $entry) { return $null }

    $document = [System.Xml.XmlDocument]::new()
    $document.XmlResolver = $null
    $entryStream = $entry.Open()
    try {
        $xmlReader = [System.Xml.XmlReader]::Create($entryStream, [TeamsBlastMessenger.TbmXlsx]::CreateReaderSettings())
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

    return [TeamsBlastMessenger.TbmXlsx]::GetNumberFormatKind($FormatId, $FormatCode)
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

    return [TeamsBlastMessenger.TbmXlsx]::FormatNumber($Value, $Kind, $Date1904)
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

        $sheetEntry = Get-TbmZipEntry -Archive $archive -EntryName $sheetParts[$SheetName]
        if ($null -eq $sheetEntry) { throw [System.IO.InvalidDataException]::new(('Worksheet part for "{0}" not found.' -f $SheetName)) }
        $sheetStream = $sheetEntry.Open()
        try { $rows = [TeamsBlastMessenger.TbmXlsx]::ReadSheet($sheetStream, $sharedStrings, $styleKinds, $date1904) }
        finally { $sheetStream.Dispose() }

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

    $entry = Get-TbmZipEntry -Archive $Archive -EntryName 'xl/sharedStrings.xml'
    if ($null -eq $entry) { return , [System.Collections.Generic.List[string]]::new() }
    $stream = $entry.Open()
    try { return , [TeamsBlastMessenger.TbmXlsx]::ReadSharedStrings($stream) }
    finally { $stream.Dispose() }
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

    $entry = Get-TbmZipEntry -Archive $Archive -EntryName 'xl/styles.xml'
    if ($null -eq $entry) { return , [System.Collections.Generic.List[string]]::new() }
    $stream = $entry.Open()
    try { return , [TeamsBlastMessenger.TbmXlsx]::ReadStyleKinds($stream) }
    finally { $stream.Dispose() }
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
    $importExcel = Get-TbmAvailableModule -Name 'ImportExcel'
    if ($extension -ne '.xls' -and $importExcel.Length -gt 0) {
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
