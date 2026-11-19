#Requires -Version 5.1
<#
    Micro-benchmarks behind the project's coding rules, plus an XLSX import timing.
    Not part of the Pester suite. Run on the PowerShell you deploy to (Windows PowerShell 5.1):

        powershell.exe -NoProfile -File .\Tests\Benchmarks.ps1
        powershell.exe -NoProfile -File .\Tests\Benchmarks.ps1 -Rows 20000

    This file deliberately contains the slow idioms it measures ([void], | Out-Null, +=, Where-Object),
    so the coding-standards test skips it.
#>
[CmdletBinding()]
param(
    # Size of the generated workbook for the XLSX import timing.
    [ValidateRange(100, 200000)]
    [int]$Rows = 10000,

    # Iterations for the idiom measurements.
    [ValidateRange(1000, 10000000)]
    [int]$Iterations = 200000
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Measure-TbmIdiom {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][int]$Count,
        [Parameter(Mandatory)][scriptblock]$Script
    )
    $null = & $Script
    $best = [double]::MaxValue
    for ($round = 0; $round -lt 3; $round++) {
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $null = & $Script
        $watch.Stop()
        if ($watch.Elapsed.TotalMilliseconds -lt $best) { $best = $watch.Elapsed.TotalMilliseconds }
    }
    $perItem = $best * 1000000 / $Count
    '{0,-38} {1,10:N1} ms  {2,10:N0} ns/item' -f $Name, $best, $perItem
}

'PowerShell {0} ({1})' -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition
''
$n = $Iterations
$small = [System.Math]::Max(1000, [int]($Iterations / 10))
$list = [System.Collections.Generic.List[int]]::new()

'-- Discarding a return value'
Measure-TbmIdiom -Name '[void]$list.Add(i)' -Count $n -Script { $list.Clear(); for ($i = 0; $i -lt $n; $i++) { [void]$list.Add($i) } }
Measure-TbmIdiom -Name '$null = $list.Add(i)' -Count $n -Script { $list.Clear(); for ($i = 0; $i -lt $n; $i++) { $null = $list.Add($i) } }
Measure-TbmIdiom -Name '$list.Add(i) > $null' -Count $n -Script { $list.Clear(); for ($i = 0; $i -lt $n; $i++) { $list.Add($i) > $null } }
Measure-TbmIdiom -Name '$list.Add(i) | Out-Null' -Count $small -Script { $list.Clear(); for ($i = 0; $i -lt $small; $i++) { $list.Add($i) | Out-Null } }

''
'-- Filtering'
$data = [int[]](1..$n)
Measure-TbmIdiom -Name '| Where-Object { }' -Count $n -Script { $data | Where-Object { $_ % 2 -eq 0 } }
Measure-TbmIdiom -Name '.Where({ })' -Count $n -Script { $data.Where({ $_ % 2 -eq 0 }) }
Measure-TbmIdiom -Name 'foreach + if + List.Add' -Count $n -Script { $kept = [System.Collections.Generic.List[int]]::new(); foreach ($item in $data) { if ($item % 2 -eq 0) { $kept.Add($item) } }; , $kept }

''
'-- Accumulating'
Measure-TbmIdiom -Name '$array += item' -Count $small -Script { $array = @(); for ($i = 0; $i -lt $small; $i++) { $array += $i }; , $array }
Measure-TbmIdiom -Name 'List[int].Add' -Count $small -Script { $items = [System.Collections.Generic.List[int]]::new(); for ($i = 0; $i -lt $small; $i++) { $items.Add($i) }; , $items }
Measure-TbmIdiom -Name '$string += text' -Count $small -Script { $text = ''; for ($i = 0; $i -lt $small; $i++) { $text += 'x' + $i + ',' }; $text }
Measure-TbmIdiom -Name 'StringBuilder.Append' -Count $small -Script { $builder = [System.Text.StringBuilder]::new(); for ($i = 0; $i -lt $small; $i++) { $null = $builder.Append('x').Append($i).Append(',') }; $builder.ToString() }

''
'-- Calling per item'
function Get-BenchAdvanced { [CmdletBinding()] param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Value) return $Value.Length }
function Get-BenchSimple { param($Value) $Value.Length }
Measure-TbmIdiom -Name 'advanced function' -Count $small -Script { for ($i = 0; $i -lt $small; $i++) { $null = Get-BenchAdvanced -Value 'AB12' } }
Measure-TbmIdiom -Name 'simple function' -Count $small -Script { for ($i = 0; $i -lt $small; $i++) { $null = Get-BenchSimple 'AB12' } }
Measure-TbmIdiom -Name '.NET static method' -Count $small -Script { for ($i = 0; $i -lt $small; $i++) { $null = [System.String]::IsNullOrEmpty('AB12') } }

''
'-- XLSX import ({0:N0} rows x 10 columns)' -f $Rows
$projectRoot = [System.IO.Path]::GetDirectoryName($PSScriptRoot)
$workRoot = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), 'tbm-bench-' + [guid]::NewGuid().ToString('N'))
$null = [System.IO.Directory]::CreateDirectory($workRoot)
[System.Environment]::SetEnvironmentVariable('TBM_APPDATA', [System.IO.Path]::Combine($workRoot, 'appdata'))
try {
    Add-Type -AssemblyName System.IO.Compression
    $xlsxPath = [System.IO.Path]::Combine($workRoot, 'bench.xlsx')
    $ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $sheet = [System.Text.StringBuilder]::new()
    $null = $sheet.Append('<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="').Append($ns).Append('"><sheetData>')
    $null = $sheet.Append('<row r="1"><c r="A1" t="inlineStr"><is><t>Email</t></is></c><c r="B1" t="inlineStr"><is><t>First Name</t></is></c><c r="C1" t="inlineStr"><is><t>Joined</t></is></c><c r="D1" t="inlineStr"><is><t>Score</t></is></c><c r="E1" t="inlineStr"><is><t>Pct</t></is></c><c r="F1" t="inlineStr"><is><t>Active</t></is></c><c r="G1" t="inlineStr"><is><t>Team</t></is></c><c r="H1" t="inlineStr"><is><t>City</t></is></c><c r="I1" t="inlineStr"><is><t>Code</t></is></c><c r="J1" t="inlineStr"><is><t>Note</t></is></c></row>')
    for ($r = 2; $r -le $Rows + 1; $r++) {
        $null = $sheet.AppendFormat([System.Globalization.CultureInfo]::InvariantCulture,
            '<row r="{0}"><c r="A{0}" t="inlineStr"><is><t>user{0}@contoso.com</t></is></c><c r="B{0}" t="inlineStr"><is><t>First{1}</t></is></c><c r="C{0}" s="1"><v>{2}</v></c><c r="D{0}"><v>{3}</v></c><c r="E{0}" s="2"><v>0.{4}</v></c><c r="F{0}" t="b"><v>{5}</v></c><c r="G{0}" t="inlineStr"><is><t>Team {6}</t></is></c><c r="H{0}" t="inlineStr"><is><t>City {7}</t></is></c><c r="I{0}"><v>{0}</v></c><c r="J{0}" t="inlineStr"><is><t>Note for row {0}</t></is></c></row>',
            $r, ($r % 97), (40000 + ($r % 3000)), (($r * 7) % 1000), ($r % 100), ($r % 2), ($r % 12), ($r % 40))
    }
    $null = $sheet.Append('</sheetData></worksheet>')
    $parts = [ordered]@{
        '[Content_Types].xml'        = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>'
        'xl/workbook.xml'            = '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="' + $ns + '" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="People" sheetId="1" r:id="rId1"/></sheets></workbook>'
        'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
        'xl/styles.xml'              = '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="' + $ns + '"><cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="10"/></cellXfs></styleSheet>'
        'xl/worksheets/sheet1.xml'   = $sheet.ToString()
    }
    $file = [System.IO.FileStream]::new($xlsxPath, [System.IO.FileMode]::Create)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($file, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $parts.Keys) {
                $writer = [System.IO.StreamWriter]::new($zip.CreateEntry($name).Open(), [System.Text.UTF8Encoding]::new($false))
                try { $writer.Write($parts[$name]) } finally { $writer.Dispose() }
            }
        }
        finally { $zip.Dispose() }
    }
    finally { $file.Dispose() }

    Import-Module ([System.IO.Path]::Combine($projectRoot, 'TeamsBlastMessenger.psd1')) -Force
    InModuleScope TeamsBlastMessenger -Parameters @{ Path = $xlsxPath } {
        param($Path)
        $null = Initialize-TbmEnvironment
        $null = Import-TbmRecipientFile -Path $Path
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $result = Import-TbmRecipientFile -Path $Path
        $watch.Stop()
        '{0,-38} {1,10:N1} ms  ({2:N0} rows, reader {3})' -f 'Import-TbmRecipientFile', $watch.Elapsed.TotalMilliseconds, $result.Table.Rows.Count, $result.Reader
        $watch.Restart()
        $summary = Update-TbmRecipientValidation -Table $result.Table -RecipientColumn 'Email'
        $watch.Stop()
        '{0,-38} {1,10:N1} ms  ({2:N0} valid)' -f 'Update-TbmRecipientValidation', $watch.Elapsed.TotalMilliseconds, $summary.Valid
    }
}
finally {
    Remove-Module TeamsBlastMessenger -ErrorAction SilentlyContinue
    [System.Environment]::SetEnvironmentVariable('TBM_APPDATA', $null)
    if ([System.IO.Directory]::Exists($workRoot)) { [System.IO.Directory]::Delete($workRoot, $true) }
}
