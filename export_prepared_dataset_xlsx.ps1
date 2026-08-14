$ErrorActionPreference = 'Stop'

$preparedDir = 'C:\Users\User\Documents\G3\prepared_dataset'
$outputPath = Join-Path $preparedDir 'prepared_dataset.xlsx'
$tempRoot = Join-Path $preparedDir 'xlsx_build'

$sheetSpecs = @(
    @{ Name = 'dataset_summary'; File = 'dataset_summary.csv' },
    @{ Name = 'indicator_reference'; File = 'indicator_reference.csv' },
    @{ Name = 'ml_entities'; File = 'ml_entities.csv' },
    @{ Name = 'unknown_codes'; File = 'unknown_improvement_codes.csv' },
    @{ Name = 'ml_dataset_wide'; File = 'ml_dataset_wide.csv' },
    @{ Name = 'ml_dataset_long'; File = 'ml_dataset_long.csv' }
)

function Remove-IfExists {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force
    }
}

function New-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Escape-Xml {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) {
        return ''
    }

    return [System.Security.SecurityElement]::Escape([string]$Value)
}

function Get-ExcelColumnName {
    param([int]$Index)

    $name = ''
    $current = $Index
    while ($current -gt 0) {
        $remainder = ($current - 1) % 26
        $name = ([char][int](65 + $remainder)) + $name
        $current = [math]::Floor(($current - 1) / 26)
    }

    return $name
}

function Get-CellXml {
    param(
        [string]$CellRef,
        [AllowNull()][object]$Value,
        [bool]$IsHeader
    )

    $text = if ($null -eq $Value) { '' } else { [string]$Value }
    if (-not $IsHeader -and $text -match '^-?\d+(?:\.\d+)?$') {
        return '<c r="' + $CellRef + '"><v>' + $text + '</v></c>'
    }

    return '<c r="' + $CellRef + '" t="inlineStr"><is><t xml:space="preserve">' + (Escape-Xml $text) + '</t></is></c>'
}

function Convert-CsvToWorksheetXml {
    param(
        [Parameter(Mandatory = $true)][string]$CsvPath,
        [Parameter(Mandatory = $true)][string]$SheetName
    )

    $rows = Import-Csv -LiteralPath $CsvPath
    $headers = @()
    if ($rows.Count -gt 0) {
        $headers = @($rows[0].PSObject.Properties.Name)
    }
    else {
        $firstLine = Get-Content -LiteralPath $CsvPath -TotalCount 1
        if ($firstLine) {
            $headers = @($firstLine.Split(',') | ForEach-Object { $_.Trim('"') })
        }
    }

    $dimensionEnd = if ($headers.Count -gt 0) {
        (Get-ExcelColumnName $headers.Count) + ($rows.Count + 1)
    }
    else {
        'A1'
    }

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$builder.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$builder.AppendLine('<dimension ref="A1:' + $dimensionEnd + '"/>')
    [void]$builder.AppendLine('<sheetViews><sheetView workbookViewId="0"/></sheetViews>')
    [void]$builder.AppendLine('<sheetFormatPr defaultRowHeight="15"/>')
    [void]$builder.AppendLine('<sheetData>')

    if ($headers.Count -gt 0) {
        [void]$builder.Append('<row r="1">')
        for ($col = 0; $col -lt $headers.Count; $col++) {
            $cellRef = (Get-ExcelColumnName ($col + 1)) + '1'
            [void]$builder.Append((Get-CellXml -CellRef $cellRef -Value $headers[$col] -IsHeader $true))
        }
        [void]$builder.AppendLine('</row>')
    }

    for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex++) {
        $excelRow = $rowIndex + 2
        [void]$builder.Append('<row r="' + $excelRow + '">')
        for ($col = 0; $col -lt $headers.Count; $col++) {
            $header = $headers[$col]
            $cellRef = (Get-ExcelColumnName ($col + 1)) + $excelRow
            [void]$builder.Append((Get-CellXml -CellRef $cellRef -Value $rows[$rowIndex].$header -IsHeader $false))
        }
        [void]$builder.AppendLine('</row>')
    }

    [void]$builder.AppendLine('</sheetData>')
    [void]$builder.AppendLine('</worksheet>')
    return $builder.ToString()
}

function Write-Utf8File {
    param(
        [string]$Path,
        [string]$Content
    )

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

Remove-IfExists -Path $tempRoot
New-Directory -Path $tempRoot
New-Directory -Path (Join-Path $tempRoot '_rels')
New-Directory -Path (Join-Path $tempRoot 'xl')
New-Directory -Path (Join-Path $tempRoot 'xl\_rels')
New-Directory -Path (Join-Path $tempRoot 'xl\worksheets')

$contentTypes = New-Object System.Text.StringBuilder
[void]$contentTypes.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
[void]$contentTypes.AppendLine('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
[void]$contentTypes.AppendLine('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
[void]$contentTypes.AppendLine('<Default Extension="xml" ContentType="application/xml"/>')
[void]$contentTypes.AppendLine('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
[void]$contentTypes.AppendLine('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')

$workbook = New-Object System.Text.StringBuilder
[void]$workbook.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
[void]$workbook.AppendLine('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
[void]$workbook.AppendLine('<sheets>')

$workbookRels = New-Object System.Text.StringBuilder
[void]$workbookRels.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
[void]$workbookRels.AppendLine('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
[void]$workbookRels.AppendLine('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')

$relationshipId = 2
$sheetId = 1
foreach ($spec in $sheetSpecs) {
    $csvPath = Join-Path $preparedDir $spec.File
    if (-not (Test-Path -LiteralPath $csvPath)) {
        throw "Missing CSV file: $csvPath"
    }

    $worksheetXml = Convert-CsvToWorksheetXml -CsvPath $csvPath -SheetName $spec.Name
    $worksheetPath = Join-Path $tempRoot ('xl\worksheets\sheet' + $sheetId + '.xml')
    Write-Utf8File -Path $worksheetPath -Content $worksheetXml

    [void]$contentTypes.AppendLine('<Override PartName="/xl/worksheets/sheet' + $sheetId + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    [void]$workbook.AppendLine('<sheet name="' + (Escape-Xml $spec.Name) + '" sheetId="' + $sheetId + '" r:id="rId' + $relationshipId + '"/>')
    [void]$workbookRels.AppendLine('<Relationship Id="rId' + $relationshipId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + $sheetId + '.xml"/>')

    $relationshipId++
    $sheetId++
}

[void]$workbook.AppendLine('</sheets>')
[void]$workbook.AppendLine('</workbook>')

[void]$workbookRels.AppendLine('</Relationships>')
[void]$contentTypes.AppendLine('</Types>')

$rootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@

$styles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
'@

Write-Utf8File -Path (Join-Path $tempRoot '[Content_Types].xml') -Content $contentTypes.ToString()
Write-Utf8File -Path (Join-Path $tempRoot '_rels\.rels') -Content $rootRels
Write-Utf8File -Path (Join-Path $tempRoot 'xl\workbook.xml') -Content $workbook.ToString()
Write-Utf8File -Path (Join-Path $tempRoot 'xl\_rels\workbook.xml.rels') -Content $workbookRels.ToString()
Write-Utf8File -Path (Join-Path $tempRoot 'xl\styles.xml') -Content $styles

Remove-IfExists -Path $outputPath
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempRoot, $outputPath)

Write-Output ('XLSX_PATH=' + $outputPath)
