$ErrorActionPreference = 'Stop'

$outputDir = 'C:\Users\User\Documents\G3\prepared_dataset_ml'
$csvExportDir = Join-Path $outputDir 'csv_exports'
$sourceWorkbook = Join-Path $outputDir 'ml_model_datasets.xlsx'
$tempRoot = Join-Path $outputDir 'ml_model_datasets_with_qa_build'
$outputWorkbook = Join-Path $outputDir 'ml_model_datasets_with_qa.xlsx'

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
    param([string]$CsvPath)

    $rows = Import-Csv -LiteralPath $CsvPath
    $headers = if ($rows.Count -gt 0) { @($rows[0].PSObject.Properties.Name) } else { @() }
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

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($sourceWorkbook, $tempRoot)

$sheetSpecs = @(
    @{ Name = 'qa_summary'; CsvPath = (Join-Path $csvExportDir 'qa_summary.csv'); SheetId = 3; RelationshipId = 4 },
    @{ Name = 'data_dictionary'; CsvPath = (Join-Path $csvExportDir 'data_dictionary.csv'); SheetId = 4; RelationshipId = 5 },
    @{ Name = 'model_feature_notes'; CsvPath = (Join-Path $csvExportDir 'model_feature_notes.csv'); SheetId = 5; RelationshipId = 6 }
)

foreach ($sheetSpec in $sheetSpecs) {
    $worksheetXml = Convert-CsvToWorksheetXml -CsvPath $sheetSpec.CsvPath
    $worksheetPath = Join-Path $tempRoot ('xl\worksheets\sheet' + $sheetSpec.SheetId + '.xml')
    Write-Utf8File -Path $worksheetPath -Content $worksheetXml
}

$workbookPath = Join-Path $tempRoot 'xl\workbook.xml'
$workbookContent = Get-Content -LiteralPath $workbookPath -Raw
foreach ($sheetSpec in $sheetSpecs) {
    $sheetLine = '<sheet name="' + (Escape-Xml $sheetSpec.Name) + '" sheetId="' + $sheetSpec.SheetId + '" r:id="rId' + $sheetSpec.RelationshipId + '"/>'
    $workbookContent = $workbookContent -replace '</sheets>', ($sheetLine + "`r`n</sheets>")
}
Write-Utf8File -Path $workbookPath -Content $workbookContent

$workbookRelsPath = Join-Path $tempRoot 'xl\_rels\workbook.xml.rels'
$workbookRelsContent = Get-Content -LiteralPath $workbookRelsPath -Raw
foreach ($sheetSpec in $sheetSpecs) {
    $relLine = '<Relationship Id="rId' + $sheetSpec.RelationshipId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + $sheetSpec.SheetId + '.xml"/>'
    $workbookRelsContent = $workbookRelsContent -replace '</Relationships>', ($relLine + "`r`n</Relationships>")
}
Write-Utf8File -Path $workbookRelsPath -Content $workbookRelsContent

$contentTypesPath = Join-Path $tempRoot '[Content_Types].xml'
$contentTypesContent = Get-Content -LiteralPath $contentTypesPath -Raw
foreach ($sheetSpec in $sheetSpecs) {
    $overrideLine = '<Override PartName="/xl/worksheets/sheet' + $sheetSpec.SheetId + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    $contentTypesContent = $contentTypesContent -replace '</Types>', ($overrideLine + "`r`n</Types>")
}
Write-Utf8File -Path $contentTypesPath -Content $contentTypesContent

Remove-IfExists -Path $outputWorkbook
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempRoot, $outputWorkbook)

Write-Output ('OUTPUT_XLSX=' + $outputWorkbook)
