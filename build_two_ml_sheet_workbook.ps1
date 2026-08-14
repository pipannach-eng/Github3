$ErrorActionPreference = 'Stop'

$outputDir = 'C:\Users\User\Documents\G3\prepared_dataset_ml'
$csvExportDir = Join-Path $outputDir 'csv_exports'
$tempXlsxRoot = Join-Path $outputDir 'model_sheets_xlsx_build'
$outputPath = Join-Path $outputDir 'ml_model_datasets.xlsx'

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

function Export-RowsToCsv {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Rows,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ($Rows.Count -eq 0) {
        $Rows = @([pscustomobject]@{ status = 'no_rows' })
    }

    $Rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function Normalize-Text {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) {
        return ''
    }

    return ([string]$Value).Trim()
}

function Add-DictionaryRows {
    param(
        [System.Collections.Generic.List[object]]$DictionaryRows,
        [string]$SheetName,
        [string[]]$Columns,
        [string]$Role,
        [string]$DescriptionPrefix
    )

    foreach ($column in $Columns) {
        $DictionaryRows.Add([pscustomobject]@{
            sheet_name   = $SheetName
            column_name  = $column
            role         = $Role
            data_type    = 'string_or_numeric'
            description  = $DescriptionPrefix + $column
        })
    }
}

$datasetCleanPath = Join-Path $csvExportDir 'dataset_clean.csv'
$benchmarkPath = Join-Path $csvExportDir 'model_dataset_benchmark.csv'
$leakageSafePath = Join-Path $csvExportDir 'model_dataset_leakage_safe.csv'
$rowCountsPath = Join-Path $csvExportDir 'row_counts_by_year.csv'
$codeMappingPath = Join-Path $csvExportDir 'code_mapping.csv'
$indicatorReferencePath = Join-Path $csvExportDir 'indicator_reference.csv'

$datasetCleanRows = @(Import-Csv -LiteralPath $datasetCleanPath)
$benchmarkRows = @(Import-Csv -LiteralPath $benchmarkPath)
$leakageSafeRows = @(Import-Csv -LiteralPath $leakageSafePath)
$rowCountRows = @(Import-Csv -LiteralPath $rowCountsPath)
$codeMappingRows = @(Import-Csv -LiteralPath $codeMappingPath)
$indicatorReferenceRows = @(Import-Csv -LiteralPath $indicatorReferencePath)
$indicatorColumns = @($indicatorReferenceRows | ForEach-Object { $_.feature_column })
$failLabelObserved = Normalize-Text (($datasetCleanRows | Where-Object { $_.Quality_level -eq 'D' } | Select-Object -First 1).target_label)
$passLabelObserved = Normalize-Text (($datasetCleanRows | Where-Object { $_.Quality_level -ne 'D' } | Select-Object -First 1).target_label)

$qualityDistribution = @($datasetCleanRows | Group-Object Quality_level | Sort-Object Name)
$targetDistribution = @($datasetCleanRows | Group-Object target_label | Sort-Object Name)
$duplicateGroupCount = @($datasetCleanRows | Group-Object record_key | Where-Object { $_.Count -gt 1 }).Count
$duplicateRowCount = @($datasetCleanRows | Group-Object record_key | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
if ($null -eq $duplicateRowCount) {
    $duplicateRowCount = 0
}

$indicatorNonBinaryRowCount = 0
$indicatorBinaryFailColumns = New-Object System.Collections.Generic.List[string]
foreach ($indicatorColumn in $indicatorColumns) {
    $hasNonBinary = $false
    foreach ($row in $datasetCleanRows) {
        $value = Normalize-Text $row.$indicatorColumn
        if ($value -ne '0' -and $value -ne '1') {
            $hasNonBinary = $true
            $indicatorNonBinaryRowCount++
        }
    }
    if ($hasNonBinary) {
        $indicatorBinaryFailColumns.Add($indicatorColumn)
    }
}

$improveScoreMismatchCount = 0
$targetMismatchCount = 0
foreach ($row in $datasetCleanRows) {
    $indicatorSum = 0
    foreach ($indicatorColumn in $indicatorColumns) {
        $indicatorSum += [int](Normalize-Text $row.$indicatorColumn)
    }

    if ($indicatorSum -ne [int](Normalize-Text $row.ImproveScore)) {
        $improveScoreMismatchCount++
    }

    $expectedTarget = if ((Normalize-Text $row.Quality_level) -eq 'D') { $failLabelObserved } else { $passLabelObserved }
    if ($expectedTarget -ne (Normalize-Text $row.target_label)) {
        $targetMismatchCount++
    }
}

$importantMissingColumns = @(
    'assessment_year_be',
    'assessment_year_ce',
    'record_key',
    'center_code_anon',
    'ministry',
    'sub_agency_1',
    'sub_agency_2',
    'sub_agency_3',
    'opening_date_iso',
    'Quality_level',
    'target_label'
)

$missingSummaryImportant = foreach ($columnName in $importantMissingColumns) {
    $missingCount = 0
    foreach ($row in $datasetCleanRows) {
        if ([string]::IsNullOrWhiteSpace((Normalize-Text $row.$columnName))) {
            $missingCount++
        }
    }

    [pscustomobject]@{
        metric_category = 'missing_values'
        metric_name     = $columnName
        metric_value    = $missingCount
        detail          = 'Important field missing count in dataset_clean.'
    }
}

$qaSummaryRows = New-Object System.Collections.Generic.List[object]
$qaSummaryRows.Add([pscustomobject]@{
    metric_category = 'dataset_overview'
    metric_name     = 'source_file_count'
    metric_value    = (@($rowCountRows | Where-Object { $_.assessment_year_be -ne 'TOTAL' })).Count
    detail          = 'Number of annual source files included.'
})

foreach ($rowCount in ($rowCountRows | Where-Object { $_.assessment_year_be -ne 'TOTAL' })) {
    $qaSummaryRows.Add([pscustomobject]@{
        metric_category = 'rows_by_year'
        metric_name     = 'assessment_year_be_' + $rowCount.assessment_year_be
        metric_value    = [int]$rowCount.row_count
        detail          = $rowCount.source_file
    })
}

$qaSummaryRows.Add([pscustomobject]@{
    metric_category = 'dataset_overview'
    metric_name     = 'total_rows'
    metric_value    = $datasetCleanRows.Count
    detail          = 'Total rows across all years.'
})
$qaSummaryRows.Add([pscustomobject]@{
    metric_category = 'dataset_overview'
    metric_name     = 'distinct_education_centers'
    metric_value    = $codeMappingRows.Count
    detail          = 'Distinct center_code_anon count.'
})
$qaSummaryRows.Add([pscustomobject]@{
    metric_category = 'dataset_overview'
    metric_name     = 'indicator_column_count'
    metric_value    = $indicatorColumns.Count
    detail          = 'Count of indicator_* columns.'
})

foreach ($group in $qualityDistribution) {
    $qaSummaryRows.Add([pscustomobject]@{
        metric_category = 'quality_level_distribution'
        metric_name     = $group.Name
        metric_value    = $group.Count
        detail          = 'Distribution of Quality_level.'
    })
}

foreach ($group in $targetDistribution) {
    $qaSummaryRows.Add([pscustomobject]@{
        metric_category = 'target_label_distribution'
        metric_name     = $group.Name
        metric_value    = $group.Count
        detail          = 'Distribution of target_label.'
    })
}

foreach ($row in $missingSummaryImportant) {
    $qaSummaryRows.Add($row)
}

$qaSummaryRows.Add([pscustomobject]@{
    metric_category = 'duplicate_check'
    metric_name     = 'duplicate_record_key_groups'
    metric_value    = $duplicateGroupCount
    detail          = 'Count of duplicated record_key groups.'
})
$qaSummaryRows.Add([pscustomobject]@{
    metric_category = 'duplicate_check'
    metric_name     = 'duplicate_record_key_rows'
    metric_value    = $duplicateRowCount
    detail          = 'Rows involved in duplicated record_key groups.'
})
$qaSummaryRows.Add([pscustomobject]@{
    metric_category = 'validation'
    metric_name     = 'indicator_binary_only'
    metric_value    = if ($indicatorBinaryFailColumns.Count -eq 0) { 'pass' } else { 'fail' }
    detail          = if ($indicatorBinaryFailColumns.Count -eq 0) { 'All indicator_* values are 0/1.' } else { 'Non-binary indicator columns: ' + ($indicatorBinaryFailColumns -join ',') }
})
$qaSummaryRows.Add([pscustomobject]@{
    metric_category = 'validation'
    metric_name     = 'indicator_non_binary_row_count'
    metric_value    = $indicatorNonBinaryRowCount
    detail          = 'Total non-binary indicator cell count.'
})
$qaSummaryRows.Add([pscustomobject]@{
    metric_category = 'validation'
    metric_name     = 'ImproveScore_matches_indicator_sum'
    metric_value    = if ($improveScoreMismatchCount -eq 0) { 'pass' } else { 'fail' }
    detail          = 'Mismatch row count=' + $improveScoreMismatchCount
})
$qaSummaryRows.Add([pscustomobject]@{
    metric_category = 'validation'
    metric_name     = 'target_label_matches_quality_level'
    metric_value    = if ($targetMismatchCount -eq 0) { 'pass' } else { 'fail' }
    detail          = 'Mismatch row count=' + $targetMismatchCount
})

$dataDictionaryRows = New-Object System.Collections.Generic.List[object]
$benchmarkFeatureColumns = @('ministry', 'affiliation_1', 'affiliation_2', 'affiliation_3') + $indicatorColumns
$benchmarkTargetColumns = @('target_label', 'Quality_level')
$leakageSafeFeatureColumns = @(
    'assessment_year_be',
    'assessment_year_ce',
    'ministry',
    'affiliation_1',
    'affiliation_2',
    'affiliation_3',
    'center_age_years'
) + @($indicatorColumns | ForEach-Object { 'prev_' + $_ }) + @(
    'prev_ImproveScore',
    'prev_Quality_level',
    'prev_target_label',
    'past_2yr_fail_rate',
    'past_3yr_avg_ImproveScore',
    'consecutive_pass_years',
    'consecutive_fail_years',
    'ever_failed_before'
) 
$leakageSafeTargetColumns = @('target_label', 'Quality_level')

Add-DictionaryRows -DictionaryRows $dataDictionaryRows -SheetName 'model_dataset_benchmark' -Columns $benchmarkFeatureColumns -Role 'feature' -DescriptionPrefix 'Benchmark feature: '
Add-DictionaryRows -DictionaryRows $dataDictionaryRows -SheetName 'model_dataset_benchmark' -Columns $benchmarkTargetColumns -Role 'target' -DescriptionPrefix 'Benchmark target: '
Add-DictionaryRows -DictionaryRows $dataDictionaryRows -SheetName 'model_dataset_leakage_safe' -Columns $leakageSafeFeatureColumns -Role 'feature' -DescriptionPrefix 'Leakage-safe feature: '
Add-DictionaryRows -DictionaryRows $dataDictionaryRows -SheetName 'model_dataset_leakage_safe' -Columns $leakageSafeTargetColumns -Role 'target' -DescriptionPrefix 'Leakage-safe target: '

$modelFeatureNotesRows = @(
    [pscustomobject]@{
        sheet_name    = 'model_dataset_benchmark'
        note_type     = 'purpose'
        feature_group = 'all_features'
        note          = 'Classification benchmark dataset using current-year indicator_* plus organization fields.'
    },
    [pscustomobject]@{
        sheet_name    = 'model_dataset_benchmark'
        note_type     = 'warning'
        feature_group = 'indicator_*'
        note          = 'Contains intentional data leakage because target_label and Quality_level are derived from same-year indicator_* values.'
    },
    [pscustomobject]@{
        sheet_name    = 'model_dataset_leakage_safe'
        note_type     = 'purpose'
        feature_group = 'historical_features'
        note          = 'Leakage-reduced dataset for forward prediction using prior-year history and aggregate historical behavior.'
    },
    [pscustomobject]@{
        sheet_name    = 'model_dataset_leakage_safe'
        note_type     = 'feature_rule'
        feature_group = 'prev_indicator_*'
        note          = 'Previous-year indicator features are populated only when the immediately previous Buddhist year exists for the same center.'
    },
    [pscustomobject]@{
        sheet_name    = 'model_dataset_leakage_safe'
        note_type     = 'feature_rule'
        feature_group = 'past_2yr_fail_rate'
        note          = 'Calculated from up to two prior years only; blank when no prior years exist.'
    },
    [pscustomobject]@{
        sheet_name    = 'model_dataset_leakage_safe'
        note_type     = 'feature_rule'
        feature_group = 'past_3yr_avg_ImproveScore'
        note          = 'Average of ImproveScore over up to three prior years; blank when no prior years exist.'
    },
    [pscustomobject]@{
        sheet_name    = 'model_dataset_leakage_safe'
        note_type     = 'feature_rule'
        feature_group = 'consecutive_pass_years/consecutive_fail_years'
        note          = 'Counts only contiguous prior years immediately before the current year.'
    },
    [pscustomobject]@{
        sheet_name    = 'model_dataset_leakage_safe'
        note_type     = 'feature_rule'
        feature_group = 'center_age_years'
        note          = 'Derived from opening_date_iso to year-end of assessment_year_ce; blank when opening date is unavailable or invalid.'
    }
)

$qaSummaryPath = Join-Path $csvExportDir 'qa_summary.csv'
$dataDictionaryPath = Join-Path $csvExportDir 'data_dictionary.csv'
$modelFeatureNotesPath = Join-Path $csvExportDir 'model_feature_notes.csv'

Export-RowsToCsv -Rows $qaSummaryRows -Path $qaSummaryPath
Export-RowsToCsv -Rows $dataDictionaryRows -Path $dataDictionaryPath
Export-RowsToCsv -Rows $modelFeatureNotesRows -Path $modelFeatureNotesPath

$sheetSpecs = @(
    @{
        Name = 'model_dataset_benchmark'
        CsvPath = $benchmarkPath
    },
    @{
        Name = 'model_dataset_leakage_safe'
        CsvPath = $leakageSafePath
    },
    @{
        Name = 'qa_summary'
        CsvPath = $qaSummaryPath
    },
    @{
        Name = 'data_dictionary'
        CsvPath = $dataDictionaryPath
    },
    @{
        Name = 'model_feature_notes'
        CsvPath = $modelFeatureNotesPath
    }
)

Remove-IfExists -Path $tempXlsxRoot
New-Directory -Path $tempXlsxRoot
New-Directory -Path (Join-Path $tempXlsxRoot '_rels')
New-Directory -Path (Join-Path $tempXlsxRoot 'xl')
New-Directory -Path (Join-Path $tempXlsxRoot 'xl\_rels')
New-Directory -Path (Join-Path $tempXlsxRoot 'xl\worksheets')

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

$sheetId = 1
$relationshipId = 2
foreach ($sheetSpec in $sheetSpecs) {
    $worksheetXml = Convert-CsvToWorksheetXml -CsvPath $sheetSpec.CsvPath
    $worksheetPath = Join-Path $tempXlsxRoot ('xl\worksheets\sheet' + $sheetId + '.xml')
    Write-Utf8File -Path $worksheetPath -Content $worksheetXml

    [void]$contentTypes.AppendLine('<Override PartName="/xl/worksheets/sheet' + $sheetId + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    [void]$workbook.AppendLine('<sheet name="' + (Escape-Xml $sheetSpec.Name) + '" sheetId="' + $sheetId + '" r:id="rId' + $relationshipId + '"/>')
    [void]$workbookRels.AppendLine('<Relationship Id="rId' + $relationshipId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + $sheetId + '.xml"/>')

    $sheetId++
    $relationshipId++
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
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
'@

Write-Utf8File -Path (Join-Path $tempXlsxRoot '[Content_Types].xml') -Content $contentTypes.ToString()
Write-Utf8File -Path (Join-Path $tempXlsxRoot '_rels\.rels') -Content $rootRels
Write-Utf8File -Path (Join-Path $tempXlsxRoot 'xl\workbook.xml') -Content $workbook.ToString()
Write-Utf8File -Path (Join-Path $tempXlsxRoot 'xl\_rels\workbook.xml.rels') -Content $workbookRels.ToString()
Write-Utf8File -Path (Join-Path $tempXlsxRoot 'xl\styles.xml') -Content $styles

Remove-IfExists -Path $outputPath
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempXlsxRoot, $outputPath)

Write-Output ('OUTPUT_XLSX=' + $outputPath)
