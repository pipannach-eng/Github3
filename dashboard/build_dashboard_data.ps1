$ErrorActionPreference = 'Stop'

$baseDir = 'C:\Users\User\Documents\G3'
$dashboardDir = Join-Path $baseDir 'dashboard'
$dataFile = Join-Path $dashboardDir 'data.js'
$sourceWorkbook = (
    Get-ChildItem -LiteralPath (Join-Path $baseDir 'prepared_dataset_ml') -Filter 'ml_dataset_workbook*.xlsx' -File |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
).FullName
$recordsPath = Join-Path $baseDir 'prepared_dataset_ml\csv_exports\improvement_indicators.csv'
$referencePath = Join-Path $baseDir 'prepared_dataset_ml\csv_exports\indicator_reference.csv'
$codeMappingPath = Join-Path $baseDir 'prepared_dataset_ml\csv_exports\code_mapping.csv'

function Normalize-Text {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim()
}

$records = Import-Csv -LiteralPath $recordsPath
$reference = Import-Csv -LiteralPath $referencePath
$codeMapping = Import-Csv -LiteralPath $codeMappingPath
$centerNameLookup = @{}
foreach ($row in $codeMapping) {
    $centerNameLookup[(Normalize-Text $row.center_code_anon)] = Normalize-Text $row.center_name
}

$indicatorColumns = @($reference.feature_column)

$recordRows = foreach ($row in $records) {
    $recordKey = Normalize-Text $row.record_key
    $ordered = [ordered]@{
        assessment_year_be = [int]$row.assessment_year_be
        assessment_year_ce = [int]$row.assessment_year_ce
        record_key = $recordKey
        center_code_anon = Normalize-Text $row.center_code_anon
        center_name = if ($centerNameLookup.ContainsKey((Normalize-Text $row.center_code_anon))) {
            $centerNameLookup[(Normalize-Text $row.center_code_anon)]
        } else {
            ''
        }
        improvement_items_raw = Normalize-Text $row.improvement_items_raw
        ImproveScore = [int]$row.ImproveScore
        Quality_level = Normalize-Text $row.Quality_level
        target_label = Normalize-Text $row.target_label
        improvement_data_status = Normalize-Text $row.improvement_data_status
    }

    foreach ($column in $indicatorColumns) {
        $ordered[$column] = [int]$row.$column
    }

    [pscustomobject]$ordered
}

$qualityCounts = $recordRows |
    Group-Object Quality_level |
    Sort-Object Name |
    ForEach-Object {
        [pscustomobject]@{
            quality_level = $_.Name
            count = $_.Count
        }
    }

$targetCounts = $recordRows |
    Group-Object target_label |
    Sort-Object Name |
    ForEach-Object {
        [pscustomobject]@{
            target_label = $_.Name
            count = $_.Count
        }
    }

$payload = [ordered]@{
    source = [ordered]@{
        workbook = $sourceWorkbook
        sheet = 'improvement_indicators'
        generated_at = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        total_records = $recordRows.Count
    }
    quality_counts = $qualityCounts
    target_counts = $targetCounts
    indicator_reference = @(
        $reference | ForEach-Object {
            [pscustomobject]@{
                item_code = Normalize-Text $_.item_code
                feature_column = Normalize-Text $_.feature_column
                indicator_text = Normalize-Text $_.indicator_text
                source_row = [int]$_.source_row
                source_raw = Normalize-Text $_.source_raw
            }
        }
    )
    records = @($recordRows)
}

$json = $payload | ConvertTo-Json -Depth 6 -Compress
$content = "window.DASHBOARD_DATA = $json;"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($dataFile, $content, $utf8NoBom)

Write-Output ('DATA_JS=' + $dataFile)
Write-Output ('TOTAL_RECORDS=' + $recordRows.Count)
