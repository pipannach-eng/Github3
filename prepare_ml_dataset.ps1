$ErrorActionPreference = 'Stop'

$sourceDir = 'C:\Users\User\Documents\G3\Data set'
$outputDir = 'C:\Users\User\Documents\G3\prepared_dataset'

function Resolve-RequiredFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Pattern
    )

    $matches = @(Get-ChildItem -LiteralPath $sourceDir -Filter $Pattern -File)
    if ($matches.Count -ne 1) {
        throw "Expected exactly one file for pattern [$Pattern], found [$($matches.Count)]."
    }

    return $matches[0]
}

$annualFiles = @(
    @{ Year = 2563; File = (Resolve-RequiredFile -Pattern '*2563*REP004-IMPROVEMENT 12*.xls').Name },
    @{ Year = 2564; File = (Resolve-RequiredFile -Pattern '*2564*REP004-IMPROVEMENT 12*.xls').Name },
    @{ Year = 2565; File = (Resolve-RequiredFile -Pattern '*2565*REP004-IMPROVEMENT 12*.xls').Name },
    @{ Year = 2566; File = (Resolve-RequiredFile -Pattern '*2566*REP004-IMPROVEMENT 12*.xls').Name },
    @{ Year = 2567; File = (Resolve-RequiredFile -Pattern '*2567*REP004-IMPROVEMENT 12*.xls').Name },
    @{ Year = 2568; File = (Resolve-RequiredFile -Pattern '*2568*REP004-IMPROVEMENT 12*.xls').Name }
)

$referenceFile = (Resolve-RequiredFile -Pattern '*2568*REP006 13*.xls').FullName

function Get-ExcelTable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$Path;Extended Properties='Excel 8.0;HDR=No;IMEX=1';"
    $connection = New-Object System.Data.OleDb.OleDbConnection($connectionString)
    $connection.Open()

    try {
        $schema = $connection.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)
        $sheet = @($schema | Where-Object { $_.TABLE_NAME -like '*$' -or $_.TABLE_NAME -like '*$''' })[0]
        if (-not $sheet) {
            throw "No worksheet found in $Path"
        }

        $command = $connection.CreateCommand()
        $command.CommandText = 'SELECT * FROM [' + [string]$sheet.TABLE_NAME + ']'

        $adapter = New-Object System.Data.OleDb.OleDbDataAdapter($command)
        $dataSet = New-Object System.Data.DataSet
        [void]$adapter.Fill($dataSet)
        return ,$dataSet.Tables[0]
    }
    finally {
        $connection.Close()
    }
}

function Normalize-Text {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    return ([string]$Value).Trim()
}

function Parse-IndicatorReference {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$Table
    )

    $indicators = New-Object System.Collections.Generic.List[object]

    for ($rowIndex = 4; $rowIndex -lt $Table.Rows.Count; $rowIndex++) {
        $raw = Normalize-Text $Table.Rows[$rowIndex][0]
        if ([string]::IsNullOrWhiteSpace($raw)) {
            continue
        }

        $singleLine = ($raw -replace '\s+', ' ').Trim()
        if ($singleLine -notmatch '^\((\d+)\)\s*([0-9.]+(?:\s*[ก-ฮ])?)\s*(.*)$') {
            continue
        }

        $indicators.Add([pscustomobject]@{
            indicator_seq  = [int]$Matches[1]
            indicator_code = ($Matches[2] -replace '\s+', '').Trim()
            indicator_text = $Matches[3].Trim()
            raw_value      = $singleLine
        })
    }

    return $indicators
}

function Split-ImprovementCodes {
    param(
        [AllowNull()]
        [string]$Value
    )

    $codes = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $codes
    }

    foreach ($part in ($Value -split ',')) {
        $code = $part.Trim()
        if (-not [string]::IsNullOrWhiteSpace($code) -and -not $codes.Contains($code)) {
            $codes.Add($code)
        }
    }

    return $codes
}

function Resolve-IndicatorCode {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Code,
        [Parameter(Mandatory = $true)]
        [hashtable]$IndicatorLookup
    )

    if ($IndicatorLookup.ContainsKey($Code)) {
        return $Code
    }

    if ($Code.Length -gt 1) {
        $lastChar = $Code.Substring($Code.Length - 1, 1)
        if ($lastChar -notmatch '[0-9.]') {
            $baseCode = $Code.Substring(0, $Code.Length - 1)
            if ($IndicatorLookup.ContainsKey($baseCode)) {
                return $baseCode
            }
        }
    }

    return $null
}

if (-not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

$referenceTable = Get-ExcelTable -Path $referenceFile
$indicators = Parse-IndicatorReference -Table $referenceTable
if ($indicators.Count -eq 0) {
    throw 'Could not parse any indicators from the reference file.'
}

$indicatorCodeSet = @{}
foreach ($indicator in $indicators) {
    $indicatorCodeSet[$indicator.indicator_code] = $true
}

$entityRows = New-Object System.Collections.Generic.List[object]
$longRows = New-Object System.Collections.Generic.List[object]
$wideRows = New-Object System.Collections.Generic.List[object]
$unknownCodeRows = New-Object System.Collections.Generic.List[object]
$summaryRows = New-Object System.Collections.Generic.List[object]

foreach ($fileInfo in $annualFiles) {
    $path = Join-Path $sourceDir $fileInfo.File
    $table = Get-ExcelTable -Path $path
    $dataRowCount = 0

    for ($rowIndex = 2; $rowIndex -lt $table.Rows.Count; $rowIndex++) {
        $row = $table.Rows[$rowIndex]

        $userCode = Normalize-Text $row[8]
        $centerName = Normalize-Text $row[9]

        if ([string]::IsNullOrWhiteSpace($userCode) -or [string]::IsNullOrWhiteSpace($centerName)) {
            continue
        }

        $dataRowCount++
        $improvementRaw = Normalize-Text $row[11]
        $improvementCodes = Split-ImprovementCodes -Value $improvementRaw
        $improvementLookup = @{}
        $resolvedImprovementCodes = New-Object System.Collections.Generic.List[string]
        foreach ($code in $improvementCodes) {
            $resolvedCode = Resolve-IndicatorCode -Code $code -IndicatorLookup $indicatorCodeSet
            if ($resolvedCode) {
                $improvementLookup[$resolvedCode] = $true
                if (-not $resolvedImprovementCodes.Contains($resolvedCode)) {
                    $resolvedImprovementCodes.Add($resolvedCode)
                }
            }
            else {
                $unknownCodeRows.Add([pscustomobject]@{
                    year                = $fileInfo.Year
                    user_code           = $userCode
                    center_name         = $centerName
                    unknown_code        = $code
                    improvement_codes   = $improvementRaw
                    source_file         = $fileInfo.File
                })
            }
        }

        $entity = [ordered]@{
            year                    = $fileInfo.Year
            source_file             = $fileInfo.File
            row_number_in_sheet     = $rowIndex + 1
            ministry                = Normalize-Text $row[1]
            sub_agency_1            = Normalize-Text $row[2]
            sub_agency_2            = Normalize-Text $row[3]
            sub_agency_3            = Normalize-Text $row[4]
            province                = Normalize-Text $row[5]
            district                = Normalize-Text $row[6]
            subdistrict             = Normalize-Text $row[7]
            user_code               = $userCode
            center_name             = $centerName
            open_date               = Normalize-Text $row[10]
            improvement_codes_raw   = $improvementRaw
            improvement_code_count  = $improvementCodes.Count
            improvement_codes_mapped = ($resolvedImprovementCodes -join ',')
        }
        $entityRows.Add([pscustomobject]$entity)

        $wide = [ordered]@{}
        foreach ($key in $entity.Keys) {
            $wide[$key] = $entity[$key]
        }

        foreach ($indicator in $indicators) {
            $label = if ($improvementLookup.ContainsKey($indicator.indicator_code)) { 1 } else { 0 }
            $columnName = 'indicator_' + $indicator.indicator_code.Replace('.', '_')
            $wide[$columnName] = $label

            $longRows.Add([pscustomobject]@{
                year                   = $fileInfo.Year
                source_file            = $fileInfo.File
                user_code              = $userCode
                center_name            = $centerName
                ministry               = Normalize-Text $row[1]
                sub_agency_1           = Normalize-Text $row[2]
                sub_agency_2           = Normalize-Text $row[3]
                sub_agency_3           = Normalize-Text $row[4]
                province               = Normalize-Text $row[5]
                district               = Normalize-Text $row[6]
                subdistrict            = Normalize-Text $row[7]
                open_date              = Normalize-Text $row[10]
                indicator_seq          = $indicator.indicator_seq
                indicator_code         = $indicator.indicator_code
                indicator_text         = $indicator.indicator_text
                needs_improvement      = $label
                improvement_codes_raw  = $improvementRaw
            })
        }

        $wideRows.Add([pscustomobject]$wide)
    }

    $summaryRows.Add([pscustomobject]@{
        year             = $fileInfo.Year
        source_file      = $fileInfo.File
        center_row_count = $dataRowCount
    })
}

$indicatorReferencePath = Join-Path $outputDir 'indicator_reference.csv'
$entityPath = Join-Path $outputDir 'ml_entities.csv'
$longPath = Join-Path $outputDir 'ml_dataset_long.csv'
$widePath = Join-Path $outputDir 'ml_dataset_wide.csv'
$unknownPath = Join-Path $outputDir 'unknown_improvement_codes.csv'
$summaryPath = Join-Path $outputDir 'dataset_summary.csv'
$readmePath = Join-Path $outputDir 'README_prepared_dataset.md'

$indicators | Export-Csv -LiteralPath $indicatorReferencePath -NoTypeInformation -Encoding UTF8
$entityRows | Export-Csv -LiteralPath $entityPath -NoTypeInformation -Encoding UTF8
$longRows | Export-Csv -LiteralPath $longPath -NoTypeInformation -Encoding UTF8
$wideRows | Export-Csv -LiteralPath $widePath -NoTypeInformation -Encoding UTF8
$unknownCodeRows | Export-Csv -LiteralPath $unknownPath -NoTypeInformation -Encoding UTF8
$summaryRows | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding UTF8

$readme = @"
# Prepared ML Dataset

Source folder:
- $sourceDir

Outputs:
- indicator_reference.csv: indicator list parsed from reference file column A
- ml_entities.csv: one row per center per year
- ml_dataset_long.csv: one row per center per year per indicator
- ml_dataset_wide.csv: one row per center per year with 0/1 indicator columns
- unknown_improvement_codes.csv: improvement codes found in annual files but not in the reference list
- dataset_summary.csv: imported center counts by year

Label meaning:
- needs_improvement = 1 means the indicator code appears in the source improvement-code field
- needs_improvement = 0 means the indicator code does not appear in that field

Assumptions:
- indicator reference comes from column A of the REP006 file
- rows without user_code or center_name are skipped
"@

Set-Content -LiteralPath $readmePath -Value $readme -Encoding UTF8

Write-Output ('OUTPUT_DIR=' + $outputDir)
Write-Output ('INDICATOR_COUNT=' + $indicators.Count)
Write-Output ('ENTITY_ROWS=' + $entityRows.Count)
Write-Output ('LONG_ROWS=' + $longRows.Count)
Write-Output ('WIDE_ROWS=' + $wideRows.Count)
Write-Output ('UNKNOWN_CODE_ROWS=' + $unknownCodeRows.Count)
