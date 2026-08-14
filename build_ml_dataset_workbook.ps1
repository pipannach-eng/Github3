$ErrorActionPreference = 'Stop'

$sourceDir = 'C:\Users\User\Documents\G3\Data set'
$outputDir = 'C:\Users\User\Documents\G3\prepared_dataset_ml'
$xlsxPath = Join-Path $outputDir 'ml_dataset_workbook.xlsx'
$tempExportDir = Join-Path $outputDir 'csv_exports'
$tempXlsxRoot = Join-Path $outputDir 'xlsx_build'

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

function Remove-IfExists {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path) {
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force
        }
        catch {
            return $false
        }
    }

    return $true
}

function New-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Normalize-Text {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) {
        return ''
    }

    return ([string]$Value).Trim()
}

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

function Split-ImprovementCodes {
    param([string]$Value)

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
        if ($singleLine -notmatch '^\((\d+)\)\s*([0-9.]+)(.*)$') {
            continue
        }

        $indicatorSeq = [int]$Matches[1]
        $indicatorCode = $Matches[2].Trim()
        $remainder = $Matches[3].Trim()

        if ($remainder -match '^(\S)\s+(.*)$') {
            $indicatorCode = $indicatorCode + $Matches[1]
            $remainder = $Matches[2].Trim()
        }

        if ([string]::IsNullOrWhiteSpace($remainder)) {
            continue
        }

        $indicators.Add([pscustomobject]@{
            indicator_seq  = $indicatorSeq
            indicator_code = $indicatorCode
            indicator_text = $remainder
            indicator_column = 'indicator_' + ($indicatorCode.Replace('.', '_'))
            source_row     = $rowIndex + 1
            raw_value      = $singleLine
        })
    }

    return $indicators
}

function Convert-BuddhistDateToIso {
    param([string]$RawValue)

    $raw = Normalize-Text $RawValue
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return [pscustomobject]@{
            opening_date_iso    = ''
            opening_date_status = 'missing'
        }
    }

    if ($raw -eq '17') {
        return [pscustomobject]@{
            opening_date_iso    = ''
            opening_date_status = 'unknown_marker_17'
        }
    }

    if ($raw -notmatch '^(\d{1,2})/(\d{1,2})/(\d{2,4})$') {
        return [pscustomobject]@{
            opening_date_iso    = ''
            opening_date_status = 'invalid'
        }
    }

    $day = [int]$Matches[1]
    $month = [int]$Matches[2]
    $year = [int]$Matches[3]
    if ($year -ge 2400) {
        $year = $year - 543
    }

    try {
        $date = [datetime]::new($year, $month, $day)
        return [pscustomobject]@{
            opening_date_iso    = $date.ToString('yyyy-MM-dd')
            opening_date_status = 'parsed'
        }
    }
    catch {
        return [pscustomobject]@{
            opening_date_iso    = ''
            opening_date_status = 'invalid'
        }
    }
}

function New-ThaiText {
    param([int[]]$CodePoints)
    return (-join ($CodePoints | ForEach-Object { [char]$_ }))
}

function Get-TargetLabel {
    param([string]$QualityLevel)
    if ($QualityLevel -eq 'D') {
        return (New-ThaiText -CodePoints @(0x0E44, 0x0E21, 0x0E48, 0x0E1C, 0x0E48, 0x0E32, 0x0E19))
    }

    return (New-ThaiText -CodePoints @(0x0E1C, 0x0E48, 0x0E32, 0x0E19))
}

function Get-QualityLevel {
    param([int]$ImproveScore)

    if ($ImproveScore -eq 0) { return 'A' }
    if ($ImproveScore -le 7) { return 'B' }
    if ($ImproveScore -le 15) { return 'C' }
    return 'D'
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

function Get-CenterAgeYears {
    param(
        [AllowNull()][object]$OpeningDateIso,
        [AllowNull()][object]$AssessmentYearCe
    )

    $openingDateText = Normalize-Text $OpeningDateIso
    $yearText = Normalize-Text $AssessmentYearCe
    if ([string]::IsNullOrWhiteSpace($openingDateText) -or [string]::IsNullOrWhiteSpace($yearText)) {
        return ''
    }

    try {
        $openingDate = [datetime]::ParseExact($openingDateText, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
        $referenceDate = [datetime]::new([int]$yearText, 12, 31)
        $ageYears = [math]::Round((($referenceDate - $openingDate).TotalDays / 365.25), 4)
        if ($ageYears -lt 0) {
            return ''
        }

        return $ageYears
    }
    catch {
        return ''
    }
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

function Export-SheetsToXlsx {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$SheetSpecs,
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    [void](Remove-IfExists -Path $tempXlsxRoot)
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
    foreach ($sheetSpec in $SheetSpecs) {
        $csvPath = Join-Path $tempExportDir ($sheetSpec.FileBase + '.csv')
        $sheetSpec.Rows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
        $worksheetXml = Convert-CsvToWorksheetXml -CsvPath $csvPath
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

    Write-Utf8File -Path (Join-Path $tempXlsxRoot '[Content_Types].xml') -Content $contentTypes.ToString()
    Write-Utf8File -Path (Join-Path $tempXlsxRoot '_rels\.rels') -Content $rootRels
    Write-Utf8File -Path (Join-Path $tempXlsxRoot 'xl\workbook.xml') -Content $workbook.ToString()
    Write-Utf8File -Path (Join-Path $tempXlsxRoot 'xl\_rels\workbook.xml.rels') -Content $workbookRels.ToString()
    Write-Utf8File -Path (Join-Path $tempXlsxRoot 'xl\styles.xml') -Content $styles

    $resolvedOutputPath = $OutputPath
    $removed = [bool](Remove-IfExists -Path $resolvedOutputPath)
    if (-not $removed) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($OutputPath)
        $extension = [System.IO.Path]::GetExtension($OutputPath)
        $directory = Split-Path -Path $OutputPath -Parent
        $resolvedOutputPath = Join-Path $directory ($baseName + '_updated_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + $extension)
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempXlsxRoot, $resolvedOutputPath)
    return $resolvedOutputPath
}

New-Directory -Path $outputDir
[void](Remove-IfExists -Path $tempExportDir)
New-Directory -Path $tempExportDir

$annualFiles = @(
    @{ Year = 2563; File = (Resolve-RequiredFile -Pattern '*2563*REP004-IMPROVEMENT 12*.xls').Name },
    @{ Year = 2564; File = (Resolve-RequiredFile -Pattern '*2564*REP004-IMPROVEMENT 12*.xls').Name },
    @{ Year = 2565; File = (Resolve-RequiredFile -Pattern '*2565*REP004-IMPROVEMENT 12*.xls').Name },
    @{ Year = 2566; File = (Resolve-RequiredFile -Pattern '*2566*REP004-IMPROVEMENT 12*.xls').Name },
    @{ Year = 2567; File = (Resolve-RequiredFile -Pattern '*2567*REP004-IMPROVEMENT 12*.xls').Name },
    @{ Year = 2568; File = (Resolve-RequiredFile -Pattern '*2568*REP004-IMPROVEMENT 12*.xls').Name }
)

$referenceFile = (Resolve-RequiredFile -Pattern '*2568*REP006 13*.xls').FullName
$referenceTable = Get-ExcelTable -Path $referenceFile
$indicators = Parse-IndicatorReference -Table $referenceTable
if ($indicators.Count -ne 75) {
    throw "Expected 75 indicators, found $($indicators.Count)."
}

$indicatorLookup = @{}
foreach ($indicator in $indicators) {
    $indicatorLookup[$indicator.indicator_code] = $indicator
}

$datasetRows = New-Object System.Collections.Generic.List[object]
$unknownImprovementRows = New-Object System.Collections.Generic.List[object]
$rowCountRows = New-Object System.Collections.Generic.List[object]

foreach ($annualFile in $annualFiles) {
    $path = Join-Path $sourceDir $annualFile.File
    $table = Get-ExcelTable -Path $path
    $validRows = 0

    for ($rowIndex = 2; $rowIndex -lt $table.Rows.Count; $rowIndex++) {
        $row = $table.Rows[$rowIndex]
        $userCode = Normalize-Text $row[8]
        $centerName = Normalize-Text $row[9]
        if ([string]::IsNullOrWhiteSpace($userCode) -or [string]::IsNullOrWhiteSpace($centerName)) {
            continue
        }

        $validRows++
        $openingDateRaw = Normalize-Text $row[10]
        $parsedDate = Convert-BuddhistDateToIso -RawValue $openingDateRaw
        $improvementItemsRaw = Normalize-Text $row[11]
        $improvementCodesRaw = Split-ImprovementCodes -Value $improvementItemsRaw
        $resolvedCodes = New-Object System.Collections.Generic.List[string]
        $unknownCodes = New-Object System.Collections.Generic.List[string]

        foreach ($code in $improvementCodesRaw) {
            $resolvedCode = Resolve-IndicatorCode -Code $code -IndicatorLookup $indicatorLookup
            if ($resolvedCode) {
                if (-not $resolvedCodes.Contains($resolvedCode)) {
                    $resolvedCodes.Add($resolvedCode)
                }
            }
            else {
                if (-not $unknownCodes.Contains($code)) {
                    $unknownCodes.Add($code)
                }

                $unknownImprovementRows.Add([pscustomobject]@{
                    assessment_year_be   = $annualFile.Year
                    assessment_year_ce   = $annualFile.Year - 543
                    source_file          = $annualFile.File
                    source_row           = $rowIndex + 1
                    user_code            = $userCode
                    center_name          = $centerName
                    unknown_code         = $code
                    improvement_items_raw = $improvementItemsRaw
                })
            }
        }

        $record = [ordered]@{
            assessment_year_be           = $annualFile.Year
            assessment_year_ce           = $annualFile.Year - 543
            source_file                  = $annualFile.File
            source_row                   = $rowIndex + 1
            source_sheet                 = 'Sheet1'
            record_key                   = "$($annualFile.Year)|$userCode"
            ministry                     = Normalize-Text $row[1]
            sub_agency_1                 = Normalize-Text $row[2]
            sub_agency_2                 = Normalize-Text $row[3]
            sub_agency_3                 = Normalize-Text $row[4]
            province                     = Normalize-Text $row[5]
            district                     = Normalize-Text $row[6]
            subdistrict                  = Normalize-Text $row[7]
            user_code                    = $userCode
            center_name                  = $centerName
            center_name_raw              = $centerName
            opening_date_raw             = $openingDateRaw
            opening_date_iso             = $parsedDate.opening_date_iso
            opening_date_status          = $parsedDate.opening_date_status
            improvement_items_raw        = $improvementItemsRaw
            improvement_items_mapped     = ($resolvedCodes -join ',')
            improvement_item_count_raw   = $improvementCodesRaw.Count
            improvement_item_count_mapped = $resolvedCodes.Count
            unknown_improvement_code_count = $unknownCodes.Count
            unknown_improvement_codes    = ($unknownCodes -join ',')
            has_unknown_improvement_code = if ($unknownCodes.Count -gt 0) { 1 } else { 0 }
        }

        foreach ($indicator in $indicators) {
            $record[$indicator.indicator_column] = if ($resolvedCodes.Contains($indicator.indicator_code)) { 1 } else { 0 }
        }

        $improveScore = 0
        foreach ($indicator in $indicators) {
            $improveScore += [int]$record[$indicator.indicator_column]
        }
        $record['ImproveScore'] = $improveScore
        $record['Quality_level'] = Get-QualityLevel -ImproveScore $improveScore
        $record['target_label'] = Get-TargetLabel -QualityLevel $record['Quality_level']

        $datasetRows.Add([pscustomobject]$record)
    }

    $rowCountRows.Add([pscustomobject]@{
        assessment_year_be = $annualFile.Year
        assessment_year_ce = $annualFile.Year - 543
        row_count          = $validRows
        source_file        = $annualFile.File
    })
}

$distinctCenters = @($datasetRows | Sort-Object user_code, center_name | Group-Object user_code | ForEach-Object { $_.Group[0] })
$codeMap = @{}
$codeMappingRows = New-Object System.Collections.Generic.List[object]
for ($index = 0; $index -lt $distinctCenters.Count; $index++) {
    $center = $distinctCenters[$index]
    $anonCode = 'SCH_G3_' + ($index + 1).ToString('0000')
    $codeMap[$center.user_code] = $anonCode
}

foreach ($group in ($datasetRows | Group-Object user_code | Sort-Object Name)) {
    $first = $group.Group | Sort-Object assessment_year_be | Select-Object -First 1
    $last = $group.Group | Sort-Object assessment_year_be -Descending | Select-Object -First 1
    $codeMappingRows.Add([pscustomobject]@{
        center_code_anon   = $codeMap[$group.Name]
        source_key_type    = 'user_code'
        source_key_value   = $group.Name
        center_name        = $first.center_name
        ministry           = $first.ministry
        first_year_be      = $first.assessment_year_be
        last_year_be       = $last.assessment_year_be
        record_count       = $group.Count
    })
}

$finalDatasetRows = New-Object System.Collections.Generic.List[object]
foreach ($row in $datasetRows) {
    $ordered = [ordered]@{ center_code_anon = $codeMap[$row.user_code] }
    foreach ($property in $row.PSObject.Properties.Name) {
        $ordered[$property] = $row.$property
    }
    $finalDatasetRows.Add([pscustomobject]$ordered)
}

$improvementIndicatorRows = New-Object System.Collections.Generic.List[object]
foreach ($row in $finalDatasetRows) {
    $indicatorRow = [ordered]@{
        assessment_year_be     = $row.assessment_year_be
        assessment_year_ce     = $row.assessment_year_ce
        record_key             = $row.record_key
        center_code_anon       = $row.center_code_anon
        improvement_items_raw  = $row.improvement_items_raw
        ImproveScore           = $row.ImproveScore
        Quality_level          = $row.Quality_level
        target_label           = $row.target_label
        improvement_data_status = if ($row.has_unknown_improvement_code -eq 1) {
            'has_unknown_codes'
        }
        elseif ([string]::IsNullOrWhiteSpace($row.improvement_items_raw)) {
            'empty'
        }
        else {
            'parsed'
        }
    }

    foreach ($indicator in $indicators) {
        $value = [int]$row.($indicator.indicator_column)
        if ($value -ne 0 -and $value -ne 1) {
            throw "Indicator column [$($indicator.indicator_column)] has non-binary value [$value] for record [$($row.record_key)]."
        }
        $indicatorRow[$indicator.indicator_column] = $value
    }

    $improvementIndicatorRows.Add([pscustomobject]$indicatorRow)
}

$improvementIndicatorsCleanRows = New-Object System.Collections.Generic.List[object]
$recordIndex = 1
foreach ($row in $finalDatasetRows) {
    $cleanRow = [ordered]@{
        no               = $recordIndex
        ministry         = $row.ministry
        affiliation_1    = $row.sub_agency_1
        affiliation_2    = $row.sub_agency_2
        affiliation_3    = $row.sub_agency_3
        center_code_anon = $row.center_code_anon
        opening_date_iso = $row.opening_date_iso
        opening_date_raw = $row.opening_date_raw
    }

    foreach ($indicator in $indicators) {
        $value = [int]$row.($indicator.indicator_column)
        if ($value -ne 0 -and $value -ne 1) {
            throw "Indicator column [$($indicator.indicator_column)] has non-binary value [$value] for clean row [$recordIndex]."
        }
        $cleanRow[$indicator.indicator_column] = $value
    }

    $cleanRow['ImproveScore'] = [int]$row.ImproveScore
    $cleanRow['Quality_level'] = [string]$row.Quality_level
    $cleanRow['target_label'] = [string]$row.target_label
    $improvementIndicatorsCleanRows.Add([pscustomobject]$cleanRow)
    $recordIndex++
}

$benchmarkDatasetRows = New-Object System.Collections.Generic.List[object]
foreach ($row in $finalDatasetRows) {
    $benchmarkRow = [ordered]@{
        ministry      = [string]$row.ministry
        affiliation_1 = [string]$row.sub_agency_1
        affiliation_2 = [string]$row.sub_agency_2
        affiliation_3 = [string]$row.sub_agency_3
    }

    foreach ($indicator in $indicators) {
        $benchmarkRow[$indicator.indicator_column] = [int]$row.($indicator.indicator_column)
    }

    $benchmarkRow['target_label'] = [string]$row.target_label
    $benchmarkRow['Quality_level'] = [string]$row.Quality_level
    $benchmarkDatasetRows.Add([pscustomobject]$benchmarkRow)
}

$passLabel = Get-TargetLabel -QualityLevel 'A'
$failLabel = Get-TargetLabel -QualityLevel 'D'
$leakageSafeDatasetRows = New-Object System.Collections.Generic.List[object]

foreach ($centerGroup in ($finalDatasetRows | Group-Object center_code_anon | Sort-Object Name)) {
    $history = @($centerGroup.Group | Sort-Object { [int]$_.assessment_year_be })

    for ($historyIndex = 0; $historyIndex -lt $history.Count; $historyIndex++) {
        $currentRow = $history[$historyIndex]
        $currentYearBe = [int]$currentRow.assessment_year_be
        $prevRow = $null

        if ($historyIndex -gt 0) {
            $candidatePrevRow = $history[$historyIndex - 1]
            if ([int]$candidatePrevRow.assessment_year_be -eq ($currentYearBe - 1)) {
                $prevRow = $candidatePrevRow
            }
        }

        $past2YearRows = @($history | Where-Object {
            $yearValue = [int]$_.assessment_year_be
            $yearValue -ge ($currentYearBe - 2) -and $yearValue -lt $currentYearBe
        })
        $past3YearRows = @($history | Where-Object {
            $yearValue = [int]$_.assessment_year_be
            $yearValue -ge ($currentYearBe - 3) -and $yearValue -lt $currentYearBe
        })
        $priorRows = if ($historyIndex -gt 0) { @($history[0..($historyIndex - 1)]) } else { @() }

        $past2YearFailRate = ''
        if ($past2YearRows.Count -gt 0) {
            $failCount = (@($past2YearRows | Where-Object { $_.target_label -eq $failLabel })).Count
            $past2YearFailRate = [math]::Round(($failCount / $past2YearRows.Count), 6)
        }

        $past3YearAvgImproveScore = ''
        if ($past3YearRows.Count -gt 0) {
            $scoreSum = 0
            foreach ($pastRow in $past3YearRows) {
                $scoreSum += [int]$pastRow.ImproveScore
            }
            $past3YearAvgImproveScore = [math]::Round(($scoreSum / $past3YearRows.Count), 6)
        }

        $consecutivePassYears = 0
        $consecutiveFailYears = 0
        $expectedYear = $currentYearBe - 1
        for ($priorIndex = $historyIndex - 1; $priorIndex -ge 0; $priorIndex--) {
            $priorRow = $history[$priorIndex]
            $priorYear = [int]$priorRow.assessment_year_be
            if ($priorYear -ne $expectedYear) {
                break
            }

            if ($priorRow.target_label -eq $passLabel) {
                $consecutivePassYears++
                $expectedYear--
            }
            else {
                break
            }
        }

        $expectedYear = $currentYearBe - 1
        for ($priorIndex = $historyIndex - 1; $priorIndex -ge 0; $priorIndex--) {
            $priorRow = $history[$priorIndex]
            $priorYear = [int]$priorRow.assessment_year_be
            if ($priorYear -ne $expectedYear) {
                break
            }

            if ($priorRow.target_label -eq $failLabel) {
                $consecutiveFailYears++
                $expectedYear--
            }
            else {
                break
            }
        }

        $leakageSafeRow = [ordered]@{
            assessment_year_be         = [int]$currentRow.assessment_year_be
            assessment_year_ce         = [int]$currentRow.assessment_year_ce
            ministry                   = [string]$currentRow.ministry
            affiliation_1              = [string]$currentRow.sub_agency_1
            affiliation_2              = [string]$currentRow.sub_agency_2
            affiliation_3              = [string]$currentRow.sub_agency_3
            center_age_years           = Get-CenterAgeYears -OpeningDateIso $currentRow.opening_date_iso -AssessmentYearCe $currentRow.assessment_year_ce
        }

        foreach ($indicator in $indicators) {
            $leakageSafeRow['prev_' + $indicator.indicator_column] = if ($null -ne $prevRow) {
                [int]$prevRow.($indicator.indicator_column)
            }
            else {
                ''
            }
        }

        $leakageSafeRow['prev_ImproveScore'] = if ($null -ne $prevRow) { [int]$prevRow.ImproveScore } else { '' }
        $leakageSafeRow['prev_Quality_level'] = if ($null -ne $prevRow) { [string]$prevRow.Quality_level } else { '' }
        $leakageSafeRow['prev_target_label'] = if ($null -ne $prevRow) { [string]$prevRow.target_label } else { '' }
        $leakageSafeRow['past_2yr_fail_rate'] = $past2YearFailRate
        $leakageSafeRow['past_3yr_avg_ImproveScore'] = $past3YearAvgImproveScore
        $leakageSafeRow['consecutive_pass_years'] = $consecutivePassYears
        $leakageSafeRow['consecutive_fail_years'] = $consecutiveFailYears
        $leakageSafeRow['ever_failed_before'] = if ((@($priorRows | Where-Object { $_.target_label -eq $failLabel })).Count -gt 0) { 1 } else { 0 }
        $leakageSafeRow['target_label'] = [string]$currentRow.target_label
        $leakageSafeRow['Quality_level'] = [string]$currentRow.Quality_level
        $leakageSafeDatasetRows.Add([pscustomobject]$leakageSafeRow)
    }
}

$duplicateRows = New-Object System.Collections.Generic.List[object]
foreach ($group in ($finalDatasetRows | Group-Object record_key | Where-Object { $_.Count -gt 1 })) {
    foreach ($row in $group.Group) {
        $duplicateRows.Add([pscustomobject]@{
            record_key          = $row.record_key
            duplicate_count     = $group.Count
            assessment_year_be  = $row.assessment_year_be
            user_code           = $row.user_code
            center_code_anon    = $row.center_code_anon
            center_name         = $row.center_name
            source_file         = $row.source_file
            source_row          = $row.source_row
        })
    }
}
if ($duplicateRows.Count -eq 0) {
    $duplicateRows.Add([pscustomobject]@{
        record_key          = ''
        duplicate_count     = 0
        assessment_year_be  = ''
        user_code           = ''
        center_code_anon    = ''
        center_name         = ''
        source_file         = ''
        source_row          = ''
    })
}

$missingSummaryRows = New-Object System.Collections.Generic.List[object]
$columnNames = @($finalDatasetRows[0].PSObject.Properties.Name)
foreach ($columnName in $columnNames) {
    $missingCount = 0
    foreach ($row in $finalDatasetRows) {
        $value = Normalize-Text $row.$columnName
        if ([string]::IsNullOrWhiteSpace($value)) {
            $missingCount++
        }
    }

    $missingSummaryRows.Add([pscustomobject]@{
        column_name    = $columnName
        missing_count  = $missingCount
        total_rows     = $finalDatasetRows.Count
        missing_ratio  = [math]::Round(($missingCount / $finalDatasetRows.Count), 6)
    })
}

$targetSummaryRows = @(
    [pscustomobject]@{
        metric = 'target_label'
        value  = Get-TargetLabel -QualityLevel 'A'
        count  = (@($finalDatasetRows | Where-Object { $_.target_label -eq (Get-TargetLabel -QualityLevel 'A') })).Count
    },
    [pscustomobject]@{
        metric = 'target_label'
        value  = Get-TargetLabel -QualityLevel 'D'
        count  = (@($finalDatasetRows | Where-Object { $_.target_label -eq (Get-TargetLabel -QualityLevel 'D') })).Count
    }
)

foreach ($level in @('A', 'B', 'C', 'D')) {
    $targetSummaryRows += [pscustomobject]@{
        metric = 'quality_level'
        value  = $level
        count  = (@($finalDatasetRows | Where-Object { $_.Quality_level -eq $level })).Count
    }
}

$targetDefinitionRows = @(
    [pscustomobject]@{
        target_name = 'target_label'
        rule_name   = 'binary_pass_fail_from_quality_level'
        rule_logic  = 'fail when Quality_level = D; pass when Quality_level is A, B, or C'
        notes       = 'Derived target label based on the requested Quality_level rule.'
    },
    [pscustomobject]@{
        target_name = 'quality_level'
        rule_name   = 'quality_bucket_from_ImproveScore'
        rule_logic  = 'A=0, B=1-7, C=8-15, D>=16 using ImproveScore = sum(indicator_* across all 75 indicators)'
        notes       = 'Rule-based label prepared from indicator columns.'
    }
)

$referenceRows = $indicators | Select-Object `
    @{ Name = 'item_code'; Expression = { $_.indicator_code } }, `
    @{ Name = 'feature_column'; Expression = { $_.indicator_column } }, `
    @{ Name = 'indicator_text'; Expression = { $_.indicator_text } }, `
    @{ Name = 'source_row'; Expression = { $_.source_row } }, `
    @{ Name = 'source_raw'; Expression = { $_.raw_value } }

$indicatorBinaryCheckRows = New-Object System.Collections.Generic.List[object]
foreach ($indicator in $indicators) {
    $values = @($improvementIndicatorRows | ForEach-Object { [string]$_.($indicator.indicator_column) } | Sort-Object -Unique)
    $nonBinaryValues = @($values | Where-Object { $_ -ne '0' -and $_ -ne '1' })
    $indicatorBinaryCheckRows.Add([pscustomobject]@{
        feature_column    = $indicator.indicator_column
        unique_values     = ($values -join ',')
        is_binary_only    = if ($nonBinaryValues.Count -eq 0) { 1 } else { 0 }
        non_binary_values = ($nonBinaryValues -join ',')
    })
}

$mlCleanValidationRows = @(
    [pscustomobject]@{
        check_name   = 'indicator_binary_only'
        status       = if ((@($indicatorBinaryCheckRows | Where-Object { $_.is_binary_only -ne 1 }).Count) -eq 0) { 'pass' } else { 'fail' }
        detail       = 'All 75 indicator columns contain only 0/1 values.'
    },
    [pscustomobject]@{
        check_name   = 'ImproveScore_integer_only'
        status       = if ((@($improvementIndicatorsCleanRows | Where-Object { ([string]$_.ImproveScore) -notmatch '^-?\d+$' }).Count) -eq 0) { 'pass' } else { 'fail' }
        detail       = 'ImproveScore is integer for every row.'
    },
    [pscustomobject]@{
        check_name   = 'Quality_level_allowed_values'
        status       = if ((@($improvementIndicatorsCleanRows | Where-Object { $_.Quality_level -notin @('A','B','C','D') }).Count) -eq 0) { 'pass' } else { 'fail' }
        detail       = 'Quality_level contains only A/B/C/D.'
    },
    [pscustomobject]@{
        check_name   = 'target_label_allowed_values'
        status       = if ((@($improvementIndicatorsCleanRows | Where-Object { $_.target_label -notin @((Get-TargetLabel -QualityLevel 'A'), (Get-TargetLabel -QualityLevel 'D')) }).Count) -eq 0) { 'pass' } else { 'fail' }
        detail       = 'target_label contains only ผ่าน/ไม่ผ่าน.'
    },
    [pscustomobject]@{
        check_name   = 'center_name_removed'
        status       = if ((@($improvementIndicatorsCleanRows[0].PSObject.Properties.Name | Where-Object { $_ -eq 'center_name' }).Count) -eq 0) { 'pass' } else { 'fail' }
        detail       = 'improvement_indicators_clean does not contain center_name.'
    }
)

$rowCountTotal = ($rowCountRows | Measure-Object row_count -Sum).Sum
$rowCountRows.Add([pscustomobject]@{
    assessment_year_be = 'TOTAL'
    assessment_year_ce = ''
    row_count          = $rowCountTotal
    source_file        = 'all_years'
})

$sheetSpecs = @(
    @{ Name = 'dataset_clean'; FileBase = 'dataset_clean'; Rows = $finalDatasetRows },
    @{ Name = 'code_mapping'; FileBase = 'code_mapping'; Rows = $codeMappingRows },
    @{ Name = 'row_counts_by_year'; FileBase = 'row_counts_by_year'; Rows = $rowCountRows },
    @{ Name = 'indicator_reference'; FileBase = 'indicator_reference'; Rows = $referenceRows },
    @{ Name = 'improvement_indicators'; FileBase = 'improvement_indicators'; Rows = $improvementIndicatorRows },
    @{ Name = 'improvement_indicators_clean'; FileBase = 'improvement_indicators_clean'; Rows = $improvementIndicatorsCleanRows },
    @{ Name = 'model_dataset_benchmark'; FileBase = 'model_dataset_benchmark'; Rows = $benchmarkDatasetRows },
    @{ Name = 'model_dataset_leakage_safe'; FileBase = 'model_dataset_leakage_safe'; Rows = $leakageSafeDatasetRows },
    @{ Name = 'indicator_binary_check'; FileBase = 'indicator_binary_check'; Rows = $indicatorBinaryCheckRows },
    @{ Name = 'ml_clean_validation'; FileBase = 'ml_clean_validation'; Rows = $mlCleanValidationRows },
    @{ Name = 'missing_value_summary'; FileBase = 'missing_value_summary'; Rows = $missingSummaryRows },
    @{ Name = 'duplicate_record_keys'; FileBase = 'duplicate_record_keys'; Rows = $duplicateRows },
    @{ Name = 'unexpected_improvement_values'; FileBase = 'unexpected_improvement_values'; Rows = $unknownImprovementRows },
    @{ Name = 'target_definition'; FileBase = 'target_definition'; Rows = $targetDefinitionRows },
    @{ Name = 'target_summary'; FileBase = 'target_summary'; Rows = $targetSummaryRows }
)

foreach ($sheetSpec in $sheetSpecs) {
    if ($sheetSpec.Rows.Count -eq 0) {
        $sheetSpec.Rows = @([pscustomobject]@{ status = 'no_rows' })
    }
}

$finalXlsxPath = Export-SheetsToXlsx -SheetSpecs $sheetSpecs -OutputPath $xlsxPath

Write-Output ('OUTPUT_DIR=' + $outputDir)
Write-Output ('XLSX_PATH=' + $finalXlsxPath)
Write-Output ('TOTAL_ROWS=' + $finalDatasetRows.Count)
Write-Output ('TOTAL_CENTERS=' + $codeMappingRows.Count)
Write-Output ('UNKNOWN_IMPROVEMENT_ROWS=' + $unknownImprovementRows.Count)
