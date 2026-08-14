Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$Root = "C:\Users\User\Documents\G3"
$InputWorkbook = Join-Path $Root "prepared_dataset_ml\ml_dataset_workbook.xlsx"
$CsvPath = Join-Path $Root "prepared_dataset_ml\csv_exports\improvement_indicators_clean.csv"
$IndicatorRefPath = Join-Path $Root "prepared_dataset_ml\csv_exports\indicator_reference.csv"
$OutDir = Join-Path $Root "prepared_dataset_ml"
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutWorkbook = Join-Path $OutDir "knn_ml_model_results_$Timestamp.xlsx"

if (-not (Test-Path $InputWorkbook)) { throw "Input workbook not found: $InputWorkbook" }
if (-not (Test-Path $CsvPath)) { throw "CSV export not found: $CsvPath" }
if (-not (Test-Path $IndicatorRefPath)) { throw "Indicator reference not found: $IndicatorRefPath" }

function Normalize-Text {
    param([object]$Value)
    if ($null -eq $Value) { return "(missing)" }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return "(missing)" }
    return $text.Trim()
}

function Get-Percent {
    param([double]$Part, [double]$Total)
    if ($Total -eq 0) { return 0.0 }
    return [Math]::Round(($Part / $Total) * 100.0, 4)
}

function Shuffle-Array {
    param(
        [object[]]$Items,
        [System.Random]$Random
    )
    $copy = @($Items)
    for ($i = $copy.Count - 1; $i -gt 0; $i--) {
        $j = $Random.Next(0, $i + 1)
        $tmp = $copy[$i]
        $copy[$i] = $copy[$j]
        $copy[$j] = $tmp
    }
    return $copy
}

function Add-SheetData {
    param(
        [object]$Workbook,
        [string]$Name,
        [object[]]$Rows,
        [string[]]$Headers
    )
    $ws = $Workbook.Worksheets.Add()
    $ws.Name = $Name
    if ($Headers.Count -eq 0) { return $ws }

    $rowCount = [Math]::Max(1, $Rows.Count + 1)
    $colCount = $Headers.Count
    $values = New-Object "object[,]" $rowCount, $colCount

    for ($c = 0; $c -lt $colCount; $c++) {
        $values[0, $c] = $Headers[$c]
    }
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        foreach ($c in 0..($colCount - 1)) {
            $name = $Headers[$c]
            $value = $Rows[$r].$name
            if ($null -eq $value) { $value = "" }
            $values[$r + 1, $c] = $value
        }
    }

    $range = $ws.Range($ws.Cells(1, 1), $ws.Cells($rowCount, $colCount))
    $range.Value2 = $values
    $ws.Rows.Item(1).Font.Bold = $true
    $ws.Rows.Item(1).Interior.Color = 14277081
    $ws.Columns.AutoFit() | Out-Null
    return $ws
}

function Write-Block {
    param(
        [object]$Worksheet,
        [int]$StartRow,
        [int]$StartCol,
        [string]$Title,
        [object[]]$Rows,
        [string[]]$Headers
    )
    $Worksheet.Cells.Item($StartRow, $StartCol).Value2 = $Title
    $Worksheet.Cells.Item($StartRow, $StartCol).Font.Bold = $true
    $Worksheet.Cells.Item($StartRow, $StartCol).Font.Size = 14
    for ($c = 0; $c -lt $Headers.Count; $c++) {
        $cell = $Worksheet.Cells.Item($StartRow + 1, $StartCol + $c)
        $cell.Value2 = $Headers[$c]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 14277081
    }
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        for ($c = 0; $c -lt $Headers.Count; $c++) {
            $name = $Headers[$c]
            $value = $Rows[$r].$name
            if ($null -eq $value) { $value = "" }
            $Worksheet.Cells.Item($StartRow + 2 + $r, $StartCol + $c).Value2 = $value
        }
    }
}

function Add-Chart {
    param(
        [object]$Worksheet,
        [int]$Left,
        [int]$Top,
        [int]$Width,
        [int]$Height,
        [string]$Title,
        [object]$SourceRange,
        [int]$ChartType
    )
    $chartObj = $Worksheet.ChartObjects().Add($Left, $Top, $Width, $Height)
    $chart = $chartObj.Chart
    $chart.SetSourceData($SourceRange)
    $chart.ChartType = $ChartType
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $Title
    $chart.HasLegend = $false
    return $chartObj
}

function ConvertTo-ExcelColumnName {
    param([int]$ColumnNumber)
    $name = ""
    while ($ColumnNumber -gt 0) {
        $mod = ($ColumnNumber - 1) % 26
        $name = [char](65 + $mod) + $name
        $ColumnNumber = [int](($ColumnNumber - $mod) / 26)
    }
    return $name
}

function Escape-XmlText {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    return [System.Security.SecurityElement]::Escape([string]$Value)
}

function Test-NumericCell {
    param([object]$Value)
    if ($null -eq $Value) { return $false }
    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int] -or $Value -is [long] -or $Value -is [single] -or $Value -is [double] -or $Value -is [decimal]) { return $true }
    return $false
}

function Write-XlsxSheetXml {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$EntryName,
        [object[]]$Rows,
        [string[]]$Headers
    )
    $entry = $Zip.CreateEntry($EntryName)
    $stream = $entry.Open()
    $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
    try {
        $writer.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        $writer.Write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
        $rowNumber = 1
        $writer.Write("<row r=`"$rowNumber`">")
        for ($c = 0; $c -lt $Headers.Count; $c++) {
            $ref = "$(ConvertTo-ExcelColumnName ($c + 1))$rowNumber"
            $text = Escape-XmlText $Headers[$c]
            $writer.Write("<c r=`"$ref`" t=`"inlineStr`"><is><t>$text</t></is></c>")
        }
        $writer.Write('</row>')

        foreach ($row in $Rows) {
            $rowNumber++
            $writer.Write("<row r=`"$rowNumber`">")
            for ($c = 0; $c -lt $Headers.Count; $c++) {
                $name = $Headers[$c]
                $value = $row.$name
                $ref = "$(ConvertTo-ExcelColumnName ($c + 1))$rowNumber"
                if (Test-NumericCell $value) {
                    $num = ([string]$value).Replace(",", ".")
                    $writer.Write("<c r=`"$ref`"><v>$num</v></c>")
                } else {
                    $text = Escape-XmlText $value
                    $writer.Write("<c r=`"$ref`" t=`"inlineStr`"><is><t>$text</t></is></c>")
                }
            }
            $writer.Write('</row>')
        }
        $writer.Write('</sheetData></worksheet>')
    }
    finally {
        $writer.Dispose()
        $stream.Dispose()
    }
}

function Add-ZipTextEntry {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$EntryName,
        [string]$Text
    )
    $entry = $Zip.CreateEntry($EntryName)
    $stream = $entry.Open()
    $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
    try { $writer.Write($Text) }
    finally {
        $writer.Dispose()
        $stream.Dispose()
    }
}

function Export-WorkbookOpenXml {
    param(
        [string]$Path,
        [object[]]$Sheets
    )
    if (Test-Path $Path) { Remove-Item -LiteralPath $Path -Force }
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew)
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $overrides = New-Object System.Text.StringBuilder
        [void]$overrides.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
        [void]$overrides.Append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
        [void]$overrides.Append('<Default Extension="xml" ContentType="application/xml"/>')
        [void]$overrides.Append('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
        for ($i = 0; $i -lt $Sheets.Count; $i++) {
            [void]$overrides.Append("<Override PartName=`"/xl/worksheets/sheet$($i + 1).xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>")
        }
        [void]$overrides.Append('</Types>')
        Add-ZipTextEntry -Zip $zip -EntryName "[Content_Types].xml" -Text $overrides.ToString()

        Add-ZipTextEntry -Zip $zip -EntryName "_rels/.rels" -Text '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'

        $workbookXml = New-Object System.Text.StringBuilder
        [void]$workbookXml.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>')
        $relsXml = New-Object System.Text.StringBuilder
        [void]$relsXml.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
        for ($i = 0; $i -lt $Sheets.Count; $i++) {
            $sheetId = $i + 1
            $safeName = Escape-XmlText $Sheets[$i].name
            [void]$workbookXml.Append("<sheet name=`"$safeName`" sheetId=`"$sheetId`" r:id=`"rId$sheetId`"/>")
            [void]$relsXml.Append("<Relationship Id=`"rId$sheetId`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$sheetId.xml`"/>")
        }
        [void]$workbookXml.Append('</sheets></workbook>')
        [void]$relsXml.Append('</Relationships>')
        Add-ZipTextEntry -Zip $zip -EntryName "xl/workbook.xml" -Text $workbookXml.ToString()
        Add-ZipTextEntry -Zip $zip -EntryName "xl/_rels/workbook.xml.rels" -Text $relsXml.ToString()

        for ($i = 0; $i -lt $Sheets.Count; $i++) {
            Write-XlsxSheetXml -Zip $zip -EntryName "xl/worksheets/sheet$($i + 1).xml" -Rows $Sheets[$i].rows -Headers $Sheets[$i].headers
        }
    }
    finally {
        $zip.Dispose()
        $fs.Dispose()
    }
}

function New-ConfusionKey {
    param([string]$Actual, [string]$Predicted)
    return "$Actual`u{241F}$Predicted"
}

function Get-ClassificationMetrics {
    param(
        [string]$Target,
        [string]$Model,
        [string]$Method,
        [string[]]$Classes,
        [object[]]$Predictions
    )

    $matrix = @{}
    foreach ($actual in $Classes) {
        foreach ($pred in $Classes) {
            $matrix[(New-ConfusionKey $actual $pred)] = 0
        }
    }
    foreach ($p in $Predictions) {
        $key = New-ConfusionKey $p.actual $p.predicted
        if (-not $matrix.ContainsKey($key)) { $matrix[$key] = 0 }
        $matrix[$key] = [int]$matrix[$key] + 1
    }

    $total = [Math]::Max(1, $Predictions.Count)
    $correct = 0
    foreach ($class in $Classes) {
        $correct += [int]$matrix[(New-ConfusionKey $class $class)]
    }
    $accuracy = [Math]::Round($correct / $total, 6)

    $perClass = New-Object System.Collections.Generic.List[object]
    $macroPrecision = 0.0
    $macroRecall = 0.0
    $macroF1 = 0.0
    $weightedF1 = 0.0

    foreach ($class in $Classes) {
        $tp = [int]$matrix[(New-ConfusionKey $class $class)]
        $fp = 0
        $fn = 0
        foreach ($other in $Classes) {
            if ($other -ne $class) {
                $fp += [int]$matrix[(New-ConfusionKey $other $class)]
                $fn += [int]$matrix[(New-ConfusionKey $class $other)]
            }
        }
        $support = $tp + $fn
        $precision = if (($tp + $fp) -eq 0) { 0.0 } else { $tp / ($tp + $fp) }
        $recall = if (($tp + $fn) -eq 0) { 0.0 } else { $tp / ($tp + $fn) }
        $f1 = if (($precision + $recall) -eq 0) { 0.0 } else { 2.0 * $precision * $recall / ($precision + $recall) }

        $macroPrecision += $precision
        $macroRecall += $recall
        $macroF1 += $f1
        $weightedF1 += $f1 * $support

        $perClass.Add([pscustomobject]@{
            target = $Target
            model = $Model
            evaluation_method = $Method
            class = $class
            precision = [Math]::Round($precision, 6)
            recall = [Math]::Round($recall, 6)
            "F1-score" = [Math]::Round($f1, 6)
            support = $support
            TP = $tp
            FP = $fp
            FN = $fn
        })
    }

    $classCount = [Math]::Max(1, $Classes.Count)
    $overall = [pscustomobject]@{
        target = $Target
        evaluation_method = $Method
        model = $Model
        accuracy = $accuracy
        "macro precision" = [Math]::Round($macroPrecision / $classCount, 6)
        "macro recall" = [Math]::Round($macroRecall / $classCount, 6)
        "macro F1" = [Math]::Round($macroF1 / $classCount, 6)
        "weighted F1" = [Math]::Round($weightedF1 / $total, 6)
    }

    $matrixRows = New-Object System.Collections.Generic.List[object]
    foreach ($actual in $Classes) {
        foreach ($pred in $Classes) {
            $matrixRows.Add([pscustomobject]@{
                target = $Target
                model = $Model
                evaluation_method = $Method
                actual_class = $actual
                predicted_class = $pred
                count = [int]$matrix[(New-ConfusionKey $actual $pred)]
            })
        }
    }

    return [pscustomobject]@{
        overall = $overall
        per_class = $perClass
        confusion = $matrixRows
    }
}

Write-Host "Reading data from improvement_indicators_clean..."
$rows = @(Import-Csv -Path $CsvPath -Encoding UTF8)
$indicatorRef = @(Import-Csv -Path $IndicatorRefPath -Encoding UTF8)
if ($rows.Count -eq 0) { throw "No rows found in $CsvPath" }

$allColumns = @($rows[0].PSObject.Properties.Name)
$indicatorCols = @($allColumns | Where-Object { $_ -like "indicator_*" } | Sort-Object)
if ($indicatorCols.Count -ne 75) { throw "Expected 75 indicator columns, found $($indicatorCols.Count)." }

$requiredCols = @("ministry", "affiliation_1", "affiliation_2", "affiliation_3", "Quality_level", "target_label", "ImproveScore")
foreach ($col in $requiredCols) {
    if ($allColumns -notcontains $col) { throw "Required column missing: $col" }
}

$invalidIndicatorRows = New-Object System.Collections.Generic.List[object]
for ($i = 0; $i -lt $rows.Count; $i++) {
    foreach ($col in $indicatorCols) {
        $raw = Normalize-Text $rows[$i].$col
        if ($raw -notin @("0", "1")) {
            $invalidIndicatorRows.Add([pscustomobject]@{
                row_number = $i + 2
                column = $col
                value = $raw
            })
        }
        $rows[$i].$col = [int]$raw
    }
}
if ($invalidIndicatorRows.Count -gt 0) {
    throw "Indicator columns contain non-binary values. First issue: $($invalidIndicatorRows[0] | ConvertTo-Json -Compress)"
}

$categoricalCols = @("ministry", "affiliation_1", "affiliation_2", "affiliation_3")
$categoryLevels = @{}
foreach ($col in $categoricalCols) {
    $categoryLevels[$col] = @($rows | ForEach-Object { Normalize-Text $_.$col } | Sort-Object -Unique)
}

$featureNames = New-Object System.Collections.Generic.List[string]
foreach ($col in $categoricalCols) {
    foreach ($level in $categoryLevels[$col]) {
        $featureNames.Add("$col=$level")
    }
}
foreach ($col in $indicatorCols) { $featureNames.Add($col) }

$vectors = New-Object System.Collections.Generic.List[object]
foreach ($row in $rows) {
    $vector = New-Object double[] $featureNames.Count
    $pos = 0
    foreach ($col in $categoricalCols) {
        $value = Normalize-Text $row.$col
        foreach ($level in $categoryLevels[$col]) {
            $vector[$pos] = if ($value -eq $level) { 1.0 } else { 0.0 }
            $pos++
        }
    }
    foreach ($col in $indicatorCols) {
        $vector[$pos] = [double]$row.$col
        $pos++
    }
    $vectors.Add($vector)
}

function New-StratifiedSplit {
    param(
        [object[]]$Rows,
        [string]$TargetColumn,
        [double]$TrainRatio,
        [int]$Seed
    )
    $rnd = [System.Random]::new($Seed)
    $train = New-Object System.Collections.Generic.HashSet[int]
    $test = New-Object System.Collections.Generic.HashSet[int]
    $byClass = @{}
    for ($i = 0; $i -lt $Rows.Count; $i++) {
        $class = Normalize-Text $Rows[$i].$TargetColumn
        if (-not $byClass.ContainsKey($class)) {
            $byClass[$class] = New-Object System.Collections.Generic.List[int]
        }
        $byClass[$class].Add($i)
    }
    foreach ($class in $byClass.Keys) {
        $indices = @(Shuffle-Array -Items @($byClass[$class]) -Random $rnd)
        $n = $indices.Count
        $trainN = [int][Math]::Floor($n * $TrainRatio)
        if ($n -gt 1) {
            $trainN = [Math]::Max(1, [Math]::Min($n - 1, $trainN))
        }
        for ($j = 0; $j -lt $indices.Count; $j++) {
            if ($j -lt $trainN) { [void]$train.Add([int]$indices[$j]) }
            else { [void]$test.Add([int]$indices[$j]) }
        }
    }
    return [pscustomobject]@{
        train = @($train)
        test = @($test)
        by_class = $byClass
    }
}

function Invoke-Knn {
    param(
        [object[]]$Rows,
        [System.Collections.Generic.List[object]]$Vectors,
        [int[]]$TrainIndices,
        [int[]]$TestIndices,
        [string]$TargetColumn,
        [string[]]$Classes,
        [int]$K
    )
    $preds = New-Object System.Collections.Generic.List[object]
    foreach ($testIdx in $TestIndices) {
        $testVector = $Vectors[$testIdx]
        $distances = New-Object System.Collections.Generic.List[object]
        foreach ($trainIdx in $TrainIndices) {
            $trainVector = $Vectors[$trainIdx]
            $distance = 0.0
            for ($f = 0; $f -lt $testVector.Length; $f++) {
                $delta = $testVector[$f] - $trainVector[$f]
                $distance += $delta * $delta
            }
            $distances.Add([pscustomobject]@{
                index = $trainIdx
                distance = $distance
                class = Normalize-Text $Rows[$trainIdx].$TargetColumn
            })
        }
        $neighbors = @($distances | Sort-Object distance | Select-Object -First $K)
        $votes = @{}
        $distanceSums = @{}
        foreach ($class in $Classes) {
            $votes[$class] = 0
            $distanceSums[$class] = 0.0
        }
        foreach ($n in $neighbors) {
            $votes[$n.class] = [int]$votes[$n.class] + 1
            $distanceSums[$n.class] = [double]$distanceSums[$n.class] + [double]$n.distance
        }

        $bestClass = $Classes[0]
        $bestVotes = -1
        $bestAvgDistance = [double]::PositiveInfinity
        foreach ($class in $Classes) {
            $voteCount = [int]$votes[$class]
            $avgDistance = if ($voteCount -eq 0) { [double]::PositiveInfinity } else { [double]$distanceSums[$class] / $voteCount }
            if (($voteCount -gt $bestVotes) -or (($voteCount -eq $bestVotes) -and ($avgDistance -lt $bestAvgDistance))) {
                $bestClass = $class
                $bestVotes = $voteCount
                $bestAvgDistance = $avgDistance
            }
        }

        $preds.Add([pscustomobject]@{
            row_index = $testIdx
            actual = Normalize-Text $Rows[$testIdx].$TargetColumn
            predicted = $bestClass
        })
    }
    return $preds
}

function Invoke-KnnForKValues {
    param(
        [object[]]$Rows,
        [System.Collections.Generic.List[object]]$Vectors,
        [int[]]$TrainIndices,
        [int[]]$TestIndices,
        [string]$TargetColumn,
        [string[]]$Classes,
        [int[]]$KValues
    )
    $maxK = ($KValues | Measure-Object -Maximum).Maximum
    $result = @{}
    foreach ($k in $KValues) {
        $result[$k] = New-Object System.Collections.Generic.List[object]
    }

    foreach ($testIdx in $TestIndices) {
        $testVector = $Vectors[$testIdx]
        $distances = New-Object System.Collections.Generic.List[object]
        foreach ($trainIdx in $TrainIndices) {
            $trainVector = $Vectors[$trainIdx]
            $distance = 0.0
            for ($f = 0; $f -lt $testVector.Length; $f++) {
                $delta = $testVector[$f] - $trainVector[$f]
                $distance += $delta * $delta
            }
            $distances.Add([pscustomobject]@{
                distance = $distance
                class = Normalize-Text $Rows[$trainIdx].$TargetColumn
            })
        }
        $neighbors = @($distances | Sort-Object distance | Select-Object -First $maxK)

        foreach ($k in $KValues) {
            $votes = @{}
            $distanceSums = @{}
            foreach ($class in $Classes) {
                $votes[$class] = 0
                $distanceSums[$class] = 0.0
            }
            foreach ($n in @($neighbors | Select-Object -First $k)) {
                $votes[$n.class] = [int]$votes[$n.class] + 1
                $distanceSums[$n.class] = [double]$distanceSums[$n.class] + [double]$n.distance
            }

            $bestClass = $Classes[0]
            $bestVotes = -1
            $bestAvgDistance = [double]::PositiveInfinity
            foreach ($class in $Classes) {
                $voteCount = [int]$votes[$class]
                $avgDistance = if ($voteCount -eq 0) { [double]::PositiveInfinity } else { [double]$distanceSums[$class] / $voteCount }
                if (($voteCount -gt $bestVotes) -or (($voteCount -eq $bestVotes) -and ($avgDistance -lt $bestAvgDistance))) {
                    $bestClass = $class
                    $bestVotes = $voteCount
                    $bestAvgDistance = $avgDistance
                }
            }

            $result[$k].Add([pscustomobject]@{
                row_index = $testIdx
                actual = Normalize-Text $Rows[$testIdx].$TargetColumn
                predicted = $bestClass
            })
        }
    }
    return $result
}

$targetConfig = @(
    [pscustomobject]@{ name = "target_label"; classes = @("ผ่าน", "ไม่ผ่าน") },
    [pscustomobject]@{ name = "Quality_level"; classes = @("A", "B", "C", "D") }
)
$kValues = @(3, 5, 7)

$overallMetrics = New-Object System.Collections.Generic.List[object]
$perClassMetrics = New-Object System.Collections.Generic.List[object]
$confusionMatrices = New-Object System.Collections.Generic.List[object]
$foldMetrics = New-Object System.Collections.Generic.List[object]
$metricLineData = New-Object System.Collections.Generic.List[object]

foreach ($target in $targetConfig) {
    $split = New-StratifiedSplit -Rows $rows -TargetColumn $target.name -TrainRatio 0.70 -Seed 2568
    foreach ($class in $target.classes) {
        $classIndices = @(0..($rows.Count - 1) | Where-Object { (Normalize-Text $rows[$_].($target.name)) -eq $class })
        $trainCount = @($split.train | Where-Object { (Normalize-Text $rows[$_].($target.name)) -eq $class }).Count
        $testCount = @($split.test | Where-Object { (Normalize-Text $rows[$_].($target.name)) -eq $class }).Count
        foreach ($k in $kValues) {
            $foldMetrics.Add([pscustomobject]@{
                target = $target.name
                evaluation_method = "stratified split 70:30"
                model = "KNN k=$k"
                fold = "split_test"
                class = $class
                train_count = $trainCount
                test_count = $testCount
                total_count = $classIndices.Count
                train_percent = Get-Percent $trainCount $classIndices.Count
                test_percent = Get-Percent $testCount $classIndices.Count
            })
        }
    }
    Write-Host "Training/evaluating KNN k=3,5,7 for $($target.name)..."
    $predictionsByK = Invoke-KnnForKValues -Rows $rows -Vectors $vectors -TrainIndices @($split.train) -TestIndices @($split.test) -TargetColumn $target.name -Classes $target.classes -KValues $kValues
    foreach ($k in $kValues) {
        $preds = $predictionsByK[$k]
        $metrics = Get-ClassificationMetrics -Target $target.name -Model "KNN k=$k" -Method "stratified split 70:30" -Classes $target.classes -Predictions $preds
        $overallMetrics.Add($metrics.overall)
        foreach ($x in $metrics.per_class) { $perClassMetrics.Add($x) }
        foreach ($x in $metrics.confusion) { $confusionMatrices.Add($x) }

        foreach ($metricName in @("accuracy", "macro F1", "weighted F1")) {
            $metricLineData.Add([pscustomobject]@{
                target = $target.name
                metric = $metricName
                model = "KNN k=$k"
                k = $k
                value = $metrics.overall.$metricName
            })
        }
    }
}

$overviewRows = New-Object System.Collections.Generic.List[object]
foreach ($label in @("ผ่าน", "ไม่ผ่าน")) {
    $count = @($rows | Where-Object { (Normalize-Text $_.target_label) -eq $label }).Count
    $overviewRows.Add([pscustomobject]@{
        data_type = "target_label_distribution"
        label = $label
        count = $count
        percent = Get-Percent $count $rows.Count
        feature_column = ""
        indicator_text = ""
    })
}
foreach ($label in @("A", "B", "C", "D")) {
    $count = @($rows | Where-Object { (Normalize-Text $_.Quality_level) -eq $label }).Count
    $overviewRows.Add([pscustomobject]@{
        data_type = "Quality_level_distribution"
        label = $label
        count = $count
        percent = Get-Percent $count $rows.Count
        feature_column = ""
        indicator_text = ""
    })
}

$indicatorTextByFeature = @{}
foreach ($ref in $indicatorRef) {
    $indicatorTextByFeature[$ref.feature_column] = $ref.indicator_text
}
$indicatorCounts = New-Object System.Collections.Generic.List[object]
foreach ($col in $indicatorCols) {
    $sum = 0
    foreach ($row in $rows) { $sum += [int]$row.$col }
    $indicatorCounts.Add([pscustomobject]@{
        feature_column = $col
        count = $sum
        indicator_text = if ($indicatorTextByFeature.ContainsKey($col)) { $indicatorTextByFeature[$col] } else { "" }
    })
}
$topIndicators = @($indicatorCounts | Sort-Object count -Descending | Select-Object -First 15)
foreach ($ind in $topIndicators) {
    $overviewRows.Add([pscustomobject]@{
        data_type = "top_indicator_counts"
        label = $ind.feature_column
        count = $ind.count
        percent = Get-Percent $ind.count $rows.Count
        feature_column = $ind.feature_column
        indicator_text = $ind.indicator_text
    })
}

$targetLabelCounts = @($overviewRows | Where-Object { $_.data_type -eq "target_label_distribution" })
$qualityCounts = @($overviewRows | Where-Object { $_.data_type -eq "Quality_level_distribution" })
$minorTarget = ($targetLabelCounts | Sort-Object count | Select-Object -First 1)
$majorTarget = ($targetLabelCounts | Sort-Object count -Descending | Select-Object -First 1)
$minorQuality = ($qualityCounts | Sort-Object count | Select-Object -First 1)
$majorQuality = ($qualityCounts | Sort-Object count -Descending | Select-Object -First 1)
$overviewRows.Add([pscustomobject]@{
    data_type = "class_imbalance"
    label = "target_label majority/minority ratio"
    count = if ($minorTarget.count -eq 0) { 0 } else { [Math]::Round($majorTarget.count / $minorTarget.count, 4) }
    percent = ""
    feature_column = ""
    indicator_text = "majority=$($majorTarget.label), minority=$($minorTarget.label)"
})
$overviewRows.Add([pscustomobject]@{
    data_type = "class_imbalance"
    label = "Quality_level majority/minority ratio"
    count = if ($minorQuality.count -eq 0) { 0 } else { [Math]::Round($majorQuality.count / $minorQuality.count, 4) }
    percent = ""
    feature_column = ""
    indicator_text = "majority=$($majorQuality.label), minority=$($minorQuality.label)"
})

$modelDataset = New-Object System.Collections.Generic.List[object]
foreach ($row in $rows) {
    $obj = [ordered]@{}
    foreach ($col in $categoricalCols) { $obj[$col] = Normalize-Text $row.$col }
    foreach ($col in $indicatorCols) { $obj[$col] = [int]$row.$col }
    $obj["target_label"] = Normalize-Text $row.target_label
    $obj["Quality_level"] = Normalize-Text $row.Quality_level
    $modelDataset.Add([pscustomobject]$obj)
}

$notes = @(
    [pscustomobject]@{ item = "source_workbook"; note = $InputWorkbook },
    [pscustomobject]@{ item = "source_sheet"; note = "improvement_indicators_clean" },
    [pscustomobject]@{ item = "feature_policy"; note = "รอบนี้ใช้ ministry, affiliation_1, affiliation_2, affiliation_3 และ indicator_* ทั้ง 75 columns เป็น feature" },
    [pscustomobject]@{ item = "excluded_features"; note = "ไม่ใช้ target_label, Quality_level, center_code_anon, no, center_name เป็น feature" },
    [pscustomobject]@{ item = "target_leakage_warning"; note = "Quality_level และ target_label ถูกสร้างจาก indicator_* ดังนั้นผลโมเดลรอบนี้สะท้อนความสามารถในการเรียนรู้ rule จาก feature เดียวกับที่สร้างคำตอบ ไม่ใช่การพยากรณ์จากข้อมูลอิสระ" },
    [pscustomobject]@{ item = "model"; note = "KNN ทดลอง k=3, 5, 7" },
    [pscustomobject]@{ item = "evaluation"; note = "stratified train/test split 70:30, random seed 2568" },
    [pscustomobject]@{ item = "indicator_validation"; note = "ตรวจแล้ว indicator_* ทุก column เป็น binary 0/1 จำนวน $($indicatorCols.Count) columns ไม่มีทศนิยม" },
    [pscustomobject]@{ item = "records"; note = "จำนวนระเบียนทั้งหมด $($rows.Count)" },
    [pscustomobject]@{ item = "class_imbalance_target_label"; note = "ผ่าน=$(@($rows | Where-Object { $_.target_label -eq 'ผ่าน' }).Count), ไม่ผ่าน=$(@($rows | Where-Object { $_.target_label -eq 'ไม่ผ่าน' }).Count)" },
    [pscustomobject]@{ item = "class_imbalance_Quality_level"; note = "A=$(@($rows | Where-Object { $_.Quality_level -eq 'A' }).Count), B=$(@($rows | Where-Object { $_.Quality_level -eq 'B' }).Count), C=$(@($rows | Where-Object { $_.Quality_level -eq 'C' }).Count), D=$(@($rows | Where-Object { $_.Quality_level -eq 'D' }).Count)" }
)

$dashboardRows = New-Object System.Collections.Generic.List[object]
$dashboardRows.Add([pscustomobject]@{
    section = "source"
    label = "workbook/sheet"
    value = "ml_dataset_workbook.xlsx / improvement_indicators_clean"
    percent = ""
    bar = ""
    note = "สร้างจาก feature หลัก ministry, affiliation_1, affiliation_2, affiliation_3 และ indicator_*"
})
$dashboardRows.Add([pscustomobject]@{
    section = "model_warning"
    label = "target leakage note"
    value = ""
    percent = ""
    bar = ""
    note = "Quality_level และ target_label ถูกสร้างจาก indicator_* ผลรอบนี้จึงเหมาะสำหรับตรวจ workflow/rule มากกว่าการยืนยันพลังพยากรณ์จากข้อมูลอิสระ"
})
foreach ($item in $targetLabelCounts) {
    $barLen = [int][Math]::Round(([double]$item.percent) / 2)
    $dashboardRows.Add([pscustomobject]@{
        section = "target_label_distribution"
        label = $item.label
        value = $item.count
        percent = $item.percent
        bar = ("|" * $barLen)
        note = ""
    })
}
foreach ($item in $qualityCounts) {
    $barLen = [int][Math]::Round(([double]$item.percent) / 2)
    $dashboardRows.Add([pscustomobject]@{
        section = "Quality_level_distribution"
        label = $item.label
        value = $item.count
        percent = $item.percent
        bar = ("|" * $barLen)
        note = ""
    })
}
foreach ($item in @($topIndicators | Select-Object -First 10)) {
    $barLen = [int][Math]::Round((Get-Percent $item.count $rows.Count) / 2)
    $dashboardRows.Add([pscustomobject]@{
        section = "top_indicator_counts"
        label = $item.feature_column
        value = $item.count
        percent = Get-Percent $item.count $rows.Count
        bar = ("|" * $barLen)
        note = $item.indicator_text
    })
}
foreach ($item in $overallMetrics) {
    $dashboardRows.Add([pscustomobject]@{
        section = "overall_metrics"
        label = "$($item.target) / $($item.model)"
        value = $item.accuracy
        percent = ""
        bar = ""
        note = "macro F1=$($item.'macro F1'), weighted F1=$($item.'weighted F1')"
    })
}

$sheets = @(
    [pscustomobject]@{ name = "dashboard"; rows = $dashboardRows; headers = @("section", "label", "value", "percent", "bar", "note") },
    [pscustomobject]@{ name = "model_dataset"; rows = $modelDataset; headers = @($categoricalCols + $indicatorCols + @("target_label", "Quality_level")) },
    [pscustomobject]@{ name = "overall_metrics"; rows = $overallMetrics; headers = @("target", "evaluation_method", "model", "accuracy", "macro precision", "macro recall", "macro F1", "weighted F1") },
    [pscustomobject]@{ name = "per_class_metrics"; rows = $perClassMetrics; headers = @("target", "model", "evaluation_method", "class", "precision", "recall", "F1-score", "support", "TP", "FP", "FN") },
    [pscustomobject]@{ name = "fold_metrics"; rows = $foldMetrics; headers = @("target", "evaluation_method", "model", "fold", "class", "train_count", "test_count", "total_count", "train_percent", "test_percent") },
    [pscustomobject]@{ name = "confusion_matrices"; rows = $confusionMatrices; headers = @("target", "model", "evaluation_method", "actual_class", "predicted_class", "count") },
    [pscustomobject]@{ name = "overview_visual_data"; rows = $overviewRows; headers = @("data_type", "label", "count", "percent", "feature_column", "indicator_text") },
    [pscustomobject]@{ name = "metric_line_data"; rows = $metricLineData; headers = @("target", "metric", "model", "k", "value") },
    [pscustomobject]@{ name = "model_notes"; rows = $notes; headers = @("item", "note") }
)

Write-Host "Writing OpenXML workbook..."
Export-WorkbookOpenXml -Path $OutWorkbook -Sheets $sheets
Write-Host "Saved: $OutWorkbook"

[pscustomobject]@{
    output_workbook = $OutWorkbook
    source_workbook = $InputWorkbook
    source_sheet = "improvement_indicators_clean"
    records = $rows.Count
    indicator_columns = $indicatorCols.Count
    target_label_pass = @($rows | Where-Object { $_.target_label -eq "ผ่าน" }).Count
    target_label_fail = @($rows | Where-Object { $_.target_label -eq "ไม่ผ่าน" }).Count
    quality_A = @($rows | Where-Object { $_.Quality_level -eq "A" }).Count
    quality_B = @($rows | Where-Object { $_.Quality_level -eq "B" }).Count
    quality_C = @($rows | Where-Object { $_.Quality_level -eq "C" }).Count
    quality_D = @($rows | Where-Object { $_.Quality_level -eq "D" }).Count
} | Format-List

return

$excel = $null
$wb = $null
try {
    Write-Host "Writing Excel workbook..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) {
        $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
    }

    $dashboard = $wb.Worksheets.Item(1)
    $dashboard.Name = "dashboard"
    $dashboard.Cells.Item(1, 1).Value2 = "KNN ML Dashboard: การประเมินผ่านเกณฑ์มาตรฐานสถานพัฒนาเด็กปฐมวัยแห่งชาติ"
    $dashboard.Cells.Item(1, 1).Font.Bold = $true
    $dashboard.Cells.Item(1, 1).Font.Size = 16
    $dashboard.Cells.Item(2, 1).Value2 = "Source: ml_dataset_workbook.xlsx / sheet improvement_indicators_clean"
    $dashboard.Cells.Item(3, 1).Value2 = "หมายเหตุ: Quality_level และ target_label ถูกสร้างจาก indicator_* จึงควรอ่านผลโมเดลเป็นการทดลองเรียนรู้ rule ไม่ใช่หลักฐานพยากรณ์จาก feature อิสระ"

    Write-Host "Writing dashboard summary blocks..."
    Write-Block -Worksheet $dashboard -StartRow 5 -StartCol 1 -Title "Target label distribution" -Rows $targetLabelCounts -Headers @("label", "count", "percent")
    Write-Block -Worksheet $dashboard -StartRow 5 -StartCol 6 -Title "Quality level distribution" -Rows $qualityCounts -Headers @("label", "count", "percent")
    Write-Block -Worksheet $dashboard -StartRow 15 -StartCol 1 -Title "Top indicator counts" -Rows $topIndicators -Headers @("feature_column", "count", "indicator_text")

    $bestRows = @($overallMetrics | Sort-Object target, @{ Expression = "macro F1"; Descending = $true })
    Write-Block -Worksheet $dashboard -StartRow 15 -StartCol 6 -Title "Overall metrics" -Rows $bestRows -Headers @("target", "evaluation_method", "model", "accuracy", "macro precision", "macro recall", "macro F1", "weighted F1")

    Write-Host "Creating dashboard charts..."
    $chartTargetRange = $dashboard.Range($dashboard.Cells(6, 1), $dashboard.Cells(6 + $targetLabelCounts.Count, 2))
    Add-Chart -Worksheet $dashboard -Left 20 -Top 390 -Width 330 -Height 230 -Title "Distribution target_label" -SourceRange $chartTargetRange -ChartType 51 | Out-Null
    $chartQualityRange = $dashboard.Range($dashboard.Cells(6, 6), $dashboard.Cells(6 + $qualityCounts.Count, 7))
    Add-Chart -Worksheet $dashboard -Left 370 -Top 390 -Width 330 -Height 230 -Title "Distribution Quality_level" -SourceRange $chartQualityRange -ChartType 51 | Out-Null
    $topRange = $dashboard.Range($dashboard.Cells(16, 1), $dashboard.Cells(16 + [Math]::Min(10, $topIndicators.Count), 2))
    Add-Chart -Worksheet $dashboard -Left 720 -Top 390 -Width 410 -Height 230 -Title "Top indicator_* counts" -SourceRange $topRange -ChartType 51 | Out-Null

    $modelHeaders = @($categoricalCols + $indicatorCols + @("target_label", "Quality_level"))
    Write-Host "Writing model_dataset..."
    Add-SheetData -Workbook $wb -Name "model_dataset" -Rows $modelDataset -Headers $modelHeaders | Out-Null
    Write-Host "Writing overall_metrics..."
    Add-SheetData -Workbook $wb -Name "overall_metrics" -Rows $overallMetrics -Headers @("target", "evaluation_method", "model", "accuracy", "macro precision", "macro recall", "macro F1", "weighted F1") | Out-Null
    Write-Host "Writing per_class_metrics..."
    Add-SheetData -Workbook $wb -Name "per_class_metrics" -Rows $perClassMetrics -Headers @("target", "model", "evaluation_method", "class", "precision", "recall", "F1-score", "support", "TP", "FP", "FN") | Out-Null
    Write-Host "Writing fold_metrics..."
    Add-SheetData -Workbook $wb -Name "fold_metrics" -Rows $foldMetrics -Headers @("target", "evaluation_method", "model", "fold", "class", "train_count", "test_count", "total_count", "train_percent", "test_percent") | Out-Null
    Write-Host "Writing confusion_matrices..."
    Add-SheetData -Workbook $wb -Name "confusion_matrices" -Rows $confusionMatrices -Headers @("target", "model", "evaluation_method", "actual_class", "predicted_class", "count") | Out-Null
    Write-Host "Writing overview_visual_data..."
    Add-SheetData -Workbook $wb -Name "overview_visual_data" -Rows $overviewRows -Headers @("data_type", "label", "count", "percent", "feature_column", "indicator_text") | Out-Null
    Write-Host "Writing metric_line_data..."
    Add-SheetData -Workbook $wb -Name "metric_line_data" -Rows $metricLineData -Headers @("target", "metric", "model", "k", "value") | Out-Null
    Write-Host "Writing model_notes..."
    Add-SheetData -Workbook $wb -Name "model_notes" -Rows $notes -Headers @("item", "note") | Out-Null

    $metricWs = $wb.Worksheets.Item("metric_line_data")
    $metricWs.Activate()
    $lastMetricRow = $metricLineData.Count + 1

    $dashboard.Activate()
    $wb.SaveAs($OutWorkbook, 51)
    Write-Host "Saved: $OutWorkbook"
}
finally {
    if ($wb) { $wb.Close($true) | Out-Null }
    if ($excel) { $excel.Quit() | Out-Null }
    if ($wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

[pscustomobject]@{
    output_workbook = $OutWorkbook
    source_workbook = $InputWorkbook
    source_sheet = "improvement_indicators_clean"
    records = $rows.Count
    indicator_columns = $indicatorCols.Count
    target_label_pass = @($rows | Where-Object { $_.target_label -eq "ผ่าน" }).Count
    target_label_fail = @($rows | Where-Object { $_.target_label -eq "ไม่ผ่าน" }).Count
    quality_A = @($rows | Where-Object { $_.Quality_level -eq "A" }).Count
    quality_B = @($rows | Where-Object { $_.Quality_level -eq "B" }).Count
    quality_C = @($rows | Where-Object { $_.Quality_level -eq "C" }).Count
    quality_D = @($rows | Where-Object { $_.Quality_level -eq "D" }).Count
} | Format-List
