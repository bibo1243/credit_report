param(
    [string]$ExcelPath,
    [string]$ExcelDirectory = "E:\OneDrive - tkcy\AI",
    [string]$JsonPath = (Join-Path $PSScriptRoot "goods.json"),
    [string]$JsPath = (Join-Path $PSScriptRoot "goods-data.js")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Decode-Utf8Base64 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    return [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($Value))
}

function Write-Utf8NoBom {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

function Get-LatestExcelFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory
    )

    $file = Get-ChildItem -LiteralPath $Directory -File -Filter "*.xlsx" |
        Where-Object { $_.Name -notlike '~$*' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $file) {
        throw "No Excel file (*.xlsx) was found in the source folder."
    }

    return $file.FullName
}

function Resolve-ShortcutTarget {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    $shell = $null
    $shortcut = $null

    try {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($Path)
        return $shortcut.TargetPath
    }
    finally {
        foreach ($comObject in @($shortcut, $shell)) {
            if ($null -ne $comObject) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject)
            }
        }
    }
}

function Get-LatestShortcutTarget {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory
    )

    $shortcuts = Get-ChildItem -LiteralPath $Directory -File -Filter "*.lnk" |
        Sort-Object LastWriteTime -Descending

    foreach ($shortcut in $shortcuts) {
        $targetPath = Resolve-ShortcutTarget -Path $shortcut.FullName
        if ($targetPath -and [System.IO.Path]::GetExtension($targetPath).ToLowerInvariant() -eq ".xlsx") {
            return $targetPath
        }
    }

    return $null
}

function Normalize-Header {
    param(
        [string]$Value
    )

    if ($null -eq $Value) {
        return ""
    }

    return ([string]$Value) -replace "[\s　]", ""
}

function Find-HeaderMap {
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,
        [Parameter(Mandatory = $true)]
        [string[]]$RequiredHeaders
    )

    for ($sheetIndex = 1; $sheetIndex -le $Workbook.Worksheets.Count; $sheetIndex++) {
        $worksheet = $Workbook.Worksheets.Item($sheetIndex)
        $usedRange = $worksheet.UsedRange
        $usedRows = $usedRange.Rows.Count
        $usedCols = $usedRange.Columns.Count
        $values = $usedRange.Value2
        $maxHeaderRow = [Math]::Min(10, $usedRows)

        for ($row = 1; $row -le $maxHeaderRow; $row++) {
            $map = @{}

            for ($col = 1; $col -le $usedCols; $col++) {
                $header = Normalize-Header $values[$row, $col]
                if ($RequiredHeaders -contains $header -and -not $map.ContainsKey($header)) {
                    $map[$header] = $col
                }
            }

            $allMatched = $true
            foreach ($headerName in $RequiredHeaders) {
                if (-not $map.ContainsKey($headerName)) {
                    $allMatched = $false
                    break
                }
            }

            if ($allMatched) {
                return @{
                    Worksheet = $worksheet
                    HeaderRow = $row
                    ColumnMap = $map
                    UsedRows = $usedRows
                    Values = $values
                }
            }
        }
    }

    throw "Required headers were not found in the workbook."
}

function Normalize-Quantity {
    param(
        $Value
    )

    if ($null -eq $Value) {
        return ""
    }

    try {
        $number = [double]$Value
        if ([Math]::Abs($number - [Math]::Round($number, 0)) -lt 0.000001) {
            return [string][int][Math]::Round($number, 0)
        }
        return [string]$number
    }
    catch {
        return ([string]$Value).Trim()
    }
}

$dateKey = Decode-Utf8Base64 "5pel5pyf"
$nameKey = Decode-Utf8Base64 "5aSn5b635aeT5ZCN"
$itemKey = Decode-Utf8Base64 "54mp5ZOB5ZCN56ix"
$quantityKey = Decode-Utf8Base64 "5pW46YeP"
$unitKey = Decode-Utf8Base64 "5Zau5L2N"
$recipientKey = Decode-Utf8Base64 "5Y+X6LSI5Zau5L2N"
$requiredHeaders = @($dateKey, $nameKey, $itemKey, $quantityKey, $unitKey, $recipientKey)

if (-not $ExcelPath) {
    $shortcutTarget = Get-LatestShortcutTarget -Directory $ExcelDirectory
    if ($shortcutTarget -and (Test-Path -LiteralPath $shortcutTarget)) {
        $ExcelPath = $shortcutTarget
    }
    else {
        $ExcelPath = Get-LatestExcelFile -Directory $ExcelDirectory
    }
}

if (-not (Test-Path -LiteralPath $ExcelPath)) {
    throw "Excel file not found: $ExcelPath"
}

$excel = $null
$workbook = $null
$worksheet = $null

try {
    Write-Host "Updating goods files from Excel..." -ForegroundColor Cyan
    Write-Host "Source: $ExcelPath" -ForegroundColor DarkCyan

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($ExcelPath, 0, $true)
    $headerInfo = Find-HeaderMap -Workbook $workbook -RequiredHeaders $requiredHeaders
    $worksheet = $headerInfo.Worksheet
    $headerRow = $headerInfo.HeaderRow
    $columnMap = $headerInfo.ColumnMap
    $usedRows = $headerInfo.UsedRows
    $values = $headerInfo.Values

    $rows = New-Object System.Collections.Generic.List[object]

    for ($row = $headerRow + 1; $row -le $usedRows; $row++) {
        $date = [string]$values[$row, $columnMap[$dateKey]]
        $name = [string]$values[$row, $columnMap[$nameKey]]
        $itemName = [string]$values[$row, $columnMap[$itemKey]]
        $quantity = Normalize-Quantity $values[$row, $columnMap[$quantityKey]]
        $unit = [string]$values[$row, $columnMap[$unitKey]]
        $recipient = [string]$values[$row, $columnMap[$recipientKey]]

        if ([string]::IsNullOrWhiteSpace($date) -and
            [string]::IsNullOrWhiteSpace($name) -and
            [string]::IsNullOrWhiteSpace($itemName) -and
            [string]::IsNullOrWhiteSpace($quantity) -and
            [string]::IsNullOrWhiteSpace($unit) -and
            [string]::IsNullOrWhiteSpace($recipient)) {
            continue
        }

        if ([string]::IsNullOrWhiteSpace($name) -and [string]::IsNullOrWhiteSpace($itemName)) {
            continue
        }

        $rows.Add([ordered]@{
            $dateKey = $date.Trim()
            $nameKey = $name.Trim()
            $itemKey = $itemName.Trim()
            $quantityKey = $quantity
            $unitKey = $unit.Trim()
            $recipientKey = $recipient.Trim()
        })
    }

    $json = $rows | ConvertTo-Json -Depth 3 -Compress
    $js = "window.__GOODS_DATA__ = $json;"

    Write-Utf8NoBom -Path $JsonPath -Content $json
    Write-Utf8NoBom -Path $JsPath -Content $js

    Write-Host ""
    Write-Host "Update complete." -ForegroundColor Green
    Write-Host ("Worksheet: {0}" -f $worksheet.Name)
    Write-Host ("Rows exported: {0}" -f $rows.Count)
    Write-Host ("JSON: {0}" -f $JsonPath)
    Write-Host ("JS: {0}" -f $JsPath)
}
finally {
    if ($workbook) {
        $workbook.Close($false) | Out-Null
    }

    if ($excel) {
        $excel.Quit()
    }

    foreach ($comObject in @($worksheet, $workbook, $excel)) {
        if ($null -ne $comObject) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject)
        }
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
