param(
    [string]$ExcelPath,
    [string]$JsonPath = (Join-Path $PSScriptRoot "donations.json"),
    [string]$JsPath = (Join-Path $PSScriptRoot "donations-data.js")
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
        throw "No Excel file (*.xlsx) was found in this folder."
    }

    return $file.FullName
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
                $header = [string]$values[$row, $col]
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

function Parse-Amount {
    param(
        $Value,
        [string]$Text
    )

    if ($null -ne $Value) {
        try {
            return [int][Math]::Round([double]$Value, 0)
        }
        catch {
        }
    }

    $cleaned = (($Text | Out-String) -replace "[,\s]", "").Trim()
    if ([string]::IsNullOrWhiteSpace($cleaned)) {
        return 0
    }

    try {
        return [int][Math]::Round([double]::Parse($cleaned), 0)
    }
    catch {
        return 0
    }
}

$dateKey = Decode-Utf8Base64 "5pel5pyf"
$nameKey = Decode-Utf8Base64 "5aSn5b635aeT5ZCN"
$amountKey = Decode-Utf8Base64 "6YeR6aGN"
$unitKey = Decode-Utf8Base64 "5Y+X6LSI5Zau5L2N"
$purposeKey = Decode-Utf8Base64 "5oyH5a6a55So6YCU"
$requiredHeaders = @($dateKey, $nameKey, $amountKey, $unitKey, $purposeKey)

if (-not $ExcelPath) {
    $ExcelPath = Get-LatestExcelFile -Directory $PSScriptRoot
}

if (-not (Test-Path -LiteralPath $ExcelPath)) {
    throw "Excel file not found: $ExcelPath"
}

$excel = $null
$workbook = $null
$worksheet = $null

try {
    Write-Host "Updating donation files from Excel..." -ForegroundColor Cyan
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
        $unit = [string]$values[$row, $columnMap[$unitKey]]
        $purpose = [string]$values[$row, $columnMap[$purposeKey]]
        $amountValue = $values[$row, $columnMap[$amountKey]]
        $amount = Parse-Amount -Value $amountValue -Text ([string]$amountValue)

        if ([string]::IsNullOrWhiteSpace($date) -and
            [string]::IsNullOrWhiteSpace($name) -and
            [string]::IsNullOrWhiteSpace($unit) -and
            [string]::IsNullOrWhiteSpace($purpose) -and
            $amount -eq 0) {
            continue
        }

        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        $rows.Add([ordered]@{
            $dateKey = $date.Trim()
            $nameKey = $name.Trim()
            $amountKey = $amount
            $unitKey = $unit.Trim()
            $purposeKey = $purpose.Trim()
        })
    }

    $json = $rows | ConvertTo-Json -Depth 3 -Compress
    $js = "window.__DONATIONS_DATA__ = $json;"

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
