param(
    [Parameter(Mandatory=$true)]
    [string]$TestDir
)

$ErrorActionPreference = 'Stop'
$excel = $null

function Release-ComObject($value) {
    if ($null -ne $value) {
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($value) | Out-Null } catch { }
    }
}

function Normalize-Value($value) {
    if ($null -eq $value) { return '' }
    return [string]$value
}

function Compare-Cell($sourceCell, $targetCell, [string]$location) {
    $sourceValue = $sourceCell.Value2
    $targetValue = $targetCell.Value2

    if ($sourceValue -is [double] -and $targetValue -is [double]) {
        if ([math]::Abs($sourceValue - $targetValue) -gt 0.00000001) {
            throw "Value mismatch at ${location}: $sourceValue vs $targetValue"
        }
    } elseif ((Normalize-Value $sourceValue) -ne (Normalize-Value $targetValue)) {
        throw "Value mismatch at ${location}: '$(Normalize-Value $sourceValue)' vs '$(Normalize-Value $targetValue)'"
    }

    $sourceFormat = [string]$sourceCell.NumberFormat
    $targetFormat = [string]$targetCell.NumberFormat
    if ($sourceFormat -ne $targetFormat) {
        throw "Number format mismatch at ${location}: '$sourceFormat' vs '$targetFormat'"
    }
}

function Validate-Pair(
    [string]$sourcePath,
    [string]$targetPath,
    [int]$expectedFormat,
    [string]$label
) {
    $sourceWorkbook = $null
    $targetWorkbook = $null

    try {
        $sourceWorkbook = $excel.Workbooks.Open($sourcePath, 0, $true)
        $targetWorkbook = $excel.Workbooks.Open($targetPath, 0, $true)

        if ([int]$targetWorkbook.FileFormat -ne $expectedFormat) {
            throw "$label has unexpected FileFormat $($targetWorkbook.FileFormat), expected $expectedFormat"
        }
        if ($sourceWorkbook.Worksheets.Count -ne $targetWorkbook.Worksheets.Count) {
            throw "$label changed worksheet count"
        }

        for ($sheetIndex = 1; $sheetIndex -le $sourceWorkbook.Worksheets.Count; $sheetIndex++) {
            $sourceSheet = $sourceWorkbook.Worksheets.Item($sheetIndex)
            $targetSheet = $targetWorkbook.Worksheets.Item($sheetIndex)
            if ($sourceSheet.Name -ne $targetSheet.Name) {
                throw "$label changed worksheet name at index $sheetIndex"
            }

            $sourceUsed = $sourceSheet.UsedRange
            $targetUsed = $targetSheet.UsedRange
            $sourceRows = [int]$sourceUsed.Rows.Count
            $targetRows = [int]$targetUsed.Rows.Count
            $sourceColumns = [int]$sourceUsed.Columns.Count
            $targetColumns = [int]$targetUsed.Columns.Count
            if ($sourceRows -ne $targetRows -or $sourceColumns -ne $targetColumns) {
                throw "$label changed used range on '$($sourceSheet.Name)': ${sourceRows}x${sourceColumns} vs ${targetRows}x${targetColumns}"
            }

            for ($row = 1; $row -le $sourceRows; $row++) {
                for ($column = 1; $column -le $sourceColumns; $column++) {
                    $sourceCell = $sourceSheet.Cells.Item($row, $column)
                    $targetCell = $targetSheet.Cells.Item($row, $column)
                    Compare-Cell $sourceCell $targetCell "$($sourceSheet.Name)!R${row}C${column}"
                }
            }
        }

        Write-Output "PASS|$label|$([int]$targetWorkbook.FileFormat)|$($targetWorkbook.Worksheets.Count) sheets"
    } catch {
        Write-Output "FAIL|$label|$($_.Exception.Message -replace '[`r`n]', ' ')"
        throw
    } finally {
        if ($null -ne $targetWorkbook) {
            try { $targetWorkbook.Close($false) } catch { }
            Release-ComObject $targetWorkbook
        }
        if ($null -ne $sourceWorkbook) {
            try { $sourceWorkbook.Close($false) } catch { }
            Release-ComObject $sourceWorkbook
        }
    }
}

try {
    try {
        $excel = New-Object -ComObject Excel.Application
    } catch {
        Write-Output "SKIP|Excel COM unavailable|$($_.Exception.Message -replace '[`r`n]', ' ')"
        exit 0
    }

    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    Validate-Pair (Join-Path $TestDir 'source.xlsb') (Join-Path $TestDir 'converted-from-xlsb.xlsx') 51 'XLSB-to-XLSX'
    Validate-Pair (Join-Path $TestDir 'source.xlsx') (Join-Path $TestDir 'converted-from-xlsx.xlsb') 50 'XLSX-to-XLSB'
} catch {
    exit 1
} finally {
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
