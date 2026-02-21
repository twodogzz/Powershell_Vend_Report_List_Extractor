# ================================
# Vend Report List Extractor
# Final, robust, merged-header-safe version
# ================================

$FolderPath = "E:\OneDrive\Estland&Co\01 Business Admin\01 Accounts\02 Accounts Receivable\Vend Marketplace\Sales Reports\Final Vend Reports"
$OutputPath = "E:\OneDrive\Estland&Co\01 Business Admin\01 Accounts\02 Accounts Receivable\Vend Marketplace\Sales Reports\Vend_Report_Summary.xlsx"

# Start Excel COM
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

# Create summary workbook
$SummaryWB = $Excel.Workbooks.Add()
$SummarySheet = $SummaryWB.Sheets.Item(1)
$SummarySheet.Name = "Summary"

# Header row
$SummarySheet.Cells.Item(1,1).Value2 = "Filename"
$SummarySheet.Cells.Item(1,2).Value2 = "Earliest Date"
$SummarySheet.Cells.Item(1,3).Value2 = "Latest Date"
$SummarySheet.Cells.Item(1,4).Value2 = "Payout Amount"

$row = 2

# Logging
$Log = @()

# Safe date parser
function Convert-ToSafeDate {
    param($value)

    if ($value -eq $null -or $value -eq "") { return $null }

    if ($value -is [DateTime]) { return $value }

    if ($value -is [double]) {
        try { return [DateTime]::FromOADate($value) }
        catch { return $null }
    }

    $parsed = New-Object DateTime
    if ([DateTime]::TryParse($value.ToString(), [ref]$parsed)) {
        return $parsed
    }

    return $null
}

# Process each workbook
Get-ChildItem -Path $FolderPath -Filter *.xlsx | ForEach-Object {

    $File = $_.FullName
    Write-Host "Processing $($_.Name)..."

    try {
        $WB = $Excel.Workbooks.Open($File)
        $Sheet = $WB.Sheets.Item(1)

        $UsedRange = $Sheet.UsedRange
        $RowCount = $UsedRange.Rows.Count

        # -----------------------------------------
        # Detect header row by finding "Date" in col B
        # -----------------------------------------
        $HeaderRow = $null
        for ($r = 1; $r -le $RowCount; $r++) {
            $cellText = $Sheet.Cells.Item($r, 2).Text
            if ($cellText -eq "Date") {
                $HeaderRow = $r
                break
            }
        }

        if (-not $HeaderRow) {
            throw "Could not locate header row (missing 'Date' column header)."
        }

        $FirstDataRow = $HeaderRow + 1

        $Dates = @()
        $Payout = $null

        # -----------------------------------------
        # Read data rows only (skip merged header junk)
        # -----------------------------------------
        for ($i = $FirstDataRow; $i -le $RowCount; $i++) {

            # Column B (Date)
            $rawDate = $Sheet.Cells.Item($i, 2).Value
            $safeDate = Convert-ToSafeDate $rawDate
            if ($safeDate -is [DateTime]) {
                $Dates += $safeDate
            }

            # Column G (Payout)
            if (-not $Payout) {
                $rawPayout = $Sheet.Cells.Item($i, 7).Value2
                if ($rawPayout -is [double] -or $rawPayout -is [int] -or $rawPayout -is [decimal]) {
                    $Payout = $rawPayout
                }
            }
        }

        # Ensure only DateTime objects remain
        $Dates = $Dates | Where-Object { $_ -is [DateTime] }

        # Write results
        $SummarySheet.Cells.Item($row, 1).Value2 = $_.Name
        $SummarySheet.Cells.Item($row, 2).Value2 = ($Dates | Measure-Object -Minimum).Minimum
        $SummarySheet.Cells.Item($row, 3).Value2 = ($Dates | Measure-Object -Maximum).Maximum
        $SummarySheet.Cells.Item($row, 4).Value2 = $Payout

        $row++

        $WB.Close()

    } catch {
        # Safe logging
        $Log += ("ERROR processing {0}: {1}" -f $_.Name, $_.Exception.Message.ToString())
        Write-Host "Error in $($_.Name)"
    }
}

# Save summary
$SummaryWB.SaveAs($OutputPath)
$SummaryWB.Close()
$Excel.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($SummarySheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($SummaryWB) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "Summary created at $OutputPath"

# Optional: Write log file
if ($Log.Count -gt 0) {
    $Log | Out-File -FilePath "$FolderPath\VendReportExtractor_Errors.log"
    Write-Host "Errors logged to VendReportExtractor_Errors.log"
}