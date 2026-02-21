$WB = (New-Object -ComObject Excel.Application).Workbooks.Open("E:\OneDrive\Estland&Co\01 Business Admin\01 Accounts\02 Accounts Receivable\Vend Marketplace\Sales Reports\Final Vend Reports\2-Vend Virginia Shop #3 2024-10-20.xlsx")
$Sheet = $WB.Sheets.Item(1)

foreach ($r in 1..6) {
    foreach ($c in 1..7) {
        $cell = $Sheet.Cells.Item($r,$c)
        Write-Host "$r,$c | Text='$($cell.Text)' | Value='$($cell.Value)' | Format='$($cell.NumberFormat)' | MergeArea=$($cell.MergeArea.Address)"
    }
}