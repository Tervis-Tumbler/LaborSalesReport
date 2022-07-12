function Import-ShopperTrakSales100Days {
  $ShopperTrakDataPath = "\\$($env:USERDNSDOMAIN)\applications\Shopify\ShopperTrak\Outbound"
  $SubstitutionTable = Import-Csv -Path "\\$($env:USERDNSDOMAIN)\applications\Shopify\ShopperTrak\Config\Paylocity_Shopify_StoreIDs.csv"
  $CurrentDate = Get-Date
  $MinimumDate = $CurrentDate.AddDays(-100).ToString("yyyyMMdd")
  
  $DataFiles = Get-ChildItem -File -Path $ShopperTrakDataPath -Filter "SALES_*.txt"
  $DataFiles100Days = $DataFiles | Where-Object BaseName -gt "SALES_$MinimumDate"
  
  # $i = 0
  $j = -1
  
  foreach ($DataFile in $DataFiles100Days) {
    # $i++; Write-Host $i
    $Import = Import-Csv `
      -Path $DataFile.FullName `
      -Header "Location","Date","Time","OrderId","Total","Quantity"
    $Import | ForEach-Object {
      $j++
      $DateString = $_.Date.ToString() + $_.Time.ToString()
      $Date = [DateTime]::ParseExact($DateString, "yyyyMMddHHmmss", $null)
      $LocationName = $SubstitutionTable |
        Where-Object StoreID -eq $_.Location |
        Select-Object -ExpandProperty PaylocityName
      # The following line is temporary
      if (-not $LocationName) { $LocationName = "BRANSON-TANGER2"}
      [PSCustomObject]@{
        Index = $j
        OrderId = $_.OrderId
        Date = $Date
        Location = $LocationName
        Total = $_.Total
        Quantity = $_.Quantity
      }
    }
  }
}
