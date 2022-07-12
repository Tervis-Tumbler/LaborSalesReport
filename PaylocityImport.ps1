function Import-PaylocityLabor100Days {
  $PaylocityProcessed = "\\$($env:USERDNSDOMAIN)\applications\Shopify\Paylocity\Processed"
  $CurrentDate = Get-Date
  $MinimumDate = $CurrentDate.AddDays(-100).ToString("yyyyMMdd")
  
  $DataFiles = Get-ChildItem -File -Path $PaylocityProcessed -Filter "LABOR_*.xml"
  $DataFiles100Days = $DataFiles | Where-Object BaseName -gt "LABOR_$MinimumDate"

  foreach ($DataFile in $DataFiles100Days) {
    Import-Clixml -Path $DataFile.FullName
  }
}
