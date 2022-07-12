function Invoke-PaylocityLaborExportConversion {
  $PaylocityBasePath = "\\$($env:USERDNSDOMAIN)\applications\Shopify\Paylocity"
  $PaylocityInbound = "$PaylocityBasePath\Inbound"
  $PaylocityProcessed = "$PaylocityBasePath\Processed"
  $PaylocityArchive = "$PaylocityBasePath\Archive"
  
  $FilesToProcess = Get-ChildItem -File -Path $PaylocityInbound -Filter "*.csv"
  
  foreach ($File in $FilesToProcess) {
    $ImportOptions = @{
      # Path = "$PaylocityBasePath\Inbound\Stores Weekly Payroll report (B7234).csv"
      Path = $File.FullName
      Header = "Location","TotalHours","NA","TotalPay"
    }
    
    $LaborCsv = Import-Csv @ImportOptions
    
    $Dates = $LaborCsv[1].Location -split " - "
    $StartDate = Get-Date $Dates[0]
    $EndDate = Get-Date $Dates[1]
    
    $StartObject = $LaborCsv | Where-Object {$_.Location -eq "Location"}
    
    $HeaderIndex = $LaborCsv.IndexOf($StartObject)
    
    $NumberStyles = (
      [System.Globalization.NumberStyles]::AllowThousands -bor
      [System.Globalization.NumberStyles]::AllowDecimalPoint -bor
      [System.Globalization.NumberStyles]::AllowCurrencySymbol
    )
    
    $PaylocityData = $LaborCsv | Select-Object  -Skip ($HeaderIndex + 1) | ForEach-Object {
      $TotalHours = [System.Decimal]::Parse($_.TotalHours, $NumberStyles)
      $TotalPay = [System.Decimal]::Parse($_.TotalPay, $NumberStyles)
      [PSCustomObject]@{
        Location = $_.Location
        TotalHours = $TotalHours
        TotalPay = $TotalPay
        StartDate = $StartDate
        EndDate = $EndDate
      }
    }
  
    $Filename = "LABOR_$($StartDate.toString("yyyyMMdd"))"
    $PaylocityData | Export-Clixml -Path "$PaylocityProcessed\$Filename.xml" -Force
    Move-Item -Path $File.FullName -Destination "$PaylocityArchive\$Filename.csv" -Force
  }
}
