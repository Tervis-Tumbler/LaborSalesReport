#requires -modules ImportExcel

param (
  [string]$ReportDate,
  [string]$ProductionYearStartDate = "1/2/2022",
  [string]$ExportPath = "\\$($env:USERDNSDOMAIN)\applications\Shopify\Paylocity\Reports"
)

# Import helper functions
. $PSScriptRoot\PaylocityConvert.ps1
. $PSScriptRoot\PaylocityImport.ps1
. $PSScriptRoot\ShopperTrakImport.ps1
. $PSScriptRoot\Calendar.ps1
. $PSScriptRoot\MailSettings.ps1

# Convert and import data
Write-Progress -Activity "Labor Sales Report" -CurrentOperation "Importing data"
Invoke-PaylocityLaborExportConversion
$PaylocityData = Import-PaylocityLabor100Days
$ShopperTrakData = Import-ShopperTrakSales100Days
$Calendar = Get-LaborSalesCalendar -ProductionYearStartDate "1/2/2022"

# Add date metadata to ShopperTrak
$i = 0
$Count = $ShopperTrakData.Count
$ShopperTrakDataWithDateIndices = $ShopperTrakData | ForEach-Object {
  if ($i % 1000 -eq 0) {Write-Progress -Activity "Labor Sales Report" -CurrentOperation "Adding ShopperTrak Date Indices" -Status "$i of $Count" -PercentComplete ($i * 100 / $Count)}
  $i++
  $WeekIndex = Get-IndexOfDateFromList -Date $_.Date -DateList $Calendar.WeekStartDates
  $MonthIndex = Get-IndexOfDateFromList -Date $_.Date -DateList $Calendar.MonthStartDates
  $QuarterIndex = Get-IndexOfDateFromList -Date $_.Date -DateList $Calendar.QuarterStartDates
  
  $_ |
  Add-Member -MemberType NoteProperty -Name WeekIndex -Value $WeekIndex -PassThru -Force |
  Add-Member -MemberType NoteProperty -Name MonthIndex -Value $MonthIndex -PassThru -Force |
  Add-Member -MemberType NoteProperty -Name QuarterIndex -Value $QuarterIndex -PassThru -Force
}

# Add date metadata to Paylocity
$i = 0
$Count = $PaylocityData.Count
$PaylocityDataWithDateIndices = $PaylocityData | ForEach-Object {
  if ($i % 1000 -eq 0) {Write-Progress -Activity "Labor Sales Report" -CurrentOperation "Adding Paylocity Date Indices" -Status "$i of $Count" -PercentComplete ($i * 100 / $Count)}
  $i++
  $WeekIndex = Get-IndexOfDateFromList -Date $_.StartDate -DateList $Calendar.WeekStartDates
  $MonthIndex = Get-IndexOfDateFromList -Date $_.StartDate -DateList $Calendar.MonthStartDates
  $QuarterIndex = Get-IndexOfDateFromList -Date $_.StartDate -DateList $Calendar.QuarterStartDates

  $_ |
  Add-Member -MemberType NoteProperty -Name WeekIndex -Value $WeekIndex -PassThru -Force |
  Add-Member -MemberType NoteProperty -Name MonthIndex -Value $MonthIndex -PassThru -Force |
  Add-Member -MemberType NoteProperty -Name QuarterIndex -Value $QuarterIndex -PassThru -Force
}

Write-Progress -Activity "Labor Sales Report" -CurrentOperation "Calculating Dates"

# Get calendar dates
$CurrentDate = (Get-Date).AddDays(-7)
if ($ReportDate) { $CurrentDate = Get-Date $ReportDate }
$CurrentWeekIndex = 
  Get-IndexOfDateFromList -Date $CurrentDate -DateList $Calendar.WeekStartDates
$CurrentMonthIndex = 
  Get-IndexOfDateFromList -Date $CurrentDate -DateList $Calendar.MonthStartDates
$CurrentQuarterIndex = 
  Get-IndexOfDateFromList -Date $CurrentDate -DateList $Calendar.QuarterStartDates

# Calculate *-To-Date values for each store
$Timespans = "Week","Month","Quarter"
$DateAggregrations = foreach ($Timespan in $Timespans) {
  Write-Progress -Activity "Labor Sales Report" -CurrentOperation "Calculating $Timespan-To-Date Values"
  switch ($Timespan) {
    "Week"    {
      $CurrentTimespanIndex = $CurrentWeekIndex
      $TimespanIndexProp = "WeekIndex"
      $TimespanAcronym = "WTD"
      break
    }
    "Month"   {
      $CurrentTimespanIndex = $CurrentMonthIndex
      $TimespanIndexProp = "MonthIndex"
      $TimespanAcronym = "MTD"
      break
    }
    "Quarter" {
      $CurrentTimespanIndex = $CurrentQuarterIndex
      $TimespanIndexProp = "QuarterIndex"
      $TimespanAcronym = "QTD"
      break
    }
    Default {}
  }
  
  $CurrentTimespanLabor = $PaylocityDataWithDateIndices | 
    Where-Object $TimespanIndexProp -eq $CurrentTimespanIndex |
    Where-Object WeekIndex -le $CurrentWeekIndex |
    Group-Object Location
  
  foreach ($Labor in $CurrentTimespanLabor) {
    $LaborLocation = $Labor.Name
    $OperationHours = $ShopperTrakDataWithDateIndices |
      Where-Object Location -eq $LaborLocation |
      Select-Object -First 1 |
      Select-Object -ExpandProperty OperationHours

    if (-not $OperationHours) { continue }
    $LaborHours = $Labor.Group |
      Measure-Object -Property TotalHours -Sum |
      Select-Object -ExpandProperty Sum
    $LaborPay = $Labor.Group |
      Measure-Object -Property TotalPay -Sum |
      Select-Object -ExpandProperty Sum
  
    $TimespanTotalSales = $ShopperTrakDataWithDateIndices |
      Where-Object $TimespanIndexProp -eq $CurrentTimespanIndex |
      Where-Object WeekIndex -le $CurrentWeekIndex |
      Where-Object Location -eq $LaborLocation |
      Measure-Object -Property Total -Sum |
      Select-Object -ExpandProperty Sum
    
    $SalesConversion = $LaborPay / $TimespanTotalSales
  
    [PSCustomObject]@{
      Store = $LaborLocation,$OperationHours -join "_"
      "$TimespanAcronym`_Hours" = [Math]::Round($LaborHours, 2)
      "$TimespanAcronym`_Wages" = [Math]::Round($LaborPay, 2)
      "$TimespanAcronym`_Payroll%" = [Math]::Round($SalesConversion, 4)
      "$TimespanAcronym`_Sales" = [Math]::Round($TimespanTotalSales, 2)
    }
  }
}

# Join data by store
Write-Progress -Activity "Labor Sales Report" -CurrentOperation "Aggregating data by store"
$JoinedByStore = $DateAggregrations | Group-Object -Property Store | ForEach-Object {
  $StoreName = ($_.Name -split "_")[0]
  $OperationHours = ($_.Name -split "_")[1]
  $Hashtable = @{ Store = $StoreName; OperationHours = $OperationHours }
  $_.Group | ForEach-Object {
    $_.PSObject.Properties |
      Where-Object {
        $_.MemberType -eq "NoteProperty" -and
        $_.Name -ne "Store"
      } |
      ForEach-Object {
        $Hashtable += (@{ "$($_.Name)" = $_.Value })
      }
  }
  New-Object -TypeName psobject -Property $Hashtable
}

# Calculate totals
$TotalsRow = [PSCustomObject]@{
  Store = "Totals"
}

$TotalsFields = $JoinedByStore[0].PSObject.Properties |
  Where-Object {
    $_.MemberType -eq "NoteProperty" -and
    $_.Name -ne "Store" -and
    $_.Name -notlike "*Payroll*"
  } | Select-Object -ExpandProperty "Name"

foreach ($Field in $TotalsFields) {
  $Total = $JoinedByStore | Measure-Object -Property $Field -Sum | Select-Object -ExpandProperty Sum
  $TotalsRow | Add-Member -NotePropertyName $Field -NotePropertyValue $Total
}

$JoinedByStore += $TotalsRow

# Export report
Write-Progress -Activity "Labor Sales Report" -CurrentOperation "Exporting report"
$ReportDate = $Calendar.WeekStartDates[$CurrentWeekIndex].ToString("yyyyMMdd")
# $ReportTitle = "LaborSalesReport_$ReportDate.csv"
$ReportXlsx = "LaborSalesReport_$ReportDate.xlsx"
Remove-Item -Path $ExportPath\$ReportXlsx -Force -ErrorAction Continue
$xl = $JoinedByStore | Select-Object -Property `
  Store,
  OperationHours,
  WTD_Hours,
  WTD_Wages,
  "WTD_Payroll%",
  WTD_Sales,
  MTD_Hours,
  MTD_Wages,
  "MTD_Payroll%",
  MTD_Sales,
  QTD_Hours,
  QTD_Wages,
  "QTD_Payroll%",
  QTD_Sales |
  Export-Excel -Path $ExportPath\$ReportXlsx -AutoSize -PassThru
  
$Sheet = $xl.Sheet1
$Sheet.InsertRow(1,1)

# Set Date, WTD,MTD,QTD headers
# $Sheet.Cells["A1"].Value = "Week of $($Calendar.WeekStartDates[$CurrentWeekIndex].ToString("MM/dd/yyyy"))"
Set-ExcelRange -Address $Sheet.Cells["A1:B1"] -Merge -Value "Week of $($Calendar.WeekStartDates[$CurrentWeekIndex].ToString("MM/dd/yyyy"))"
Set-ExcelRange -Address $Sheet.Cells["C1:F1"] -Merge -Value "WTD" -Bold -Underline -HorizontalAlignment "Center"
Set-ExcelRange -Address $Sheet.Cells["G1:J1"] -Merge -Value "MTD" -Bold -Underline -HorizontalAlignment "Center"
Set-ExcelRange -Address $Sheet.Cells["K1:N1"] -Merge -Value "QTD" -Bold -Underline -HorizontalAlignment "Center"

# Set column headers
$Columns = "A","B","C","D","E","F","G","H","I","J","K","L","M","N"
$Headings = "Store","Hours of Operation","Hours","Wages","Payroll %","Sales","Hours","Wages","Payroll %","Sales","Hours","Wages","Payroll %","Sales"
$Columns | ForEach-Object {
  $Shift,$Headings = $Headings
  $Sheet.Cells["$_`2"].Value = $Shift
}
Set-ExcelRange -Address $Sheet.Cells["A2:N2"] -HorizontalAlignment "Center" -Bold 

# Format cells
Set-ExcelRange -Address $Sheet.Cells["B1:B$($Sheet.Dimension.Rows)"] -AutoSize -HorizontalAlignment "Center"
Set-ExcelRange -Address $Sheet.Cells["D1:F$($Sheet.Dimension.Rows)"] -AutoSize -NumberFormat "Currency"
Set-ExcelRange -Address $Sheet.Cells["H1:J$($Sheet.Dimension.Rows)"] -AutoSize -NumberFormat "Currency"
Set-ExcelRange -Address $Sheet.Cells["L1:N$($Sheet.Dimension.Rows)"] -AutoSize -NumberFormat "Currency"
Set-ExcelRange -Address $Sheet.Cells["E1:E$($Sheet.Dimension.Rows)"] -AutoSize -NumberFormat "Percentage"
Set-ExcelRange -Address $Sheet.Cells["I1:I$($Sheet.Dimension.Rows)"] -AutoSize -NumberFormat "Percentage"
Set-ExcelRange -Address $Sheet.Cells["M1:M$($Sheet.Dimension.Rows)"] -AutoSize -NumberFormat "Percentage"
Set-ExcelRange -Address $Sheet.Cells["C1:F$($Sheet.Dimension.Rows)"] -BorderAround "Medium"
Set-ExcelRange -Address $Sheet.Cells["G1:J$($Sheet.Dimension.Rows)"] -BorderAround "Medium"
Set-ExcelRange -Address $Sheet.Cells["K1:N$($Sheet.Dimension.Rows)"] -BorderAround "Medium"
Set-ExcelRange -Address $Sheet.Cells["A2:N$($Sheet.Dimension.Rows)"] -PatternColor "Blue"

for ($i = 2; $i -lt $Sheet.Dimension.Rows; $i++) {
  if ($i % 2 -eq 0) {
    $RowColor = "LightBlue"
  } else {
    $RowColor = "LightGray"
  }
  Set-ExcelRow -Worksheet $Sheet -Row $i -BackgroundColor $RowColor
}

# Add formulas for pecentages
Set-ExcelRange -Address $Sheet.Cells["E$($Sheet.Dimension.Rows)"] -Formula "=D$($Sheet.Dimension.Rows)/F$($Sheet.Dimension.Rows)"
Set-ExcelRange -Address $Sheet.Cells["I$($Sheet.Dimension.Rows)"] -Formula "=H$($Sheet.Dimension.Rows)/J$($Sheet.Dimension.Rows)"
Set-ExcelRange -Address $Sheet.Cells["M$($Sheet.Dimension.Rows)"] -Formula "=L$($Sheet.Dimension.Rows)/N$($Sheet.Dimension.Rows)"

# Save Excel file
$xl | Close-ExcelPackage
Write-Progress -Activity "Labor Sales Report" -Completed

# Email Report
$MailDateString = $Calendar.WeekStartDates[$CurrentWeekIndex].ToString("d")
$MailTitle = "Labor Sales Report for week of $MailDateString"
Send-MailMessage @MailParams -Subject $MailTitle -Attachments $ExportPath\$ReportXlsx