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
      Store = $LaborLocation
      "$TimespanAcronym`_Hours" = [Math]::Round($LaborHours)
      "$TimespanAcronym`_Wages" = [Math]::Round($LaborPay, 2)
      "$TimespanAcronym`_Payroll%" = [Math]::Round($SalesConversion, 2)
      "$TimespanAcronym`_Sales" = [Math]::Round($TimespanTotalSales, 2)
    }
  }
}

# Join data by store
Write-Progress -Activity "Labor Sales Report" -CurrentOperation "Aggregating data by store"
$JoinedByStore = $DateAggregrations | Group-Object -Property Store | ForEach-Object {
  $Hashtable = @{ Store = $_.Name }
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
$Sheet.Cells["A1"].Value = "Week of $($Calendar.WeekStartDates[$CurrentWeekIndex].ToString("MM/dd/yyyy"))"
Set-ExcelRange -Address $Sheet.Cells["B1:E1"] -Merge -Value "WTD" -Bold -Underline -HorizontalAlignment "Center"
Set-ExcelRange -Address $Sheet.Cells["F1:I1"] -Merge -Value "MTD" -Bold -Underline -HorizontalAlignment "Center"
Set-ExcelRange -Address $Sheet.Cells["J1:M1"] -Merge -Value "QTD" -Bold -Underline -HorizontalAlignment "Center"

# Set column headers
$Columns = "A","B","C","D","E","F","G","H","I","J","K","L","M"
$Headings = "Store","Hours","Wages","Payroll %","Sales","Hours","Wages","Payroll %","Sales","Hours","Wages","Payroll %","Sales"
$Columns | ForEach-Object {
  $Shift,$Headings = $Headings
  $Sheet.Cells["$_`2"].Value = $Shift
}
Set-ExcelRange -Address $Sheet.Cells["A2:M2"] -HorizontalAlignment "Center" -Bold 

# Format cells
Set-ExcelRange -Address $Sheet.Cells["C1:E$($Sheet.Dimension.Rows)"] -AutoSize -NumberFormat "Currency"
Set-ExcelRange -Address $Sheet.Cells["G1:I$($Sheet.Dimension.Rows)"] -AutoSize -NumberFormat "Currency"
Set-ExcelRange -Address $Sheet.Cells["K1:M$($Sheet.Dimension.Rows)"] -AutoSize -NumberFormat "Currency"
Set-ExcelRange -Address $Sheet.Cells["D1:D$($Sheet.Dimension.Rows)"] -AutoSize -NumberFormat "Percentage"
Set-ExcelRange -Address $Sheet.Cells["H1:H$($Sheet.Dimension.Rows)"] -AutoSize -NumberFormat "Percentage"
Set-ExcelRange -Address $Sheet.Cells["L1:L$($Sheet.Dimension.Rows)"] -AutoSize -NumberFormat "Percentage"
Set-ExcelRange -Address $Sheet.Cells["B1:E$($Sheet.Dimension.Rows)"] -BorderAround "Medium"
Set-ExcelRange -Address $Sheet.Cells["F1:I$($Sheet.Dimension.Rows)"] -BorderAround "Medium"
Set-ExcelRange -Address $Sheet.Cells["J1:M$($Sheet.Dimension.Rows)"] -BorderAround "Medium"
Set-ExcelRange -Address $Sheet.Cells["A2:M$($Sheet.Dimension.Rows)"] -PatternColor "Blue"

for ($i = 2; $i -lt $Sheet.Dimension.Rows; $i++) {
  if ($i % 2 -eq 0) {
    $RowColor = "LightBlue"
  } else {
    $RowColor = "LightGray"
  }
  Set-ExcelRow -Worksheet $Sheet -Row $i -BackgroundColor $RowColor
}

# Add formulas for pecentages
Set-ExcelRange -Address $Sheet.Cells["D$($Sheet.Dimension.Rows)"] -Formula "=C$($Sheet.Dimension.Rows)/E$($Sheet.Dimension.Rows)"
Set-ExcelRange -Address $Sheet.Cells["H$($Sheet.Dimension.Rows)"] -Formula "=G$($Sheet.Dimension.Rows)/I$($Sheet.Dimension.Rows)"
Set-ExcelRange -Address $Sheet.Cells["L$($Sheet.Dimension.Rows)"] -Formula "=K$($Sheet.Dimension.Rows)/M$($Sheet.Dimension.Rows)"

# Save Excel file
$xl | Close-ExcelPackage
Write-Progress -Activity "Labor Sales Report" -Completed