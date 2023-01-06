function Get-LaborSalesCalendar {
  param (
    $ProductionYearStartDate = "1/1/2023"
  )

  $YearStartDate = Get-Date $ProductionYearStartDate
  $WeekSpan = [TimeSpan]::FromDays(7)
  $ShortMonthSpan = [TimeSpan]::FromDays(28)
  $LongMonthSpan = [TimeSpan]::FromDays(35)
  
  $WeekStartDates = @()
  $MonthStartDates = @()
  $QuarterStartDates = @()
  
  # Define Weeks
  $DateCursor = $YearStartDate
  while ($DateCursor.Year -lt ([int]$YearStartDate.Year + 1)) {
    $WeekStartDates += $DateCursor
    $DateCursor = $DateCursor + $WeekSpan
  }
  
  # Define Months
  $DateCursor = $YearStartDate
  while ($DateCursor.Year -lt ([int]$YearStartDate.Year + 1)) {
    $MonthStartDates += $DateCursor
    $IsThirdMonth = ($MonthStartDates.Length % 3) -eq 0
    if ($IsThirdMonth) {
      $DateCursor = $DateCursor + $LongMonthSpan
    } else {
      $DateCursor = $DateCursor + $ShortMonthSpan
    }
  }
  
  # Define Quarters
  for ($Index = 0; $Index -lt $MonthStartDates.Count; $Index++) {
    $IsQuarterStart = ($Index % 3) -eq 0
    if ($IsQuarterStart) {
      $QuarterStartDates += $MonthStartDates[$Index]
    }
  }

  # Find next year start date and add to all timespans
  $NextYearStartDate = $YearStartDate.AddYears(1)
  while ($NextYearStartDate.DayOfWeek -ne "Sunday") {
    $NextYearStartDate = $NextYearStartDate.AddDays(-1)
  }
  $WeekStartDates += $NextYearStartDate
  $MonthStartDates += $NextYearStartDate
  $QuarterStartDates += $NextYearStartDate

  return [PSCustomObject]@{
    WeekStartDates = $WeekStartDates
    MonthStartDates =  $MonthStartDates
    QuarterStartDates =  $QuarterStartDates
  }
}

function Get-IndexOfDateFromList {
  param (
      [Parameter(Mandatory)][datetime]$Date,
      [Parameter(Mandatory)]$DateList
  )

  $DateIndex = -1
  
  for ($i = 0; $i -lt $DateList.Count; $i++) {
    try {
      $IsDateInTimespan =
        $Date -ge $DateList[$i] -and
        $Date -lt $DateList[$i + 1]

      if ($IsDateInTimespan) {
        $DateIndex = $i
        break
      }
    } catch {
      $DateIndex = -1
    }
  }

  return $DateIndex
}