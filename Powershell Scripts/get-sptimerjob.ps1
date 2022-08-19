//get sp timerjob

Get-SPTimerJob -WebApplication "*****" | select ID,Name, DisplayName
Get-SPTimerJob -Identity 87f1ae66-159a-4ea6-b0c4-a2114236f831 | Format-Table DisplayName,Id,LastRunTime,Status
(Get-SPTimerJob -Identity 87f1ae66-159a-4ea6-b0c4-a2114236f831).HistoryEntries | Format-Table -Property Status,StartTime,EndTime,ErrorMessage

//////////////////////////////////////////////////////
#get all timer jobs filter by date-time
Param(
    [parameter(position=0)]
    [DateTime]
        $StartTime,
    [parameter(position=1)]
    [DateTime]
        $EndTime
)
 
if(!$StartTime) {$StartTime = (Get-Date).Date}
if(!$EndTime) {$EndTime = (Get-Date).AddDays(1).Date}
 
$StartTime = $StartTime.ToUniversalTime()
$EndTime = $EndTime.ToUniversalTime()
 
$TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById(((Get-WmiObject win32_timezone).StandardName))
$WebApp = Get-SPTimerJob
$Results = $WebApp.HistoryEntries |
        where{  ($StartTime -le $_.StartTime -and $_.StartTime -le $EndTime) -or
            ($StartTime -le $_.EndTime -and $_.EndTime -le $EndTime) } |
        sort StartTime |  select JobDefinitionId,JobDefinitionTitle,WebApplicationId,WebApplicationName,ServerName,DatabaseName,ErrorMessage,
            @{Expression={[System.TimeZoneInfo]::ConvertTimeFromUtc($_.StartTime, $TZ)};Label="Start Time"},
            @{Expression={[System.TimeZoneInfo]::ConvertTimeFromUtc($_.EndTime, $TZ)};Label="End Time"},
            @{Expression={($_.EndTime - $_.StartTime).TotalSeconds};Label="Duration (secs)"}
$Results | Out-GridView -Title "Timer Job History"
#usage .\jobhist.ps1 "8/18/2022 6:00 pm" "8/18/2022 6:30 pm"
//////////////////////////////////////////////////////////////////////////////

#getTimerJob from all web application
Param(
    [parameter(position=0)]
    [DateTime]
        $StartTime,
    [parameter(position=1)]
    [DateTime]
        $EndTime
)
 
if(!$StartTime) {$StartTime = (Get-Date).Date}
if(!$EndTime) {$EndTime = (Get-Date).AddDays(1).Date}
 
$StartTime = $StartTime.ToUniversalTime()
$EndTime = $EndTime.ToUniversalTime()
 
$TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById(((Get-WmiObject win32_timezone).StandardName))
 
Get-SPWebApplication | foreach {
    $_.JobHistoryEntries |
        where{  ($StartTime -le $_.StartTime -and $_.StartTime -le $EndTime) -or
            ($StartTime -le $_.EndTime -and $_.EndTime -le $EndTime) } |
        sort StartTime |
        select  JobDefinitionTitle,
            WebApplicationName,
            ServerName,
            Status,
            @{Expression={[System.TimeZoneInfo]::ConvertTimeFromUtc($_.StartTime, $TZ)};Label="Start Time"},
            @{Expression={[System.TimeZoneInfo]::ConvertTimeFromUtc($_.EndTime, $TZ)};Label="End Time"},
            @{Expression={($_.EndTime - $_.StartTime).TotalSeconds};Label="Duration (secs)"}
} | Out-GridView -Title "Timer Job History"
#usage .\jobhist.ps1 "8/18/2022 6:00 pm" "8/18/2022 6:30 pm"


///////////////////////////////////////
#get specific timerjob
# Variables
$StartTime = "09/26/2015 01:00:00 AM"  # mm/dd/yyyy hh:mm:ss
$EndTime = "09/26/2015 01:30:00 AM"
$TimerJobName = "Immediate Alerts"
 
#To Get Yesterday's use:
#$StartDateTime = (Get-Date).AddDays(-1).ToString('MM-dd-yyyy') + " 00:00:00"
#$EndDateTime   = (Get-Date).AddDays(-1).ToString('MM-dd-yyyy') + " 23:59:59"
 
#Get the specific Timer job
$Timerjob = Get-SPTimerJob | where { $_.DisplayName -eq $TimerJobName }
 
#Get all timer job history from the web application
$Results = $Timerjob.HistoryEntries  |
      where { ($_.StartTime -ge  $StartTime) -and ($_.EndTime -le $EndTime) } |
          Select WebApplicationName,ServerName,Status,StartTime,EndTime
 
#Send results to Grid view   
$Results | Out-GridView


#Read more: https://www.sharepointdiary.com/2015/09/get-timer-job-history-using-powershell.html#ixzz7cOR9QnnN
