#Gets timer job history on all web applications
#usage .\jobhist.ps1 -JobID "87f1ae66-159a-4ea6-b0c4-a2114236f831" -StartTime "8/20/2022 00:00:00 AM" -EndTime "8/21/2022 11:59:59 PM"

Param(
    [parameter(position=0)]
    [DateTime]
        $StartTime,
    [parameter(position=1)]
    [DateTime]
        $EndTime,
    [parameter(position=2)]
    [String]
        $JobName,
    [parameter(position=3)]
    [String]
        $JobID
)
 
if(!$StartTime) {$StartTime = (Get-Date).Date::MinValue}
if(!$EndTime) {$EndTime = (Get-Date).Date::MaxValue}
 
$StartTime = $StartTime.ToUniversalTime()
$EndTime = $EndTime.ToUniversalTime()

$TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById(((Get-WmiObject win32_timezone).StandardName))
 
Get-SPWebApplication | foreach {
    $_.JobHistoryEntries |
        where{  (($StartTime -le $_.StartTime -and $_.StartTime -le $EndTime) -and
            ($StartTime -le $_.EndTime -and $_.EndTime -le $EndTime)) -and ( ($_.JobDefinitionId -eq $JobID) -or ( $_.JobDefinitionTitle -eq $JobName ) ) } |
        sort StartTime |
        select JobDefinitionId,  
            JobDefinitionTitle,
            WebApplicationName,
            ServerName,
            Status,
            @{Expression={[System.TimeZoneInfo]::ConvertTimeFromUtc($_.StartTime, $TZ)};Label="Start Time"},
            @{Expression={[System.TimeZoneInfo]::ConvertTimeFromUtc($_.EndTime, $TZ)};Label="End Time"},
            @{Expression={($_.EndTime - $_.StartTime).TotalSeconds};Label="Duration (secs)"}
} | Out-GridView -Title "Timer Job History"


/////////////////////////////////////////////////////////////////

#Gets timer job history on all farm
#usage .\jobhist.ps1 -JobID "87f1ae66-159a-4ea6-b0c4-a2114236f831" -StartTime "8/21/2022 00:00:00 AM" -EndTime "8/21/2022 11:59:59 PM"
Param(
    [parameter(position=0)]
    [DateTime]
        $StartTime,
    [parameter(position=1)]
    [DateTime]
        $EndTime,
    [parameter(position=2)]
    [string]
        $JobName,
    [parameter(position=3)]
    [string]
        $JobID
)
 
if(!$StartTime) {$StartTime = (Get-Date).Date::MinValue}
if(!$EndTime) {$EndTime = (Get-Date).Date::MaxValue}


$StartTime = $StartTime.ToUniversalTime()
$EndTime = $EndTime.ToUniversalTime()

$TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById(((Get-WmiObject win32_timezone).StandardName))
 

#Get the specific Timer job
$Timerjob = Get-SPTimerJob | where { ($_.Id -eq $JobID) -or ($_.Title -eq $JobName) -or (!$JobName -and !$JobID) }

$Results = $Timerjob.HistoryEntries  |
      where { ($StartTime -le $_.StartTime -and $_.StartTime -le $EndTime) -and
            ($StartTizme -le $_.EndTime -and $_.EndTime -le $EndTime) } |
          Select JobDefinitionId,WebApplicationName,JobDefinitionTitle,ServerName,Status,ErrorMessage,
          @{Expression={[System.TimeZoneInfo]::ConvertTimeFromUtc($_.StartTime, $TZ)};Label="Start Time"},
            @{Expression={[System.TimeZoneInfo]::ConvertTimeFromUtc($_.EndTime, $TZ)};Label="End Time"},
            @{Expression={($_.EndTime - $_.StartTime).TotalSeconds};Label="Duration (secs)"}


#Send results to Grid view   
$Results | Out-GridView

#for csv result
#$OutPutFile="C:\script\TimerJobHistory.csv"
#$Results | Export-Csv $OutPutFile -NoType

#for failed jobs
#-and ($_.Status -ne 'Succeeded')
