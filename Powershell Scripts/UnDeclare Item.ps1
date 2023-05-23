Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

$SPAssignment = Start-SPAssignment
$web = Get-SPWeb "siteurlhere" -AssignmentCollection $SPAssignment
$list = $web.lists["libraryname"].items

foreach ($item in $list)
{
  $IsRecord = [Microsoft.Office.RecordsManagement.RecordsRepository.Records]::IsRecord($item)
  if ($IsRecord -eq $true)
  {
    Write-Host "Undeclared $($item.Name)" #or .Title
    [Microsoft.Office.RecordsManagement.RecordsRepository.Records]::UndeclareItemAsRecord($item)
  }
}

Stop-SPAssignment $SPAssignment
