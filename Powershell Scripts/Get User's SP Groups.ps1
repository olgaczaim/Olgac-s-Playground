Get-SPUser -Web https://contonso.com -Limit all | Where-Object {$_.UserLogin -like "*olgac*"} | Select-Object UserLogin, @{name="Groups";expression={($_.Groups -join [environment]::NewLine)}} | Format-Table -AutoSize -Wrap

/*or we can wrap with "| Out-GridView" */
