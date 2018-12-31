#POWERSHELL SCRIPT TO EXTRACT ALL THE WSP FILES FROM CURRENT SHAREPOINT FARM

#for one solution
$farm = Get-SPFarm
$file = $farm.Solutions.Item(“solution.wsp”).SolutionFile
$file.SaveAs(“c:\Solution.wsp”)

#for multiple solutions
$dirName = "<directory path>" 
Write-Host Exporting solutions to $dirName  
foreach ($solution in Get-SPSolution)  
{  
    $id = $Solution.SolutionID  
    $title = $Solution.Name  
    $filename = $Solution.SolutionFile.Name 
    Write-Host "Exporting ‘$title’ to …\$filename" -nonewline  
    try {  
        $solution.SolutionFile.SaveAs("$dirName\$filename")  
        Write-Host " – done" -foreground green  
    }  
    catch  
    {  
        Write-Host " – error : $_" -foreground red  
    }  
}

#another option
(Get-SPFarm).Solutions | ForEach-Object{$var = (Get-Location).Path + “\” + $_.Name; $_.SolutionFile.SaveAs($var)}
