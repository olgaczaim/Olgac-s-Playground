$WebURL ="[webname]"
$FileURL ="[webname]/[Filename].xlsx"
 
#Get Web and File Objects
$web = Get-SPWeb $WebURL
$File = $web.GetFile($FileURL)
$TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById(((Get-WmiObject win32_timezone).StandardName))

#Check if File is Checked-out
if ($File.CheckOutType -ne "None")
 {
    Write-host "File is Checked Out to user: " $File.CheckedOutByUser.LoginName
    Write-host "Checked Out Type: " $File.CheckOutType
    Write-host "Checked Out On: "  $File.CheckedOutDate
 
    #To Release from Checkout, Ask the checked out user to Checkin
    #$File.Checkin("Checked in by Administrator")
    #Write-host "File has been Checked-In"
 }
  
 #Check if File is locked
 if ($File.LockId -ne $null)
 {
     Write-host "File is Loked out by:" $File.LockedByUser.LoginName
     Write-host "File Lock Type: "$file.LockType
     Write-host "File Locked On: "$file.LockedDate, $TZ
     Write-host "File Lock Expires on: "$file.LockExpires, $TZ
 
     #To Release the lock, use:
     #$File.ReleaseLock($File.LockId)
     #Write-host "Released the lock!"
 }


<# another one #>
<#
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

$web = Get-SPWeb http://contonso/site/
$list = $web.Lists[“Documents”]
$item = $list.GetItemById(15)
$file = $item.File
$file

$userId = $file.LockedByUser.ID
$user = $web.AllUsers.GetByID($userId)
$impSite= New-Object Microsoft.SharePoint.SPSite($web.Url, $user.UserToken);

$impWeb = $impSite.OpenWeb();
$impList = $impWeb.Lists[$list.Title]
$impItem = $impList.GetItemById($item.ID)
$impFile = $impItem.File

$impFile.ReleaseLock($impFile.LockId)
$file

#>

<# and another one enhanced from upper one #>
<#
$web = Get-SPWeb "http://contonso/site/"
if($web){
    $list = $web.Lists["Documents"]
    $list.Title    
    $item = $list.GetItemById(111)

    $file = $item.File
    $file.Name

    $userId = $file.LockedByUser.ID
    if($userId){
        
        $user = $web.AllUsers.GetByID($userId)
        Write-Host "File is locked by user " $user.Name
        
        $impSite = New-Object Microsoft.SharePoint.SPSite($web.Url, $user.UserToken);
        $impWeb = $impSite.OpenWeb();
        $impList = $impWeb.Lists[$list.Title]
        $impItem = $impList.GetItemById($item.ID)
        $impFile = $impItem.File
        #$impFile.ReleaseLock($impFile.LockId)
        
        Write-Host "Realesed lock for " $impFile.Name 


    }else{
        Write-Host "File is not locked"
    }
}else{
    Write-Host "Web Not Found"
}

#>
