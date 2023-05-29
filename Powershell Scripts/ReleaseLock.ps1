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
