#delete a User Profile in SharePoint:
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
#Configuration Variables
$SiteURL = "https://mypage.crescent.com"
$AccountName="Crescent\Sherif"
 
#Get Objects
$ServiceContext  = Get-SPServiceContext -site $SiteURL
$UserProfileManager = New-Object Microsoft.Office.Server.UserProfiles.UserProfileManager($ServiceContext)
 
#Get the User Profile
$UserProfile = $UserProfileManager.GetUserProfile($AccountName)
 
#remove user profile
$UserProfileManager.RemoveUserProfile($AccountName); 



#Delete All User Profiles in SharePoint
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
#Configuration Variables
$SiteURL = "https://mypage.crescent.com"
 
#Get Objects
$ServiceContext  = Get-SPServiceContext -site $SiteURL
$UserProfileManager = New-Object Microsoft.Office.Server.UserProfiles.UserProfileManager($ServiceContext)
 
#Ger all User Profiles
$UserProfiles = $UserProfileManager.GetEnumerator()
  
# Loop through user profile
Foreach ($Profile in $UserProfiles)
{
    write-host Removing User Profile: $Profile["AccountName"]
     
    #Remove User Profile
    $UserProfileManager.RemoveUserProfile($profile["AccountName"])
}

#Delete on specific domain
$DomainPrefix = "Crescent"
$UserProfiles = $UserProfileManager.GetEnumerator()| Where-Object {$_.Accountname -like "$($DomainPrefix)\*"}
