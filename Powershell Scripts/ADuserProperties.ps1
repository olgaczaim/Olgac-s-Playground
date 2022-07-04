get user properties with windows powershell

$username = "xxxxx"

$searcher = New-Object System.DirectoryServices.DirectorySearcher
$searcher.filter ="(&(objectcategory=person)(SamAccountName=$($username)))"
$searcher.FindOne().Properties



(New-Object System.DirectoryServices.DirectorySearcher("(&(objectCategory=User)(samAccountName=$($username)))")).Properties


(New-Object System.DirectoryServices.DirectorySearcher("(&(objectCategory=User)(samAccountName=$($username)))")).FindOne().GetDirectoryEntry().memberOf
