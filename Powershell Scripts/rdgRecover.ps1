# Path to RDCMan.exe
$RDCMan = "C:\Program Files (x86)\Microsoft\Remote Desktop Connection Manager\RDCMan.exe"
# Path to RDG file
$RDGFile = "$env:USERPROFILE\Documents\remotes1.rdg"
$TempLocation = "C:\temp"
 
Copy-Item $RDCMan "$TempLocation\RDCMan.dll"
Import-Module "$TempLocation\RDCMan.dll"
$EncryptionSettings = New-Object -TypeName RdcMan.EncryptionSettings
 
$XML = New-Object -TypeName XML
$XML.Load($RDGFile)
$logonCredentials = Select-XML -Xml $XML -XPath '//logonCredentials'
 
$Credentials = New-Object System.Collections.Arraylist
$logonCredentials | foreach {
    [void]$Credentials.Add([pscustomobject]@{
    Username = $_.Node.userName
    Password = $(Try{[RdcMan.Encryption]::DecryptString($_.Node.password, $EncryptionSettings)}Catch{$_.Exception.InnerException.Message})
    Domain = $_.Node.domain
    })
    } | Sort Username
 
$Credentials | Sort Username
