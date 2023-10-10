Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
#Function to Download All Files from a SharePoint Folder
Function Download-SPFolder($SPFolderURL, $DownloadPath)
{
    Try {
        #Get the Source SharePoint Folder
        $SPFolder = $web.GetFolder($SPFolderURL)
  
        $DownloadPath = Join-Path $DownloadPath $SPFolder.Name
        #Ensure the destination local folder exists!
        If (!(Test-Path -path $DownloadPath))
        {   
            #If it doesn't exists, Create
            $LocalFolder = New-Item $DownloadPath -type directory
        }
  
        #Loop through each file in the folder and download it to Destination
        ForEach ($File in $SPFolder.Files)
        {
            #Download the file
            $Data = $File.OpenBinary()
            $FilePath= Join-Path $DownloadPath $File.Name
            [System.IO.File]::WriteAllBytes($FilePath, $data)
            Write-host -f Green "`tDownloaded the File:"$File.ServerRelativeURL        
        }
  
        #Process the Sub Folders & Recursively call the function
        ForEach ($SubFolder in $SPFolder.SubFolders)
        {
            If($SubFolder.Name -ne "Forms") #Leave "Forms" Folder
            {
                #Call the function Recursively
                Download-SPFolder $SubFolder $DownloadPath
            }
        }
    }
    Catch {
        Write-host -f Red "Error Downloading Document Library:" $_.Exception.Message
    } 
}
 
#Main Function
Function Download-SPDocumentLibrary($SiteURL, $LibraryName, $DownloadPath)
{
    Try {
        #Get the  Web
        $Web = Get-SPWeb $SiteURL
 
        #Delete any existing files and folders in the download location
        If (Test-Path $DownloadPath) {Get-ChildItem -Path $DownloadPath -Recurse| ForEach-object {Remove-item -Recurse -path $_.FullName }}
 
        #Get the document Library to Download
        $Library = $Web.Lists[$LibraryName]
        Write-host -f magenta "Downloading Document Library:" $Library.Title
 
         #Call the function to download the document library
        Download-SPFolder -SPFolderURL $Library.RootFolder.Url -DownloadPath $DownloadPath
 
        Write-host -f Green "*** Download Completed  ***"
    }
    Catch {
        Write-host -f Red "Error Downloading Document Library:" $_.Exception.Message
    } 
}
 
#Runtime-Variables
$SiteURL = "http://contonso.com/"
$LibraryName ="SharedDocuments"
$DownloadPath ="C:\Downloads\SPEXPORT"
 
#Call the Function to export all document libraries from a site
Download-SPDocumentLibrary $SiteURL $LibraryName $DownloadPath


#Read more: https://www.sharepointdiary.com/2018/05/download-all-files-in-sharepoint-document-library-using-powershell.html#ixzz8FifrUApM
