Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

#Variables for Processing
$SiteURL = "http://acme/supplies"
$ListName = "Documents"
  
Try {
   $web = Get-SPWeb $SiteURL
      
    #Get the List
    $List=$web.GetList($SiteURL + "/" +$ListName) 
    
    write-host -f Red $List.Title       
   
    $Query = New-Object Microsoft.SharePoint.SPQuery
    $Query.Query ="<Query><Where><Eq><FieldRef Name='ContentType' /><Value Type='Computed'>Folder</Value></Eq></Where></Query><ViewFields /><QueryOptions />"
    $ListItems = $List.GetItems($Query)
    
    $i=0
    #Loop through each list item
    ForEach($Item in $ListItems)
    {  
        if($Item.HasUniqueRoleAssignments -eq $False){
            $i++        
            $Item.BreakRoleInheritance($True)
            write-host  -f Green "Permission Inheritance Broken for " + $Item["FileLeafRef"]
        }
        
    } 
    write-host  -f Yellow $i
}
Catch {
    write-host -f Red "Error:" $_.Exception.Message
}
