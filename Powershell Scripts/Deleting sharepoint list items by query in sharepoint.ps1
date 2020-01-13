for($i=59; $i -le 122; $i++){
Write-Host "Preparing to delete Item - $i"
Add-PSSnapin Microsoft.SharePoint.Powershell -ea SilentlyContinue
$web = get-spweb "https://xxxxxxxxxxx"
$list = $web.lists["ListName"]
$query = New-Object Microsoft.SharePoint.SPQuery

# Added Code Start
$createdOnInternalFieldName = 'xID'
$caml = '<Where><Eq><FieldRef Name="{0}"  /><Value Type="Text">{1}</Value></Eq></Where>'-f $createdOnInternalFieldName,$i
$query.Query = $caml
# Added Code End
Write-Host "Query - $caml"
$query.ViewAttributes = "Scope='Recursive'"
$query.RowLimit = 1000
$query.ViewFields = "<FieldRef Name='ID'/><FieldRef Name='Title' /><FieldRef Name='xID' />"
$query.ViewFieldsOnly = $true
do
{
    $listItems = $list.GetItems($query)
    $query.ListItemCollectionPosition = $listItems.ListItemCollectionPosition
    foreach($item in $listItems)
    {
        Write-Host "Deleting Item - $($item.Id) - $($item.Title) - $($item['DuyuruID'])"
        $list.GetItemById($item.Id).delete()
    }
}
while ($query.ListItemCollectionPosition -ne $null) 
}