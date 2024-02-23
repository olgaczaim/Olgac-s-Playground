# Working with the .NET Client Object Model

- Before you can start working with the .NET Client Object Model, you have to add a reference to the assemblies Microsoft.SharePoint.Client.dll and Microsoft.SharePoint.Client.Runtime.dll. You can find these assemblies in the 14\ISAPI folder.

#### The Where clause

The Where clause can range from very simple to very complex with one or more nested `<And>` or `<Or>` elements. The CAML for the Where clause has not been changed since SharePoint 2003, and can be used with the SharePoint Server Object Model and with the three Client Object Models. Also the Web Services are still there, for which you can also use CAML to retrieve and/or update list items.

To retrieve list items from a SharePoint list, you have to define an instance of type CamlQuery, defined within the Microsoft.SharePoint.Client.CamlQuery namespace. You can specify your CAML query within the ViewXml property. This property is of type string but its content must be XML. The root element for this property is `<View>`. The Where clause needs to be embedded within a `<Query>` element.

```csharp
Microsoft.SharePoint.Client.CamlQuery query = new Microsoft.SharePoint.Client.CamlQuery();
query.ViewXml = "<View>"
     + "<Query>"
     + "<Where><Eq><FieldRef Name='Country' /><Value Type='Text'>Belgium</Value></Eq></Where>"
     + "</Query>"
     + "</View>";
// execute the query
ListItemCollection listItems = spList.GetItems(query);
clientContext.Load(listItems);
clientContext.ExecuteQuery();

```
####OrderBy
The OrderBy element is the most simple one: you can define a sort order using one or more `<FieldRef>` elements that you include in the ViewXml property of the CamlQuery object:

```csharp
query.ViewXml = "<View>"
     + "<Query>"
     + "<Where><Eq><FieldRef Name='Country' /><Value Type='Text'>Belgium</Value></Eq></Where>"
     + "<OrderBy><FieldRef Name='City'/></OrderBy>"
     + "</Query>"
     + "</View>";

```
####ViewFields
You can also limit the number of columns returned to the client, using the good old ViewFields element, which you can also include in the ViewXml property of the CamlQuery object:

```csharp
CamlQuery query = new CamlQuery();
query.ViewXml = "<View>"  
     + "<Query>"
     + "<Where><Eq><FieldRef Name='Country' /><Value Type='Text'>Belgium</Value></Eq></Where>"
     + "</Query>"
     + "<ViewFields>"
     + "  <FieldRef Name='Title' /><FieldRef Name='City' />"
     + "</ViewFields>"
     + "</View>";
// execute the query
ListItemCollection listItems = spList.GetItems(query);
clientContext.Load(listItems);
clientContext.ExecuteQuery();

```
But this also returns a number of system columns. If you really want to limit the columns returned to the columns you specify, you have to use a LINQ query within the Load method. The code looks as follows:

```csharp
CamlQuery camlQuery = new CamlQuery();
camlQuery.ViewXml = "<View><Where><Eq><FieldRef Name='Country' /><Value Type='Text'>Belgium</Value></Eq></Where></View>";
ListItemCollection listItems = spList.GetItems(camlQuery);
clientContext.Load(listItems,
      items => items.Include(
      item => item.Id,
      item => item.DisplayName,
      item => item.HasUniqueRoleAssignments));
clientContext.ExecuteQuery();

```
####Query Options
The different Query Options need to be handled a bit differently than with the SharePoint Server Object Model.

The row limit can also be specified within the ViewXml property:

```csharp
query.ViewXml = "<View>"
     + "<Query>"
     + "<Where><Eq><FieldRef Name='Country' /><Value Type='Text'>Belgium</Value></Eq></Where>"
     + "<OrderBy><FieldRef Name='City'/></OrderBy>"
     + "</Query>"
     + "<RowLimit>5</RowLimit>"
     + "</View>";

```

#####Dates in UTC
You can choose to return dates in UTC (Coordinated Universal Time)  by setting the DatesInUtc property of the CamlQuery instance:
`query.DatesInUtc = true;`

#####Include attachment URLs
Using CAML you are able to know if list items have attachment by adding a w> element to the ViewFields element in the ViewXml property:
```csharp
query.ViewXml = "<View>"  
     + "<ViewFields>"
     + "  <FieldRef Name='Title' /><FieldRef Name='City' /><FieldRef Name='Attachments' />"
     + "</ViewFields>"
     + "</View>";

```

SharePoint will return a boolean indicating whether the list item has attachments or not.

The attechments are not stored in the list item itself, but are stored in a sub folder of the list. More specifically, the list contains a folder named Attachments and if a list item has one or more attachments, a folder is created based on the ID of the list item. This sub folder will then contain the attachment(s). The URL of the attachment is not stored in the list item itself.

CAML contains an option IncludeAttachmentURLs that can be used to retreive the URL of the attachment(s), together with the other properties of the list item. It works on the server side SPQuery and with the `<QueryOptions>` node of the GetListItems method of the Lists.asmx web service , but it doesn’t seem to be available with the CamlQuery object of the .NET Client Object Model.

If you need to retrieving the attachments itself you will have to write some extra code that retrieves the files from the Attachment folder:

```csharp
Folder folder = clientContext.Web.GetFolderByServerRelativeUrl(
    spList.RootFolder.ServerRelativeUrl + "/Attachments/" + item.Id);    
FileCollection files = attFolder.Files;    
// If you only need the URLs    
ctx.Load(files, fs => fs.Include(f => f.ServerRelativeUrl));    
ctx.ExecuteQuery();

```
#####Limitations
Following CAML subtilities doesn’t seem to be working with the CamlQuery object of the .NET Client Object Model, although they exist when retrieving list items with the server object model and the SharePoint web services:
- IncludeMandatoryColumns: this option also returns the required fields besides the other fields specified in the ViewFields property or element.
- ExpandUserField: when you query a User field, you only see the login name of the user. When you indicate that you want to expand a user field, SharePoint will also return information like the user name and the email address.
- IncludeAttachmentURLs: cfr. higher
- IncludeAttachmentVersion:

#####Files and folders options
CAML for retrieving files and folders at different levels of a document library, is always a bit more complex. To make the explanation hereunder a bit more readable, I have created a document library with the following structure:

You can easily query the files and folders in the root folder of a document library without having to use specific CAML elements. Only if you want to start querying the folder structure within a document library, you have to apply specific CAML.
To be able to better demonstrate the subtilities I created a folder structure in my Shared Documents library, and added a set of files to the different folders. 

For example, if you want to query all files and folders in your document library, no matter how deep they are nested,  you have to add a Scope attribute to the View element, and set its value to RecursiveAll:

`query.ViewXml = "<View Scope='RecursiveAll'></View>";`

You can always add a Query element in the View element and specify a Where clause to add an extra filter to the query, or an OrderBy clause to sort the result.

If you want to query only the folders, you have to add an extra where clause:

```csharp
query.ViewXml = "<View Scope='RecursiveAll'>"
     + "<Query>"
     + "   <Where>"
     + "      <Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>1</Value></Eq>"
     + "   </Where>"
     + "</Query>"
     + "</View>";

```

If you want to query only the files, the extra where clause can be changed as follows:

```csharp
query.ViewXml = "<View Scope='RecursiveAll'>"
     + "<Query>"
     + "   <Where>"
     + "      <Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>0</Value></Eq>"
     + "   </Where>"
     + "</Query>"
     + "</View>";

```

If you want to retrieve the content of a specific folder, i.e files and folders, you have to add the relative URL to that folder to the **FolderServerRelativeUrl** property of the query instance:

`query.FolderServerRelativeUrl = "/Shared Documents/Folder 1";`

If you only want to see the files of a specific sub folder, you have to set the **Scope** attribute of the **ViewXml** property to **FilesOnly**:

```csharp
query.ViewXml = "<View Scope='FilesOnly' />";
query.FolderServerRelativeUrl = "/Shared Documents/Folder 1";

```

Of course you can also query all files in a specific sub folder and its underlying sub folders. In that case you also have to specify the relative URL to the folder, but you also have to set the **Scope** attribute of the **ViewXml** property to **Recursive**:

```csharp
query.ViewXml = "<View Scope='Recursive' />";
query.FolderServerRelativeUrl = "/Shared Documents/Folder 1";

```

If you want to retrieve all files AND folders from a specific folder and its underlying sub folders, you have to set the **Scope** attribute of the **ViewXml** property to **RecursiveAll**: 

```csharp
query.ViewXml = "<View Scope='RecursiveAll' />";
query.FolderServerRelativeUrl = "/Shared Documents/Folder 1";

```
[Referance](https://karinebosch.wordpress.com/2012/02/03/caml-and-the-client-object-model/ "Referance")

