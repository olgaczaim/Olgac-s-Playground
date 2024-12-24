using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System.Net;
using System.Linq;

string siteUrl = "http://contonso.com/";
string username = "username";
string pass = "password";
string domain = "domain";
string fileUrl = "/Shared%20Documents/Test.docx";
string filedir = "/Shared Documents";
string siteColumnName = "Test_Column";


/// <summary>
/// NuGet Packages
/// PnP.Framework
/// DocumentFormat.OpenXml
/// Using PnP.Framework with .Net Core / Standard
/// </summary>


NetworkCredential credentials = new NetworkCredential(username, pass, domain);

using (ClientContext context = new PnP.Framework.AuthenticationManager().GetOnPremisesContext(siteUrl, credentials))
{
    /*dosya checking / checkout */
    //await CheckinCheckout(fileUrl, context);

    /*Add Document to specified field*/
    //await AddFileToSpFolder(filedir, context);

    /*Adding site column to library*/
    //await AddSiteColumnToListIfMissing(fileUrl, context, siteColumnName);

    /*Column value edit*/
    //await UpdateItemProperty(context, fileUrl, "https://www.google.com.tr");

    /*Creating file based on target file's header*/
    //await DocHeaderUpdate(context, fileUrl, filedir);

    /*permission check and add*/
    //await CheckPermissionAndAddReadPerm(context, fileUrl, "testuser2");

    Console.ReadLine();
}


async Task CheckinCheckout(string fileUrl, ClientContext context)
{
    try
    {
        // Get the file object by its server-relative URL
        var file = context.Web.GetFileByServerRelativeUrl(fileUrl);
        context.Load(file, f => f.CheckOutType, f => f.CheckedOutByUser, f => f.Name, f => f.TimeCreated);
        await context.ExecuteQueryAsync();

        Console.WriteLine($"File: {file.Name}");
        Console.WriteLine($"Check Out Type: {file.CheckOutType}");

        if (file.CheckOutType == CheckOutType.None)
        {
            Console.WriteLine("Checking out the file...");
            file.CheckOut();
            await context.ExecuteQueryAsync();
            Console.WriteLine("File checked out successfully.");
        }
        else if (file.CheckOutType == CheckOutType.Online && file.CheckedOutByUser.LoginName == @"i:0#.w|domain\username")
        {
            Console.WriteLine($"Checked Out By: {file.CheckedOutByUser.LoginName ?? "None"}");
            Console.WriteLine("File is checked out by the system. Undoing checkout...");
            file.UndoCheckOut();
            await context.ExecuteQueryAsync();
            Console.WriteLine("Undo check-out successful.");
        }
        else
        {
            Console.WriteLine("No action needed on the file.");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Exception: {ex.Message}");
    }
}

async Task AddFileToSpFolder(string fileDir, ClientContext context, string fileName = "testlog12.txt")
{
    byte[] fileContent = await System.IO.File.ReadAllBytesAsync(@"c:\log.txt");
    var folder = context.Web.GetFolderByServerRelativeUrl(fileDir);
    context.Load(folder);
    await context.ExecuteQueryAsync();

    Console.WriteLine($"Uploading file to folder: {fileDir}");


    using (var stream = new MemoryStream(fileContent))
    {
        var fileCreationInfo = new FileCreationInformation
        {
            ContentStream = stream,
            Url = fileName,
            Overwrite = true
        };

        folder.Files.Add(fileCreationInfo);
        await context.ExecuteQueryAsync();

        Console.WriteLine("File uploaded successfully.");
    }

}

async Task AddSiteColumnToListIfMissing(string fileUrl, ClientContext context, string siteColumnName)
{
    try
    {
        // Get the file and associated list
        var file = context.Web.GetFileByServerRelativeUrl(fileUrl);
        context.Load(file, f => f.ListItemAllFields.ParentList.Id);
        context.ExecuteQuery();

        var list = context.Web.Lists.GetById(file.ListItemAllFields.ParentList.Id);
        context.Load(list, l => l.Fields, l => l.ContentTypesEnabled);
        context.ExecuteQuery();

        // Enable content types if not already enabled
        if (!list.ContentTypesEnabled)
        {
            Console.WriteLine("Content types are not enabled. Enabling content types...");
            list.ContentTypesEnabled = true;
            list.Update();
            context.ExecuteQuery();
            Console.WriteLine("Content types have been enabled.");
        }

        // Check if the site column exists in the list
        bool columnExists = list.Fields.Any(f => f.InternalName == siteColumnName);
        if (!columnExists)
        {
            Console.WriteLine($"Adding site column '{siteColumnName}' to the list.");
            var siteColumn = context.Web.Fields.GetByInternalNameOrTitle(siteColumnName);
            context.Load(siteColumn);
            context.ExecuteQuery();

            list.Fields.Add(siteColumn);
            list.Update();
            context.ExecuteQuery();
            Console.WriteLine($"Site column '{siteColumnName}' added to the list.");
        }

        // Add the site column to the default view if not already present
        var defaultView = list.DefaultView;
        context.Load(defaultView, v => v.ViewFields);
        context.ExecuteQuery();

        if (!defaultView.ViewFields.Contains(siteColumnName))
        {
            Console.WriteLine($"Adding site column '{siteColumnName}' to the default view.");
            defaultView.ViewFields.Add(siteColumnName);
            defaultView.Update();
            context.ExecuteQuery();
            Console.WriteLine($"Site column '{siteColumnName}' added to the default view.");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error while adding site column: {ex.Message}");
    }
}

async Task UpdateItemProperty(ClientContext context, string fileUrl, string href)
{
    try
    {
        // Load the file and its associated list item
        var file = context.Web.GetFileByServerRelativeUrl(fileUrl);
        context.Load(file, f => f.ListItemAllFields);
        context.ExecuteQuery();

        // Check if the custom column exists
        if (file.ListItemAllFields.FieldValues.ContainsKey("Test_Column"))
        {
            var fieldValue = new FieldUrlValue
            {
                Url = href,
                Description = "Column Value"
            };

            // Update the custom field
            file.ListItemAllFields["Test_Column"] = fieldValue;
            file.ListItemAllFields.Update();
            context.ExecuteQuery();

            Console.WriteLine("Custom column updated successfully.");
        }
        else
        {
            Console.WriteLine("The column 'Test_Column' does not exist in the list.");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error updating item property: {ex.Message}");
    }
}

async Task DocHeaderUpdate(ClientContext context, string fileUrl, string fileDir)
{
    try
    {
        var file = context.Web.GetFileByServerRelativeUrl(fileUrl);
        context.Load(file);
        context.ExecuteQuery();

        // Download the file content into a stream
        var fileContent = file.OpenBinaryStream();
        context.ExecuteQuery();

        using (MemoryStream sourceDocStream = new MemoryStream())
        {
            fileContent.Value.CopyTo(sourceDocStream);
            sourceDocStream.Position = 0;

            using (WordprocessingDocument sourceDoc = WordprocessingDocument.Open(sourceDocStream, false))
            {
                var mainDocumentPart = sourceDoc.MainDocumentPart;
                var headerParts = mainDocumentPart.HeaderParts;

                using (MemoryStream newDocStream = new MemoryStream())
                {

                    using (WordprocessingDocument newDoc = WordprocessingDocument.Create(newDocStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                    {
                        MainDocumentPart newMainPart = newDoc.AddMainDocumentPart();
                        newMainPart.Document = new Document(new Body());

                        PageMargin margin = sourceDoc.MainDocumentPart.Document.Body.GetFirstChild<SectionProperties>().GetFirstChild<PageMargin>();


                        // Copy styles
                        if (mainDocumentPart.StyleDefinitionsPart != null)
                        {
                            newMainPart.AddPart(mainDocumentPart.StyleDefinitionsPart, mainDocumentPart.GetIdOfPart(mainDocumentPart.StyleDefinitionsPart));
                        }

                        // Copy numbering (if there are lists)
                        if (mainDocumentPart.NumberingDefinitionsPart != null)
                        {
                            newMainPart.AddPart(mainDocumentPart.NumberingDefinitionsPart, mainDocumentPart.GetIdOfPart(mainDocumentPart.NumberingDefinitionsPart));
                        }

                        // Copy other styling-related parts if needed
                        foreach (var styleRel in mainDocumentPart.Parts)
                        {
                            if (styleRel.OpenXmlPart is CustomXmlPart || styleRel.OpenXmlPart is ThemePart)
                            {
                                newMainPart.AddPart(styleRel.OpenXmlPart, styleRel.RelationshipId);
                            }
                        }

                        //Create a placeholder paragraph
                        //DocumentFormat.OpenXml.Wordprocessing.Paragraph para = body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
                        //DocumentFormat.OpenXml.Wordprocessing.Run run = para.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run());
                        //run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text("Temp"));

                        if (headerParts.Count() > 0)
                        {
                            HeaderPart newHeaderPart = newMainPart.AddNewPart<HeaderPart>();

                            foreach (var headerPart in headerParts)
                            {
                                newHeaderPart.Header = (Header)headerPart.Header.CloneNode(true);

                                foreach (var relationship in headerPart.Parts)
                                {
                                    if (relationship.OpenXmlPart is ImagePart sourceImagePart)
                                    {
                                        ImagePart newImagePart = newHeaderPart.AddImagePart(sourceImagePart.ContentType);
                                        using (Stream sourceStream = sourceImagePart.GetStream())
                                        using (Stream targetStream = newImagePart.GetStream())
                                        {
                                            sourceStream.CopyTo(targetStream);
                                        }

                                        string oldRelId = headerPart.GetIdOfPart(sourceImagePart);
                                        string newRelId = newHeaderPart.GetIdOfPart(newImagePart);

                                        foreach (var blip in newHeaderPart.Header.Descendants<DocumentFormat.OpenXml.Drawing.Blip>())
                                        {
                                            if (blip.Embed == oldRelId)
                                            {
                                                blip.Embed = newRelId;
                                            }
                                        }
                                    }
                                }

                                string headerRelId = newMainPart.GetIdOfPart(newHeaderPart);

                                SectionProperties sectionProps = new SectionProperties();
                                HeaderReference headerRef = new HeaderReference()
                                {
                                    Type = HeaderFooterValues.Default,
                                    Id = headerRelId
                                };
                                sectionProps.Append(headerRef);
                                sectionProps.Append((PageMargin)margin.CloneNode(true));
                                newMainPart.Document.Body.Append(sectionProps);
                            }
                        }

                        newMainPart.Document.Save();
                    }

                    // Save the newly created document to SharePoint
                    newDocStream.Position = 0;

                    Folder folder = context.Web.GetFolderByServerRelativeUrl(filedir);
                    context.Load(folder);
                    context.ExecuteQuery();

                    FileCreationInformation newFile = new FileCreationInformation
                    {
                        ContentStream = newDocStream,                        
                        Url = "newdoc.docx",
                        Overwrite = true
                    };

                    Microsoft.SharePoint.Client.File uploadFile = folder.Files.Add(newFile);
                    context.Load(uploadFile);
                    context.ExecuteQuery();

                    Console.WriteLine("File uploaded successfully.");
                    Console.ReadLine();
                }
            }
        }

    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error: {ex.Message}");
    }
}

async Task CheckPermissionAndAddReadPerm(ClientContext clientContext, string fileUrl, string user)
{
    try
    {
        // Get the file
        var file = clientContext.Web.GetFileByServerRelativeUrl(fileUrl);
        clientContext.Load(file, f => f.ListItemAllFields);
        clientContext.ExecuteQuery();

        var listItem = file.ListItemAllFields;

        // Ensure user
        var targetUser = clientContext.Web.EnsureUser(user);
        clientContext.Load(targetUser);
        clientContext.ExecuteQuery();

        // Check if the item has unique permissions
        clientContext.Load(listItem, i => i.HasUniqueRoleAssignments);
        clientContext.ExecuteQuery();

        // Check user's current permissions
        var roleAssignments = listItem.RoleAssignments;
        clientContext.Load(roleAssignments, roles => roles.Include(role => role.Member, role => role.RoleDefinitionBindings));
        clientContext.ExecuteQuery();

        bool userHasPermission = false;
        foreach (var roleAssignment in roleAssignments)
        {
            clientContext.Load(roleAssignment.Member);
            clientContext.ExecuteQuery();

            // Check if the user directly has permission
            if (roleAssignment.Member.PrincipalType == PrincipalType.User &&
                roleAssignment.Member.LoginName.Equals(targetUser.LoginName, StringComparison.OrdinalIgnoreCase))
            {
                userHasPermission = roleAssignment.RoleDefinitionBindings.Any(binding => binding.Name != "Limited Access");
            }

            // Check if the user is part of a group that has permissions
            if (!userHasPermission && roleAssignment.Member.PrincipalType == PrincipalType.SharePointGroup)
            {
                var group = (Group)roleAssignment.Member;
                clientContext.Load(group.Users);
                clientContext.ExecuteQuery();

                userHasPermission = group.Users.Any(usr => usr.LoginName.Equals(targetUser.LoginName, StringComparison.OrdinalIgnoreCase));
            }

            //TODO: Consider to check AD user membership
            //if (!userHasPermission && roleAssignment.Member.PrincipalType == PrincipalType.SecurityGroup)
            //{
            //   
            //}

            if (userHasPermission)
            {
                break;
            }
        }

        if (!userHasPermission)
        {
            if (!listItem.HasUniqueRoleAssignments)
            {
                listItem.BreakRoleInheritance(true, false);
                clientContext.ExecuteQuery();
            }

            Console.WriteLine("User does not have permission. Adding Read permission...");

            var readRoleDef = clientContext.Web.RoleDefinitions.GetByName("Read");
            var roleBindings = new RoleDefinitionBindingCollection(clientContext) { readRoleDef };

            listItem.RoleAssignments.Add(targetUser, roleBindings);
            clientContext.ExecuteQuery();

            Console.WriteLine("Read permission added successfully.");
        }
        else
        {
            Console.WriteLine("User already has permissions on the item.");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error: {ex.Message}");
    }
}

