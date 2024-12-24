using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using ListItem = Microsoft.SharePoint.Client.ListItem;
using View = Microsoft.SharePoint.Client.View;

namespace SpCSOM
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string siteUrl = "http://contonso.com/";
            string username = "username";
            string pass = "password";
            string domain = "domain";
            string fileurl = "/Shared%20Documents/test.docx";
            string filedir = "/Shared Documents";
            string siteColumnName = "Test_Column";
        

            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                clientContext.Credentials = new System.Net.NetworkCredential(username, pass, domain);
                Web web = clientContext.Web;
                clientContext.Load(web, w => w.ServerRelativeUrl, w => w.Url);
                Console.WriteLine("Conneciton done successfully");

                /*dosya checking / checkout */
                //CheckinCheckout(clientContext, fileurl);

                /*Add Document to specified field*/
                //AddFileToSpFolder(clientContext, filedir);

                /*Adding site column to library*/
                //AddSiteColumnToListIfMissing(clientContext, fileurl, siteColumnName);

                /*Column value edit*/
                //UpdateItemProperty(clientContext, fileurl, "https://www.google.com");


                /*Creating file based on target file's header*/
                //DocHeaderUpdate(clientContext, fileurl,filedir);

                /*permission check and add*/
                //CheckPermissionAndAddReadPerm(clientContext, fileurl, "testuser1");
            }
        }

        static void DocHeaderUpdate(ClientContext clientContext, string fileUrl, string filedir)
        {
            Microsoft.SharePoint.Client.File file = clientContext.Web.GetFileByServerRelativeUrl(fileUrl);
            clientContext.Load(file);
            clientContext.ExecuteQuery();

            ExtractHeaderAndCreateNewDoc(clientContext, fileUrl, filedir);
        }

        private static void ExtractHeaderAndCreateNewDoc(ClientContext clientContext, string fileUrl, string filedir)
        {
            FileInformation fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, fileUrl);

            using (MemoryStream sourceDocStream = new MemoryStream())
            {
                fileInfo.Stream.CopyTo(sourceDocStream);
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

                        Folder folder = clientContext.Web.GetFolderByServerRelativeUrl(filedir);
                        clientContext.Load(folder);
                        clientContext.ExecuteQuery();
                                              
                        FileCreationInformation newFile = new FileCreationInformation
                        {
                            ContentStream = newDocStream,
                            Url = "newdoc.docx",
                            Overwrite = true
                        };

                        Microsoft.SharePoint.Client.File uploadFile = folder.Files.Add(newFile);
                        clientContext.Load(uploadFile);
                        clientContext.ExecuteQuery();

                        Console.WriteLine("File uploaded successfully.");
                        Console.ReadLine();


                    }
                }
            }
        }

        
        static void CheckPermissionAndAddReadPerm(ClientContext clientContext, string fileUrl, string user)
        {
            RoleDefinition readerDef = clientContext.Web.RoleDefinitions.GetByType(RoleType.Reader);
            RoleDefinitionBindingCollection readOnlyBinding = new RoleDefinitionBindingCollection(clientContext);
            readOnlyBinding.Add(readerDef);

            Microsoft.SharePoint.Client.File file = clientContext.Web.GetFileByServerRelativeUrl(fileUrl);
            clientContext.Load(file, f => f.ListItemAllFields);
            clientContext.ExecuteQuery();

            User targetUser = clientContext.Web.EnsureUser(user);
            clientContext.Load(targetUser);
            clientContext.ExecuteQuery();

            ListItem item = file.ListItemAllFields;
            clientContext.Load(item, i => i.HasUniqueRoleAssignments);
            clientContext.ExecuteQuery();

            RoleAssignmentCollection roleAssignments = item.RoleAssignments;
            clientContext.Load(roleAssignments, roles => roles.Include(role => role.Member));
            clientContext.ExecuteQuery();

            bool userHasPermission = false;
            foreach (RoleAssignment roleAssignment in roleAssignments)
            {
                Principal member = roleAssignment.Member;
                clientContext.Load(roleAssignment, r => r.RoleDefinitionBindings);
                clientContext.ExecuteQuery();

                if (member.PrincipalType == PrincipalType.User
                    && member.LoginName.Equals(targetUser.LoginName, StringComparison.OrdinalIgnoreCase)
                    && roleAssignment.RoleDefinitionBindings.Where(x => x.Name != "Limited Access").Any())
                {
                    userHasPermission = true;
                    break;
                }

                if (member.PrincipalType == PrincipalType.SharePointGroup)
                {
                    Group group = clientContext.Web.SiteGroups.GetById(member.Id);
                    clientContext.Load(group, g => g.Users);
                    clientContext.ExecuteQuery();

                    foreach (User usr in group.Users)
                    {
                        if (usr.LoginName.Equals(targetUser.LoginName, StringComparison.OrdinalIgnoreCase))
                        {
                            userHasPermission = true;
                            break;
                        }
                    }
                }

                if (userHasPermission)
                    break;
            }

            if (!userHasPermission)
            {
                if (!item.HasUniqueRoleAssignments)
                {
                    item.BreakRoleInheritance(true, false);
                    clientContext.ExecuteQuery();
                }

                Console.WriteLine("User does not have permission. Adding Read permission...");

                RoleDefinitionBindingCollection roleBindings = new RoleDefinitionBindingCollection(clientContext);
                RoleDefinition readPermission = clientContext.Web.RoleDefinitions.GetByName("Read");
                roleBindings.Add(readPermission);

                item.RoleAssignments.Add(targetUser, roleBindings);
                clientContext.ExecuteQuery();

                Console.WriteLine("Read permission added successfully.");
            }
            else
            {
                Console.WriteLine("User already has permissions on the item.");
            }

            Console.ReadLine();
        }

        static void UpdateItemProperty(ClientContext clientContext, string fileUrl, string href)
        {
            Microsoft.SharePoint.Client.File file = clientContext.Web.GetFileByServerRelativeUrl(fileUrl);
            clientContext.Load(file, f => f.ListItemAllFields);
            clientContext.ExecuteQuery();

            if (file.ListItemAllFields.FieldValues.ContainsKey("Test_Column"))
            {
                var fieldValue = new FieldUrlValue
                {
                    Url = href,
                    Description = "Column Value"
                };

                file.ListItemAllFields["Test_Column"] = fieldValue;
                file.ListItemAllFields.Update(); 
                clientContext.ExecuteQuery();
                Console.WriteLine("Custom column updated successfully.");
            }
            else
            {
                Console.WriteLine("The column 'Test_Column' does not exist in the list.");
            }

            Console.ReadLine();
        }

        static void AddFileToSpFolder(ClientContext clientContext, string filedir)
        {
            Folder folder = clientContext.Web.GetFolderByServerRelativeUrl(filedir);
            clientContext.Load(folder);
            clientContext.ExecuteQuery();

            byte[] fileData = System.IO.File.ReadAllBytes(@"c:\log.txt");
            FileCreationInformation newFile = new FileCreationInformation
            {
                ContentStream = new MemoryStream(fileData),                
                Url = "log test.txt",
                Overwrite = true
            };

            Microsoft.SharePoint.Client.File uploadFile = folder.Files.Add(newFile);
            clientContext.Load(uploadFile);
            clientContext.ExecuteQuery();

            Console.WriteLine("File uploaded successfully.");
            Console.ReadLine();
        }

        static void CheckinCheckout(ClientContext clientContext, string fileurl)
        {

            Microsoft.SharePoint.Client.File testfile = clientContext.Web.GetFileByServerRelativeUrl(fileurl);
            clientContext.Load(testfile, f => f.Name, f => f.TimeCreated, f => f.CheckOutType, f => f.CheckedOutByUser);
            clientContext.ExecuteQuery();

            
            if (testfile.CheckOutType == CheckOutType.None)
            {
                testfile.CheckOut();
                clientContext.ExecuteQuery();
            }

            
            if (testfile.CheckOutType != CheckOutType.None && testfile.CheckedOutByUser.LoginName == @"SHAREPOINT\system")
            {
                testfile.UndoCheckOut();
                clientContext.ExecuteQuery();
            }

            Console.WriteLine(testfile.Name + " " + testfile.TimeCreated.ToString("dd.MM.yyyy"));
            Console.ReadLine();
        }

        static void AddSiteColumnToListIfMissing(ClientContext clientContext, string fileUrl, string siteColumnName)
        {
            try
            {
                Microsoft.SharePoint.Client.File file = clientContext.Web.GetFileByServerRelativeUrl(fileUrl);
                clientContext.Load(file, f => f.ListItemAllFields.ParentList.Id);
                clientContext.ExecuteQuery();

                Guid listId = file.ListItemAllFields.ParentList.Id;
                Microsoft.SharePoint.Client.List list = clientContext.Web.Lists.GetById(listId);
                clientContext.Load(list, l => l.Fields, l => l.ContentTypesEnabled);
                clientContext.ExecuteQuery();

                bool columnExists = false;
                foreach (var field in list.Fields)
                {
                    if (field.InternalName == siteColumnName)
                    {
                        columnExists = true;
                        break;
                    }
                }

                if (!columnExists)
                {
                    Console.WriteLine($"Adding site column '{siteColumnName}' to the list.");
                    Microsoft.SharePoint.Client.Field field = clientContext.Web.Fields.GetByInternalNameOrTitle(siteColumnName);
                    clientContext.Load(field);
                    clientContext.ExecuteQuery();

                    list.Fields.Add(field);
                    list.Update();
                    clientContext.ExecuteQuery();
                    Console.WriteLine($"Site column '{siteColumnName}' added to the list.");
                }

                View defaultView = list.DefaultView;
                clientContext.Load(defaultView, v => v.ViewFields);
                clientContext.ExecuteQuery();

                if (!defaultView.ViewFields.Contains(siteColumnName))
                {
                    Console.WriteLine($"Adding site column '{siteColumnName}' to the default view.");
                    defaultView.ViewFields.Add(siteColumnName);
                    defaultView.Update();
                    clientContext.ExecuteQuery();
                    Console.WriteLine($"Site column '{siteColumnName}' added to the default view.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while adding site column: " + ex.Message);
            }
        }
    }
}
