//Test.aspx
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Test.aspx.cs" Inherits="WOPITest.Test" %>

<!DOCTYPE html>
<html>
<head>
    <title>Simple WOPI Host</title>
</head>
<body>
    <form id="form1" runat="server">
        <h2>Upload a Word Document (.docx)</h2>
        <asp:FileUpload ID="FileUpload1" runat="server" />
        <asp:Button ID="UploadButton" runat="server" Text="Upload" OnClick="UploadButton_Click" />
        <br /><br />
        <asp:Literal ID="FileLink" runat="server" Mode="PassThrough"></asp:Literal>
        <br /><br />
        <iframe id="wopiFrame" width="100%" height="800px" runat="server" visible="false"></iframe>
    </form>
</body>
</html>


//Test.aspx.cs
namespace WOPITest
{
    public partial class Test : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void UploadButton_Click(object sender, EventArgs e)
        {
            if (FileUpload1.HasFile && FileUpload1.FileName.EndsWith(".docx"))
            {
                string fileId = Guid.NewGuid().ToString();
                string savePath = Server.MapPath("~/App_Data/" + fileId + ".docx");
                FileUpload1.SaveAs(savePath);

                string wopiSrc = HttpUtility.UrlEncode($"{Request.Url.GetLeftPart(UriPartial.Authority)}/wopi/files/{fileId}");
                string oosEditUrl = $"https://oos.contonso.com/we/wordeditorframe.aspx?WOPISrc={wopiSrc}";

                string oosUrl = $"https://oos.contonso.com/wv/wordviewerframe.aspx?WOPISrc={wopiSrc}";


                FileLink.Text = $"<a href=\"{oosEditUrl}\" target=\"_blank\">Open in Office Online (Edit)</a>";
                wopiFrame.Attributes["src"] = oosUrl;
                wopiFrame.Visible = true;
            }
        }
    }
}

//WopiHandler.ashx
<%@ WebHandler Language="C#" CodeBehind="WopiHandler.ashx.cs" Class="WOPITest.WopiHandler" %>

//WopiHandler.ashx.cs
public class WopiHandler : IHttpHandler
{
    public void ProcessRequest(HttpContext context)
    {
        var segments = context.Request.Url.AbsolutePath.Split('/');
        bool isContentsRequest = context.Request.Url.AbsolutePath.EndsWith("/contents", StringComparison.OrdinalIgnoreCase);
        string fileId = isContentsRequest ? segments[segments.Length - 2] : segments[segments.Length - 1];
        string filePath = context.Server.MapPath("~/App_Data/" + fileId + ".docx");

        if (context.Request.HttpMethod == "GET")
        {
            if (isContentsRequest)
            {
                context.Response.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                context.Response.WriteFile(filePath);
            }
            else
            {
                var json = new JavaScriptSerializer().Serialize(new
                {
                    BaseFileName = fileId + ".docx",
                    Size = new FileInfo(filePath).Length,
                    OwnerId = "admin",
                    UserId = "admin",
                    Version = "1",
                    SupportsUpdate = true,
                    UserCanWrite = true
                });
                context.Response.ContentType = "application/json";
                context.Response.Write(json);
            }
        }
        else if (context.Request.HttpMethod == "POST")
        {
            string overrideHeader = context.Request.Headers["X-WOPI-Override"];

            if (overrideHeader == "LOCK" || overrideHeader == "REFRESH_LOCK")
            {
                context.Response.AddHeader("X-WOPI-Lock", context.Request.Headers["X-WOPI-Lock"]);
                context.Response.StatusCode = 200;
                return;
            }

            if (overrideHeader == "UNLOCK")
            {
                context.Response.StatusCode = 200;
                return;
            }

            if (isContentsRequest)
            {
                using (var fs = File.Create(filePath))
                {
                    context.Request.InputStream.CopyTo(fs);
                }
                context.Response.StatusCode = 200;
                return;
            }

            context.Response.StatusCode = 501;
        }
        else
        {
            context.Response.StatusCode = 501;
        }
    }

    public bool IsReusable => false;
}

//web.config
	<system.webServer>
		<handlers>
			<add name="WopiHandler" path="wopi/*" verb="*" type="WOPITest.WopiHandler, WOPITest" resourceType="Unspecified" />
		</handlers>
	</system.webServer>


/*Troubleshoot
//Check this URL in browser: https://yourapp.local/wopi/files/{fileId}
Make sure it returns valid JSON like:
{
  "BaseFileName": "abc.docx",
  "Size": 12345,
  "OwnerId": "admin",
  "UserId": "admin",
  "Version": "1",
  "SupportsUpdate": true,
  "UserCanWrite": true
}

Also ensure content-type: Content-Type: application/json

GET https://yourapp.local/wopi/files/{fileId}/contents
This URL must: Return the correct .docx file content

HTTPS is required by OOS (unless started with -AllowHttp)

 Recommended File Structure in Visual Studio:
 /WopiHostWebForms/
│
├── Test.aspx
├── Test.aspx.cs
├── WopiHandler.ashx
├── WopiHandler.ashx.cs (optional, if you split code-behind)
├── Web.config
└── /App_Data/
    └── (Uploaded .docx files go here)

Also check Office Online Server can reach https://yourapp.local/ and vice versa
    
*/
