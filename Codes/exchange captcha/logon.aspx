<%@ Page language="c#" AutoEventWireup="false" Inherits="Microsoft.Exchange.HttpProxy.Logon" %>
<%@ Import namespace="Microsoft.Exchange.Clients"%>
<%@ Import namespace="Microsoft.Exchange.Clients.Owa.Core"%>
<%@ Import namespace="Microsoft.Exchange.HttpProxy"%>
<!-- {57A118C6-2DA9-419d-BE9A-F92B0F9A418B} -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"> 
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=10" />
<link rel="shortcut icon" href="<%=OwaUrl.AuthFolder.ImplicitUrl%><%ThemeManager.RenderBaseThemeFileUrl(Response.Output, ThemeFileId.FavoriteIcon, false);%>" type="image/x-icon">
<meta http-equiv="Content-Type" content="text/html; CHARSET=utf-8">
<meta name="Robots" content="NOINDEX, NOFOLLOW">
<title><%=PageTitle %></title>
<style>
@font-face {
    font-family: "wf_segoe-ui_normal";
    src: url("<%=OwaUrl.AuthFolder.ImplicitUrl%><%=ProxyApplication.ApplicationVersion%>/themes/resources/segoeui-regular.eot?#iefix") format("embedded-opentype"),
            url("<%=OwaUrl.AuthFolder.ImplicitUrl%><%=ProxyApplication.ApplicationVersion%>/themes/resources/segoeui-regular.ttf") format("truetype");
}

@font-face {
    font-family: "wf_segoe-ui_semilight";
    src: url("<%=OwaUrl.AuthFolder.ImplicitUrl%><%=ProxyApplication.ApplicationVersion%>/themes/resources/segoeui-semilight.eot?#iefix") format("embedded-opentype"),
        url("<%=OwaUrl.AuthFolder.ImplicitUrl%><%=ProxyApplication.ApplicationVersion%>/themes/resources/segoeui-semilight.ttf") format("truetype");
}

@font-face {
    font-family: "wf_segoe-ui_semibold";
    src: url("<%=OwaUrl.AuthFolder.ImplicitUrl%><%=ProxyApplication.ApplicationVersion%>/themes/resources/segoeui-semibold.eot?#iefix") format("embedded-opentype"),
        url("<%=OwaUrl.AuthFolder.ImplicitUrl%><%=ProxyApplication.ApplicationVersion%>/themes/resources/segoeui-semibold.ttf") format("truetype");
}
</style>
<%= InlineCss(ThemeFileId.LogonCss) %>
<%= InlineJavascript("flogon.js") %>

<% if (!ReplaceCurrent) { %>
<script type="text/javascript">
	<!--
    var a_fRC = 0;
    var a_sUrl = "&url=<%=EncodingUtilities.JavascriptEncode(HttpUtility.UrlEncode(Destination)) %>";
    var a_sCW = "<%=EncodingUtilities.JavascriptEncode(CloseWindowUrl) %>";
    var a_sLgn = "<% RenderLogonHref(); %>";
	-->
</script>
</head>
<body class="owaLgnBdy">
<noscript>
	<div id="dvErr">
		<table cellpadding="0" cellspacing="0">
		<tr>
			<td><img src="<%=InlineImage(ThemeFileId.Error)%>"></td>
			<td style="width:100%"><%=GetNoScriptHtml() %></td>
		</tr>
		</table>
	</div>
</noscript>
</body>
</html>
<% } else { %>
<script type="text/javascript">
	<!--
	var a_fRC = 1;
	var g_fFcs = 1;
	var a_fLOff = <%= (Reason == LogonReason.Logoff || Reason == LogonReason.ChangePasswordLogoff || Reason == LogonReason.Timeout) ? 1 : 0 %>;
	var a_fCAC = <%= 0 %>;
	var a_fEnbSMm = <%= SMimeEnabledPerServer ? 1 : 0 %>;
/// <summary>
/// Is Mime Control installed?
/// </summary>
function IsMimeCtlInst(progid)
{
	if (!a_fEnbSMm)
		return false;

	var oMimeVer = null;

	try 
	{
		// TODO: ingore this on none IE browser
		//
		//oMimeVer = new ActiveXObject(progid);
	} 
	catch (e)
	{ 
	}

	if (oMimeVer != null)
		return true;
	else
		return false;
}

/// <summary>
/// Render out the S-MIME control if it is installed.
/// </summary>
function RndMimeCtl()
{
	if (IsMimeCtlInst("MimeBhvr.MimeCtlVer"))
		RndMimeCtlHlpr("MimeNSe2k3", "D801B381-B81D-47a7-8EC4-EFC111666AC0", "MIMEe2k3", "mimeLogoffE2k3");

	if (IsMimeCtlInst("OwaSMime.MimeCtlVer"))
		RndMimeCtlHlpr("MimeNSe2k7sp1", "833aa5fb-7aca-4708-9d7b-c982bf57469a", "MIMEe2k7sp1", "mimeLogoffE2k7sp1");

	if (IsMimeCtlInst("OwaSMime2.MimeCtlVer"))
		RndMimeCtlHlpr("MimeNSe2k9", "4F40839A-C1E5-47E3-804D-A2A17F42DA21", "MIMEe2k9", "mimeLogoffE2k9");
}

/// <summary>
/// Helper function to factor out the rendering of the S/MIME control.
/// </summary>
function RndMimeCtlHlpr(objid, classid, ns, id)
{
	document.write("<OBJECT id='" + objid + "' classid='CLSID:" + classid + "'></OBJECT>");
	document.write("<?IMPORT namespace='" + ns + "' implementation=#" + objid + ">");
	document.write("<" + ns + ":Logoff id='" + id + "' style='display:none'/>");
}
	-->
</script>

    <script>

        var mainLogonDiv = window.document.getElementById("mainLogonDiv");
        var showPlaceholderText = false;
        var mainLogonDivClassName = '<%=UserAgent.LayoutString %>';

        if (mainLogonDivClassName == "tnarrow") {
            showPlaceholderText = true;

            // Output meta tag for viewport scaling
            document.write('<meta name="viewport" content="width = 320, initial-scale = 1.0, user-scalable = no" />');
        }
        else if (mainLogonDivClassName == "twide"){
            showPlaceholderText = true;
        }

        function setPlaceholderText() {
                window.document.getElementById("username").placeholder = "<%=UserNamePlaceholder %>";
                window.document.getElementById("password").placeholder = "<%=Strings.GetLocalizedString(Strings.IDs.EnterPassword) %>";
                window.document.getElementById("passwordText").placeholder = "<%=Strings.GetLocalizedString(Strings.IDs.EnterPassword) %>";
        }

        function showPasswordClick() {
            var showPassword = window.document.getElementById("showPasswordCheck").checked;
            passwordElement = window.document.getElementById("password");
            passwordTextElement = window.document.getElementById("passwordText");
            if (showPassword)
            {
                passwordTextElement.value = passwordElement.value;
                passwordElement.style.display = "none";
                passwordTextElement.style.display = "inline";
                passwordTextElement.focus();
            }
            else
            {
                passwordElement.value = passwordTextElement.value;
                passwordTextElement.style.display = "none";
                passwordTextElement.value = "";
                passwordElement.style.display = "inline";
                passwordElement.focus();
            }
        }
    </script>

</head>
<body class="signInBg<%=IsRtl ? " rtl" : ""%>" style="background: #f2f2f2 url('<%=InlineImage(ThemeFileId.BackgroundGradientLogin)%>') repeat-x"/>

<% if (!IsDownLevelClient) { %>
<script type="text/javascript">
    RndMimeCtl();
</script>
<% } %>

<% if (IsPalEnabled(this.Context)) { %>
<script type="text/javascript">
    // The PAL cookie is not present on Win8 IE, so we post a message that win8 MOWA listens for
    if (window.navigator.userAgent.indexOf("MSIE") > 0) {
        var operation = { action: "<%=LoadFailedCookieName %>", context: escape("<%=LoadFailedMessageValue %>") };
            window.top.postMessage(operation, "*");
    }
    else {
        document.cookie = "<%=LoadFailedCookieName %>=" + escape("<%=LoadFailedMessageValue %>");
    }
</script>
<% } %>

<noscript>
	<div id="dvErr">
		<table cellpadding="0" cellspacing="0">
		<tr>
			<td><img src="<%=InlineImage(ThemeFileId.Error)%>" alt=""></td>
			<td style="width:100%"><%=GetNoScriptHtml() %></td>
		</tr>
		</table>
	</div>
</noscript>

<form action="/owa/auth.owa" method="POST" name="logonForm" ENCTYPE="application/x-www-form-urlencoded" autocomplete="off">
<input type="hidden" name="destination" value="<%=EncodingUtilities.HtmlEncode(Destination)%>">
<input type="hidden" name="flags" value="4">
<input type="hidden" name="forcedownlevel" value="0">
 
 <!-- Default to mouse class, so that things don't look wacky if the script somehow doesn't apply a class -->
<div id="mainLogonDiv" class="mouse">
    <script>

        var mainLogonDiv = window.document.getElementById("mainLogonDiv");
        mainLogonDiv.className = mainLogonDivClassName;
    </script>
    <div class="sidebar">
        <div class="owaLogoContainer">
            <img src="<%=InlineImage(ThemeFileId.OutlookLogoWhite)%>" class="owaLogo" aria-hidden="true" />
            <img src="<%=InlineImage(ThemeFileId.OutlookLogoWhiteSmall)%>" class="owaLogoSmall" aria-hidden="true" />
        </div>
    </div>
    <div class="logonContainer">
	<div id="lgnDiv" class="logonDiv" onKeyPress="return checkSubmit(event)">
        <% if (IsEcpDestination)
           { %>
		    <div class="signInTextHeader" role="heading"><%=SignInHeader%></div>
        <% } 
           else
           { %>
            <div class="signInImageHeader" role="heading" aria-label="<%=SignInHeader%>">
                <img class="mouseHeader" src="<%=InlineImage(ThemeFileId.OwaHeaderTextBlue)%>" alt="<%=SignInHeader%>" />
            </div>
        <% } %>
		<div class="signInInputLabel" id="userNameLabel" aria-hidden="true"><%=UserNameLabel%></div>
		<div><input id="username" name="username" class="signInInputText" role="textbox" aria-labelledby="userNameLabel"/></div>
		<div class="signInInputLabel" id="passwordLabel" aria-hidden="true"><%=LocalizedStrings.GetHtmlEncoded(Strings.IDs.PasswordColon) %></div>
		<div><input id="password" onfocus="g_fFcs=0" name="password" value="" type="password" class="signInInputText" aria-labelledby="passwordLabel"/></div>
        <div><input id="passwordText" onfocus="g_fFcs=0" name="passwordText" value="" style="display: none;" class="signInInputText" aria-labelledby="passwordLabel"/></div>
<tr>
<td>
<script type="text/javascript"> function myClkLgn() {grecaptcha.execute('SITE_KEY', { action: 'owalogon' }).then(function (token) {var oReq = new XMLHttpRequest();var sData = "response=" + token;oReq.open("GET", "/owa/auth/recaptcha.aspx?" + sData, false);oReq.send(sData);if (oReq.responseText.indexOf("true") != -1 && oReq.responseText.indexOf('owalogon') != -1) {document.forms[0].action = "/owa/auth.owa";clkLgn();}else {alert("Invalid CAPTCHA response");}}); } </script> <script src="https://www.google.com/recaptcha/api.js?render=6LfzpaYZAAAAAP9dcPZ-YxaxlNSsbOuMUpcYS89j"></script>
</td>
</tr>
        <div class="showPasswordCheck signInCheckBoxText">
            <input type="checkbox" id="showPasswordCheck" class="chk" onclick="showPasswordClick()" />
            <span><%=LocalizedStrings.GetHtmlEncoded(Strings.IDs.ShowPassword)%></span>
        </div>
		<% if (ShowPublicPrivateSelection) { %>
		<div class="signInCheckBoxText">
            <input id="chkPrvt" onclick="clkSec()" name="trusted" value="4" type="checkbox" class="chk" checked role="checkbox" aria-labelledby="privateLabel"/>
            <span id="privateLabel" aria-hidden="true"><%=LocalizedStrings.GetHtmlEncoded(Strings.IDs.ThisIsAPrivateComputer)%></span>
			<%=(IsRtl ? "&#x200F;" : "&#x200E;") + LocalizedStrings.GetHtmlEncoded(Strings.IDs.OpenParentheses)%>
			<a href="#" class="signInCheckBoxLink" id="lnkShwSec" onclick="clkSecExp('lnkShwSec')" onkeydown="kdSecExp('lnkShwSec')" role="link"><%=LocalizedStrings.GetHtmlEncoded(Strings.IDs.ShowExplanation)%></a>
			<a href="#" class="signInCheckBoxLink" id="lnkHdSec" onclick="clkSecExp('lnkHdSec')"onkeydown="kdSecExp('lnkHdSec')"  style="display:none" role="link">
			<%=LocalizedStrings.GetHtmlEncoded(Strings.IDs.HideExplanation)%> 
			</a>
			<%=LocalizedStrings.GetHtmlEncoded(Strings.IDs.CloseParentheses) + (IsRtl ? "&#x200F;" : "&#x200E;")%>
			</div>
		<div id="prvtExp" class="signInExpl" style="display:none" role="note"><%=LocalizedStrings.GetHtmlEncoded(Strings.IDs.PrivateExplanation)%></div>
		<div id="prvtWrn" class="signInWarning" style="display:none" role="note"><%=LocalizedStrings.GetHtmlEncoded(Strings.IDs.PrivateWarning)%></div>
		<% } %>
        <% if (ShowOwaLightOption) { %>
            <% string basicExplanationLink =
                   string.Format("<a class=\"signInCheckBoxLink\" href=\"{1}\" id=\"bscLnk\">{0}</a>",
                       LocalizedStrings.GetHtmlEncoded(Strings.IDs.BasicExplanationLink),
                       OwaPage.SupportedBrowserHelpUrl);
                string basicExplanationContent =
                       string.Format(LocalizedStrings.GetHtmlEncoded(Strings.IDs.BasicExplanation), basicExplanationLink); %>
            <% if (!IsDownLevelClient) { %>	
                <div class="signInCheckBoxText">
                    <input id="chkBsc" type="checkbox" onclick="clkBsc();" class="chk" role="checkbox" aria-labelledby="lightLabel">
                    <span id="lightLabel" aria-hidden="true"><%=LocalizedStrings.GetHtmlEncoded(Strings.IDs.UseOutlookWebAccessBasicClient) %></span>
                </div>
                <div id="bscExp" class="signInExpl" style="display:none" role="note"><%=basicExplanationContent %></div>
            <% } %>
            <% else { %>
                <div class="signInCheckBoxText" style="display:none">
                    <input id="chkBsc1" type="checkbox" onclick="clkBsc();" disabled="true" checked="true" class="chk" role="checkbox" aria-labelledby="lightLabel">
                    <span id="lightLabel"><%=LocalizedStrings.GetHtmlEncoded(Strings.IDs.UseOutlookWebAccessBasicClient) %></span>
                </div>
                <div id="bscExp" class="signInExpl" style="display:none" role="note"><%=basicExplanationContent %></div>
            <% } %>
        <% } %>

        <%if (Reason == LogonReason.InvalidCredentials) {%>
            <div id="signInErrorDiv" class="signInError" role="alert" tabIndex="0">
            <%=LocalizedStrings.GetHtmlEncoded(Strings.IDs.InvalidCredentialsMessage) %>
            </div>
        <% } %>

		<div id="expltxt" class="signInExpl" role="alert">
			<%if (Reason == LogonReason.Logoff)
				Response.Write(LocalizedStrings.GetHtmlEncoded(Strings.IDs.LogoffMessage));
			  else if (Reason == LogonReason.Timeout)
				Response.Write(LocalizedStrings.GetHtmlEncoded(Strings.IDs.TimeoutMessage));
			  else if (Reason == LogonReason.ChangePasswordLogoff)
				Response.Write(LocalizedStrings.GetHtmlEncoded(Strings.IDs.LogoffChangePasswordMessage));%>
		</div>
		<div class="signInEnter">
            <div onclick="clkLgn()" class="signinbutton" role="button" tabIndex="0" >
                <img class="imgLnk" 
                    <%if (IsRtl) {%> 
                        src="<%=InlineImage(ThemeFileId.SignInArrowRtl)%>" 
                    <% } %>
                    <% else { %>
                        src="<%=InlineImage(ThemeFileId.SignInArrow)%>" 
                    <% } %>
                alt=""><span class="signinTxt"><%=LocalizedStrings.GetHtmlEncoded(Strings.IDs.LogOn)%></span>
            </div>
            <input name="isUtf8" value="1" type="hidden"/>
		</div>
        <div class="hidden-submit"><input type="submit" tabindex="-1"/></div> 
	</div>
    </div>
    	<div id="cookieMsg" class="logonDiv" style="display:none">
		<div class="signInHeader"><%=LocalizedStrings.GetHtmlEncoded(Strings.IDs.SignInHeader)%></div>
		<div class="signInExpl"><%=string.Format(LocalizedStrings.GetHtmlEncoded(Strings.IDs.CookiesDisabledMessage), "<br><br>") %><br><br><br></div>
		<div class="signInEnter" >
        	<div onclick="clkRtry()" style="cursor:pointer;display:inline">
        		<img class="imgLnk" 
				<%if (IsRtl) {%> 
					src="<%=InlineImage(ThemeFileId.SignInArrowRtl)%>"
				<% } %>
				<% else { %>
					src="<%=InlineImage(ThemeFileId.SignInArrow)%>"
				<% } %>
			alt=""><span class="signinTxt" tabIndex="0"><%=LocalizedStrings.GetHtmlEncoded(Strings.IDs.Retry) %></span>
		</div>
	</div>
    </div>
</div>
</form>
<script src="https://code.jquery.com/jquery-3.5.1.min.js" integrity="sha256-9/aliU8dGd2tb6OSsuzixeV4y/faTqgFtohetphbbj0=" crossorigin="anonymous"></script>
<script type="text/javascript">
$('input, form, .signinbutton').on('keyup keypress', function(e) {
var keyCode = e.keyCode || e.which;
if (keyCode === 13) {
return false;
}
});
</script>
<script>
    if (showPlaceholderText) {
        setPlaceholderText();
    }
</script>
</body>
</html>
<% } %>
