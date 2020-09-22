<div align="center">

## Domain Redirect


</div>

### Description

This is a basic redirection script based on the domain name. It is usefull if you are hosting multiple sites from one IP address.
 
### More Info
 
Insert your site names in the Select Case statement to replace the Your-Site, Your-Next-Site & Default-Site text. You can add to the select case for additional sites.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Darryl Brown](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/darryl-brown.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/darryl-brown-domain-redirect__4-6125/archive/master.zip)





### Source Code

```
<%@ Language=VBScript %>
<% response.buffer = true %>
<%
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~ Script:	| Domain Redirect
'~ Language:	| ASP/VBScript
'~ Version:	| 1.0
'~ By:		| Darryl A. Brown
'~ Contact:	| Brown@LunarTech.com
'~ Copyright:	| Darryl A. Brown
'~ Released:	| April 12, 2000
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~ By using this software, you have agreed to the license
'~ agreement packaged with this program. This script is
'~ provided without warranty.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 'Set up the redirect variables
 Dim strURL		'This is the URL the user typed into the browser
 Dim strSiteName	'This is the domain name
 Dim intEPos		'This is the position of the remaining dot
 strURL = Request.servervariables("HTTP_HOST")
 intURLLen = len(strURL)
 If inStr(1, UCase(strURL), "WWW") > 0 Then
  strURL = Right(strURL, (intURLLen - 4))
 End If
 intEPos = inStr(1, strURL, ".")
 strSiteName = UCase(Mid(strURL, 1, (intEPos - 1)))
 'Insert your site names in the Select Case statement below to
 'replace the Your-Site, Your-Next-Site & Default-Site text.
 'Feel free to add to the select case for additional sites.
 Select Case strSiteName
  Case "Your-Site"
   Response.redirect "http://www.Your-Site.com"
  Case "Your-Next-Site"
   Response.redirect "http://www.Your-Next-Site.com/sitedir"
  Case Else
   Response.redirect "http://www.Default-Site.com/home.html"
 End Select
%>
```

