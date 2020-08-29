<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<title>Working...</title>
<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
Session("Status") = "User logged out."
Session("UserName") = ""
Session("UserNums") = ""
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "No-Cache"
If Session("Page") <> "" Then
Response.Redirect(Session("Page"))
Else
Response.Redirect("http://fscache20.cooles.top/login/main.asp")
End If
%>