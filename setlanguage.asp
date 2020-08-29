<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
Lge = Server.URLEncode(Request.Form("Lge"))
Application(Session("UserName")&"_LG") = Lge
Response.Cookies("Lge") = Lge
Response.Cookies("Lge").Expires = Date() + 28
Response.Cookies("Lge").Path = "/login/"
%>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<title>Working...</title>