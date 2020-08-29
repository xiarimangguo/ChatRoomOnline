<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
Application(Session("UserName")&"_LT") = Now
%>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<meta http-equiv="refresh" content="1">
<title>ChatBord</title>
<body onload="window.scrollTo(0,document.body.scrollHeight);">
