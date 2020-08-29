<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<title>Working...</title>
<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
If Session("Status") = "User logged in." Then
Else
Response.Write "Unauthorized visit. <a href='javascript:history.go(-1);'>Retry</a>"
Response.End
End If
Dim Pwd,Clw
Pwd = Request.Form("Pwd")
Clw = Request.Form("Clw")
If Clw = "I am sure that I will delete myself from this server." Then
If Pwd <> "" Then
Else
Response.Write "Invalid PassWord. <a href='javascript:history.go(-1);'>Retry</a>"
Response.End
End If
Else
Response.Write "Failed deletion. <a href='javascript:history.go(-1);'>Retry</a>"
Response.End
End If
Set Dbc = Server.CreateObject("Adodb.Connection")
Ctr = "provider=sqloledb;server=den1.mssql8.gear.host;database=cooles;uid=cooles;pwd=Rs5A9K~?JWw2;"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Clk_Pwd = "select UserName,PassWord from Users where UserName = '"&Session("UserName")&"' and PassWord = '"&Pwd&"'"
Rs.Open Clk_Pwd,Ctr,1,1
If Rs.Eof Then
Rs.Close
Response.Write "Wrong PassWord. <a href='javascript:history.go(-1);'>Retry</a>"
Response.End
Else
Rs.Close
Del_Usr = "delete Users where UserName = '"&Session("UserName")&"'"
Rs.Open Del_Usr,Ctr,1,3
Response.Write "User deleted. <a href='http://fscache20.cooles.top/login/logout.asp'>Logout</a>"
Response.End
Rs.Close
End If
Set Rs = Nothing
Set Dbc = Nothing
%>