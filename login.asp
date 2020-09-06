<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<title>Working...</title>
<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
If Session("Status") <> "User logged in." Then
Else
If Session("Page") <> "" Then
Response.Redirect(Session("Page"))
Else
Response.Redirect("http://fscache20.cooles.top/login/main.asp")
End If
End If
Dim Unm,Pwd,Lge
Unm = Server.URLEncode(Request.Form("Unm"))
Pwd = Request.Form("Pwd")
Lge = Server.URLEncode(Request.Form("Lge"))
If Unm <> "" Then
If Pwd <> "" Then
Else
Response.Write "Invalid PassWord. <a href='javascript:history.go(-1);'>Retry</a>"
Response.End
End If
Else
Response.Write "Invalid UserName. <a href='javascript:history.go(-1);'>Retry</a>"
Response.End
End If
Set Dbc = Server.CreateObject("Adodb.Connection")
Ctr = "provider=sqloledb;server=den1.mssql8.gear.host;database=;uid=;pwd=;"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Clk_Unm = "select UserName from Users where UserName = '"&Unm&"'"
Rs.Open Clk_Unm,Ctr,1,1
If Rs.Eof Then
Rs.Close
Response.Write "Wrong UserName. <a href='javascript:history.go(-1);'>Retry</a>"
Response.End
Else
Rs.Close
Clk_Pwd = "select UserName,PassWord from Users where UserName = '"&Unm&"' and PassWord = '"&Pwd&"'"
Rs.Open Clk_Pwd,Ctr,1,1
If Rs.Eof Then
Rs.Close
Response.Write "Wrong PassWord. <a href='javascript:history.go(-1);'>Retry</a>"
Response.End
Else
Rs.Close
Get_Num = "select UserNums from Users where UserName = '"&Unm&"'"
Rs.Open Get_Num,Ctr,1,1
Num = Rs(0)
Rs.Close
Session("Status") = "User logged in."
Session("UserName") = Unm
Session("UserNums") = Num
If Lge <> "" Then
Application(Session("UserName")&"_LG") = Lge
Response.Cookies("Lge") = Lge
Response.Cookies("Lge").Expires = Date() + 28
Response.Cookies("Lge").Path = "/login/"
Else
End If
If Session("Page") <> "" Then
Response.Redirect(Session("Page"))
Else
Response.Redirect("http://fscache20.cooles.top/login/main.asp")
End If
End If
End If
Set Rs = Nothing
Set Dbc = Nothing
%>
