<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
Dim Unm,Pwd,rPwd,Lge
Unm = Server.URLEncode(Request.Form("Unm"))
Pwd = Request.Form("Pwd")
rPwd = Request.Form("rPwd")
Lge = Server.URLEncode(Request.Form("Lge"))
If Unm <> "" Then
If Pwd <> "" Then
If rPwd <> Pwd Then
Response.Write "Inconsistent PassWord. <a href='javascript:history.go(-1);'>Retry</a>"
Response.End
Else
End If
Else
Response.Write "Invalid PassWord. <a href='javascript:history.go(-1);'>Retry</a>"
Response.End
End If
Else
Response.Write "Invalid UserName. <a href='javascript:history.go(-1);'>Retry</a>"
Response.End
End If
Set Dbc = Server.CreateObject("Adodb.Connection")
Ctr = "provider=sqloledb;server=den1.mssql8.gear.host;database=cooles;uid=cooles;pwd=Rs5A9K~?JWw2;"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Clk_Unm = "select UserName from Users where UserName = '"&Unm&"'"
Rs.Open Clk_Unm,Ctr,1,1
If Rs.Eof Then
Rs.Close
Rgs = "insert into Users values('"&Unm&"','"&Pwd&"')"
Rs.Open Rgs,Ctr,1,3
Set Rs2 = Server.CreateObject("Adodb.RecordSet")
Gnm = "select UserNums from Users where UserName = '"&Unm&"'"
Rs2.Open Gnm,Ctr,1,1
Num = Rs2(0)
Rs2.Close
Crt_Lst = "CREATE TABLE Verifies_"&Num&"(Orders int identity(1,1),UserNums int,UserName varchar(255),Message varchar(255),States varchar(255));CREATE TABLE Friends_"&Num&"(Orders int identity(1,1),UserNums int,UserName varchar(255),States varchar(255))"
Rs2.Open Crt_Lst,Ctr,1,3
If Lge <> "" Then
Application(Session("UserName")&"_LG") = Lge
Response.Cookies("Lge") = Lge
Response.Cookies("Lge").Expires = Date() + 28
Response.Cookies("Lge").Path = "/login/"
Else
End If
%>
<%
Lge = Request.Cookies("Lge")
If Lge = "" Then
Response.Cookies("Lge") = "English"
Response.Cookies("Lge").Expires = Date() + 28
Response.Cookies("Lge").Path = "/login/"
End If
Select Case Lge
Case "English"
Lcd = "en-us"
Case "%E7%AE%80%E4%BD%93%E4%B8%AD%E6%96%87"
Lcd = "zh-cn"
Case "%E7%B9%81%E9%AB%94%E4%B8%AD%E6%96%87"
Lcd = "zh-hk"
Case "%D1%80%D1%83%D1%81%D1%81%D0%BA%D0%B8%D0%B9"
Lcd = "ru-ru"
Case "%E6%97%A5%E6%9C%AC%E8%AA%9E"
Lcd = "ja-jp"
Case "fran%C3%A7ais"
Lcd = "fr-fr"
Case "Deutsche"
Lcd = "de-de"
Case Else
Lcd = "en-us"
Response.Cookies("Lge") = "English"
Response.Cookies("Lge").Expires = Date() + 28
Response.Cookies("Lge").Path = "/login/"
End Select
%>
<%
Response.Write "Completed your registration. <a href='http://fscache20.cooles.top/login/login.html?lge="&Lcd&"'>Go back to login</a>"
Response.End
Rs2.Close
Rs.Close
Else
Rs.Close
Response.Write "UserName '<script type='text/javascript'>document.write(decodeURIComponent('"&Unm&"'));</script>' is already registered. <a href='javascript:history.go(-1);'>Retry</a>"
Response.End
End If
Set Rs2 = Nothing
Set Rs = Nothing
Set Dbc = Nothing
%>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<title>Working...</title>