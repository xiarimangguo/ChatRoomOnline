<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<meta http-equiv="refresh" content="1">
<title>List</title>
<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
Inviter = Request.QueryString("Inviter")
Partner = Request.QueryString("Partner")
Set Dbc = Server.CreateObject("Adodb.Connection")
Ctr = "provider=sqloledb;server=den1.mssql8.gear.host;database=cooles;uid=cooles;pwd=Rs5A9K~?JWw2;"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Show_Usr = "SELECT UserName FROM Users WHERE UserNums = "&Partner
Rs.Open Show_Usr,Ctr,1,1
Unm_F = Session("UserName")
Unm_S = Rs(0)
Rs.Close
Response.Write("<p style='font-size:1em;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>")
If DateDiff("s",Application(Unm_F&"_LT"),Now) <= 5 Then
Application(Unm_F&"_ST") = "Online"
Else
Application(Unm_F&"_ST") = "Offline"
End If
If Application(Unm_F&"_ST") = "Online" Then
Color_ST_s = "green"
Else
Color_ST_s = "red"
End If
Response.Write("<p style='font-size:1em;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<script type='text/javascript'>document.write(decodeURIComponent('"&Unm_F&"'));</script>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b style='color:"&Color_ST_s&";'>"&Application(Unm_F&"_ST")&"</b></p>")
If DateDiff("s",Application(Unm_S&"_LT"),Now) <= 5 Then
Application(Unm_S&"_ST") = "Online"
Else
Application(Unm_S&"_ST") = "Offline"
End If
If Application(Unm_S&"_ST") = "Online" Then
Color_ST_s = "green"
Else
Color_ST_s = "red"
End If
Response.Write("<p style='font-size:1em;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<script type='text/javascript'>document.write(decodeURIComponent('"&Unm_S&"'));</script>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b style='color:"&Color_ST_s&";'>"&Application(Unm_S&"_ST")&"</b></p>")
Set Rs = Nothing
Set Dbc = Nothing
%>