<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<meta http-equiv="refresh" content="1">
<title>List</title>
<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
If Application(Session("UserName")&"_LG") = "" Then
Application(Session("UserName")&"_LG") = "English"
End If
Select Case Application(Session("UserName")&"_LG")
Case "English"
TxtBox_1 = "Online"
TxtBox_2 = "Offline"
TxtBox_3 = ""
TxtBox_4 = ""
TxtBox_5 = ""
TxtBox_6 = ""
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E7%AE%80%E4%BD%93%E4%B8%AD%E6%96%87"
TxtBox_1 = "在线"
TxtBox_2 = "离线"
TxtBox_3 = ""
TxtBox_4 = ""
TxtBox_5 = ""
TxtBox_6 = ""
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E7%B9%81%E9%AB%94%E4%B8%AD%E6%96%87"
TxtBox_1 = "在線"
TxtBox_2 = "離線"
TxtBox_3 = ""
TxtBox_4 = ""
TxtBox_5 = ""
TxtBox_6 = ""
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%D1%80%D1%83%D1%81%D1%81%D0%BA%D0%B8%D0%B9"
TxtBox_1 = "В сети"
TxtBox_2 = "Не в сети"
TxtBox_3 = ""
TxtBox_4 = ""
TxtBox_5 = ""
TxtBox_6 = ""
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E6%97%A5%E6%9C%AC%E8%AA%9E"
TxtBox_1 = "オンライン"
TxtBox_2 = "オフライン"
TxtBox_3 = ""
TxtBox_4 = ""
TxtBox_5 = ""
TxtBox_6 = ""
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "fran%C3%A7ais"
TxtBox_1 = "En ligne"
TxtBox_2 = "Hors ligne"
TxtBox_3 = ""
TxtBox_4 = ""
TxtBox_5 = ""
TxtBox_6 = ""
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "Deutsche"
TxtBox_1 = "Online"
TxtBox_2 = "Offline"
TxtBox_3 = ""
TxtBox_4 = ""
TxtBox_5 = ""
TxtBox_6 = ""
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case Else
TxtBox_1 = "Online"
TxtBox_2 = "Offline"
TxtBox_3 = ""
TxtBox_4 = ""
TxtBox_5 = ""
TxtBox_6 = ""
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
End Select
%>
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
Unm_F_ST = TxtBox_1
Else
Application(Unm_F&"_ST") = "Offline"
Unm_F_ST = TxtBox_2
End If
If Application(Unm_F&"_ST") = "Online" Then
Color_ST_s = "green"
Else
Color_ST_s = "red"
End If
Response.Write("<p style='font-size:1em;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<script type='text/javascript'>document.write(decodeURIComponent('"&Unm_F&"'));</script>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b style='color:"&Color_ST_s&";'>"&Unm_F_ST&"</b></p>")
If DateDiff("s",Application(Unm_S&"_LT"),Now) <= 5 Then
Application(Unm_S&"_ST") = "Online"
Unm_S_ST = TxtBox_1
Else
Application(Unm_S&"_ST") = "Offline"
Unm_S_ST = TxtBox_2
End If
If Application(Unm_S&"_ST") = "Online" Then
Color_ST_s = "green"
Else
Color_ST_s = "red"
End If
Response.Write("<p style='font-size:1em;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<script type='text/javascript'>document.write(decodeURIComponent('"&Unm_S&"'));</script>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b style='color:"&Color_ST_s&";'>"&Unm_S_ST&"</b></p>")
Set Rs = Nothing
Set Dbc = Nothing
%>