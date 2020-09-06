<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
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
<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<meta http-equiv="refresh" content="1">
<title>FriendList</title>
<%
If Application(Session("UserName")&"_LG") = "" Then
Application(Session("UserName")&"_LG") = "English"
End If
Select Case Application(Session("UserName")&"_LG")
Case "English"
TxtBox_1 = "Friends"
TxtBox_2 = "All"
TxtBox_3 = "Verifies"
TxtBox_4 = "Add"
TxtBox_5 = "Chat"
TxtBox_6 = "Delete"
TxtBox_7 = "Online"
TxtBox_8 = "Offline"
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E7%AE%80%E4%BD%93%E4%B8%AD%E6%96%87"
TxtBox_1 = "好友"
TxtBox_2 = "所有"
TxtBox_3 = "验证"
TxtBox_4 = "添加"
TxtBox_5 = "聊天"
TxtBox_6 = "删除"
TxtBox_7 = "在线"
TxtBox_8 = "离线"
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E7%B9%81%E9%AB%94%E4%B8%AD%E6%96%87"
TxtBox_1 = "好友"
TxtBox_2 = "所有"
TxtBox_3 = "驗證"
TxtBox_4 = "添加"
TxtBox_5 = "聊天"
TxtBox_6 = "刪除"
TxtBox_7 = "在線"
TxtBox_8 = "離線"
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%D1%80%D1%83%D1%81%D1%81%D0%BA%D0%B8%D0%B9"
TxtBox_1 = "друзья"
TxtBox_2 = "Bce"
TxtBox_3 = "Проверяет"
TxtBox_4 = "Добавить"
TxtBox_5 = "Чат"
TxtBox_6 = "Удалить"
TxtBox_7 = "В сети"
TxtBox_8 = "Не в сети"
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E6%97%A5%E6%9C%AC%E8%AA%9E"
TxtBox_1 = "フレンズ"
TxtBox_2 = "すべて"
TxtBox_3 = "検証する"
TxtBox_4 = "追加"
TxtBox_5 = "チャット"
TxtBox_6 = "削除する"
TxtBox_7 = "オンライン"
TxtBox_8 = "オフライン"
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "fran%C3%A7ais"
TxtBox_1 = "Copains"
TxtBox_2 = "Tout"
TxtBox_3 = "Vérifie"
TxtBox_4 = "Ajouter"
TxtBox_5 = "Bavarder"
TxtBox_6 = "Supprimer"
TxtBox_7 = "En ligne"
TxtBox_8 = "Hors ligne"
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "Deutsche"
TxtBox_1 = "Freunde"
TxtBox_2 = "Alles"
TxtBox_3 = "Überprüft"
TxtBox_4 = "Hinzufügen"
TxtBox_5 = "Plaudern"
TxtBox_6 = "Löschen"
TxtBox_7 = "Online"
TxtBox_8 = "Offline"
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case Else
TxtBox_1 = "Friends"
TxtBox_2 = "All"
TxtBox_3 = "Verifies"
TxtBox_4 = "Add"
TxtBox_5 = "Chat"
TxtBox_6 = "Delete"
TxtBox_7 = "Online"
TxtBox_8 = "Offline"
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
End Select
%>
<%
Set Dbc = Server.CreateObject("Adodb.Connection")
Ctr = "provider=sqloledb;server=den1.mssql8.gear.host;database=;uid=;pwd=;"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Show_Ols = "SELECT UserName FROM Friends_"&Session("UserNums")&" ORDER BY Orders"
Rs.Open Show_Ols,Ctr,1,1
Onlines = 0
Do Until Rs.Eof
Ur = Rs(0)
If DateDiff("s",Application(Ur&"_LT"),Now) <= 5 Then
Onlines = Onlines + 1
Else
End If
Rs.MoveNext
Loop
Application("Onlines_Frs_"&Session("UserName")) = Onlines
If Application("Onlines_Frs_"&Session("UserName")) > 0 Then
Color_ST = "green"
Else
Color_ST = "red"
End If
Rs.Close
Show_Avn = "SELECT COUNT(Orders) FROM Verifies_"&Session("UserNums")
Rs.Open Show_Avn,Ctr,1,1
Application("Cnt_Vrf_"&Session("UserName")) = Rs(0).Value
If Application("Cnt_Vrf_"&Session("UserName")) > Application("Cnt_Vrs_"&Session("UserName")) Then
Application("New_Vrf_"&Session("UserName")) = Application("Cnt_Vrf_"&Session("UserName")) - Application("Cnt_Vrs_"&Session("UserName"))
Else
Application("New_Vrf_"&Session("UserName")) = 0
End If
Rs.Close
Show_Aun = "SELECT COUNT(UserName) FROM Friends_"&Session("UserNums")
Rs.Open Show_Aun,Ctr,1,1
Application("All_Frs_"&Session("UserName")) = Rs(0).Value
Response.Write("<p style='font-size:1em;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>"&TxtBox_1&"</b>&nbsp;&nbsp;(<b style='color:"&Color_ST&";'>"&Application("Onlines_Frs_"&Session("UserName"))&"</b>/<b>"&Application("All_Frs_"&Session("UserName"))&"</b>)&nbsp;&nbsp;<b>|</b>&nbsp;&nbsp;<b><a href='http://fscache20.cooles.top/login/memberlist.asp' style='color:blue;'>"&TxtBox_2&"</a></b>&nbsp;&nbsp;<b>|</b>&nbsp;&nbsp;<b>-</b>&nbsp;&nbsp;<b><a href='http://fscache20.cooles.top/login/verifylist.asp' style='color:blue;'>"&TxtBox_3&"</a></b><b id='New_Vrf' style='display:none;'></b>&nbsp;&nbsp;<b>-</b>&nbsp;&nbsp;<b><a href='http://fscache20.cooles.top/login/friendadd.html?lge="&Lcd&"' style='color:blue;'>"&TxtBox_4&"</a></b></p><iframe name='nofresh' style='display:none;width:0em;height:0em;'></iframe>")
If Application("Cnt_Vrf_"&Session("UserName")) > Application("Cnt_Vrs_"&Session("UserName")) Then
Response.Write("<script type=""text/javascript"">document.getElementById('New_Vrf').innerHTML='[+<b style=""color:green;"">"&Application("New_Vrf_"&Session("UserName"))&"</b>]';document.getElementById('New_Vrf').style.display='';</script>")
Else
End If
Rs.Close
Show_Usr = "SELECT * FROM Friends_"&Session("UserNums")&" ORDER BY Orders"
Rs.Open Show_Usr,Ctr,1,1
If Rs.Eof Then
Rs.Close
Else
Form_Num = 0
Do Until Rs.Eof
Form_Num = Form_Num + 1
Num = Rs(1)
Urs = Rs(2)
If DateDiff("s",Application(Urs&"_LT"),Now) <= 5 Then
Application(Urs&"_ST") = "Online"
Urs_ST = TxtBox_7
Else
Application(Urs&"_ST") = "Offline"
Urs_ST = TxtBox_8
End If
If Application(Urs&"_ST") = "Online" Then
Color_ST_s = "green"
Else
Color_ST_s = "red"
End If
Response.Write("<div style='float:left;width:100%;'><div style='float:left;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<script type='text/javascript'>document.write(decodeURIComponent('"&Urs&"'));</script>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b style='color:"&Color_ST_s&";'>"&Urs_ST&"</b></div><div style='float:left;'><form action='chats/chat.asp' style='float:left;' id='C_"&Form_Num&"' method='get' target='_parent'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='hidden' style='width:0;height:0;' name='Inviter' value='"&Session("UserNums")&"' /><input type='hidden' style='width:0;height:0;' name='Partner' value='"&Rs(1)&"' /><a href='#' id='C_L_"&Form_Num&"' style='' onclick=""document.getElementById('C_"&Form_Num&"').submit();"">"&TxtBox_5&"</a></form></div><div style='float:left;'><form action='frienddelete.asp' style='float:left;' id='D_"&Form_Num&"' method='post' target='nofresh'>&nbsp;&nbsp;&nbsp;&nbsp;<input type='hidden' style='width:0;height:0;' name='Orders' value='"&Rs(0)&"' /><a href='#' id='D_L_"&Form_Num&"' style='' onclick=""document.getElementById('D_"&Form_Num&"').submit();return false;"">"&TxtBox_6&"</a></form></div></div>")
Rs.MoveNext
Loop
Rs.Close
End If
Set Rs = Nothing
Set Dbc = Nothing
%>
