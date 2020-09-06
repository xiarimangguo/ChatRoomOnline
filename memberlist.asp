<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<meta http-equiv="refresh" content="1">
<title>MemberList</title>
<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
If Application(Session("UserName")&"_LG") = "" Then
Application(Session("UserName")&"_LG") = "English"
End If
Select Case Application(Session("UserName")&"_LG")
Case "English"
TxtBox_1 = "All"
TxtBox_2 = "Friends"
TxtBox_3 = "Add"
TxtBox_4 = "Verify Sended."
TxtBox_5 = "Online"
TxtBox_6 = "Offline"
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E7%AE%80%E4%BD%93%E4%B8%AD%E6%96%87"
TxtBox_1 = "所有"
TxtBox_2 = "好友"
TxtBox_3 = "添加"
TxtBox_4 = "验证已发送"
TxtBox_5 = "在线"
TxtBox_6 = "离线"
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E7%B9%81%E9%AB%94%E4%B8%AD%E6%96%87"
TxtBox_1 = "所有"
TxtBox_2 = "好友"
TxtBox_3 = "添加"
TxtBox_4 = "驗證已發送"
TxtBox_5 = "在線"
TxtBox_6 = "離線"
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%D1%80%D1%83%D1%81%D1%81%D0%BA%D0%B8%D0%B9"
TxtBox_1 = "Все"
TxtBox_2 = "друзья"
TxtBox_3 = "Добавить"
TxtBox_4 = "Подтвердить отправлено."
TxtBox_5 = "В сети"
TxtBox_6 = "Не в сети"
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E6%97%A5%E6%9C%AC%E8%AA%9E"
TxtBox_1 = "すべて"
TxtBox_2 = "フレンズ"
TxtBox_3 = "追加"
TxtBox_4 = "送信済みを確認します"
TxtBox_5 = "オンライン"
TxtBox_6 = "オフライン"
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "fran%C3%A7ais"
TxtBox_1 = "Tout"
TxtBox_2 = "Copains"
TxtBox_3 = "Ajouter"
TxtBox_4 = "Vérifiez envoyé."
TxtBox_5 = "En ligne"
TxtBox_6 = "Hors ligne"
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "Deutsche"
TxtBox_1 = "Alles"
TxtBox_2 = "Freunde"
TxtBox_3 = "Hinzufügen"
TxtBox_4 = "Überprüfen Sie Gesendet."
TxtBox_5 = "Online"
TxtBox_6 = "Offline"
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case Else
TxtBox_1 = "All"
TxtBox_2 = "Friends"
TxtBox_3 = "Add"
TxtBox_4 = "Verify Sended."
TxtBox_5 = "Online"
TxtBox_6 = "Offline"
TxtBox_7 = ""
TxtBox_8 = ""
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
Set Rs2 = Server.CreateObject("Adodb.RecordSet")
Show_Ols = "SELECT UserName FROM Users ORDER BY UserNums"
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
Application("Onlines") = Onlines
If Application("Onlines") > 0 Then
Color_ST = "green"
Else
Color_ST = "red"
End If
Rs.Close
Show_Aun = "SELECT COUNT(UserName) FROM Users"
Rs.Open Show_Aun,Ctr,1,1
Application("All") = Rs(0).Value
Response.Write("<p style='font-size:1em;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>"&TxtBox_1&"</b>&nbsp;&nbsp;(<b style='color:"&Color_ST&";'>"&Application("Onlines")&"</b>/<b>"&Application("All")&"</b>)&nbsp;&nbsp;<b>|</b>&nbsp;&nbsp;<b><a href='http://fscache20.cooles.top/login/friendlist.asp' style='color:blue;'>"&TxtBox_2&"</a></b></p><iframe name='nofresh' style='display:none;width:0em;height:0em;'></iframe>")
Rs.Close
Show_Usr = "SELECT * FROM Users ORDER BY UserNums"
Rs.Open Show_Usr,Ctr,1,1
If Rs.Eof Then
Rs.Close
Else
Form_Num = 0
Do Until Rs.Eof
Form_Num = Form_Num + 1
Urs = Rs(1)
If DateDiff("s",Application(Urs&"_LT"),Now) <= 5 Then
Application(Urs&"_ST") = "Online"
Urs_ST = TxtBox_5
Else
Application(Urs&"_ST") = "Offline"
Urs_ST = TxtBox_6
End If
If Application(Urs&"_ST") = "Online" Then
Color_ST_s = "green"
Else
Color_ST_s = "red"
End If
Response.Write("<div style='float:left;width:100%;'><div style='float:left;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<script type='text/javascript'>document.write(decodeURIComponent('"&Urs&"'));</script>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b style='color:"&Color_ST_s&";'>"&Urs_ST&"</b></div><div style='float:left;'><form action='sendverify.asp' style='float:left;' id='A_"&Form_Num&"' method='post' target='nofresh'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='hidden' style='width:0;height:0;' name='Num' value='"&Rs(0)&"' /><input type='hidden' style='width:0;height:0;' name='Unm' value='"&Rs(1)&"' /><a href='#' id='A_L_"&Form_Num&"' style='' onclick=""document.getElementById('A_"&Form_Num&"').submit();document.getElementById('A_L_"&Form_Num&"').onclick='function(){return false;}';document.getElementById('A_L_"&Form_Num&"').innerHTML='"&TxtBox_4&"';document.getElementById('A_L_"&Form_Num&"').style.color='gray';"">"&TxtBox_3&"</a></form></div></div>")
Clk_Frd = "select * from Friends_"&Session("UserNums")&" where UserNums = '"&Rs(0)&"'"
Rs2.Open Clk_Frd,Ctr,1,1
If Not Rs2.Eof Then
Response.Write("<script type=""text/javascript"">function disable(){return false;}document.getElementById('A_L_"&Form_Num&"').onclick=disable;document.getElementById('A_L_"&Form_Num&"').style.color='gray';</script>")
Else
End If
Rs2.Close
Rs.MoveNext
Loop
Rs.Close
End If
Set Rs2 = Nothing
Set Rs = Nothing
Set Dbc = Nothing
%>
