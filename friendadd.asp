<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<title>FriendAdd</title>
<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
If Application(Session("UserName")&"_LG") = "" Then
Application(Session("UserName")&"_LG") = "English"
End If
Select Case Application(Session("UserName")&"_LG")
Case "English"
TxtBox_1 = "Friends"
TxtBox_2 = "Verifies"
TxtBox_3 = "Add"
TxtBox_4 = " results for "
TxtBox_5 = "Add"
TxtBox_6 = "Verify Sended."
TxtBox_7 = " results for "
TxtBox_8 = "Add"
TxtBox_9 = "Verify Sended."
TxtBox_10 = "There is not any results for "
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E7%AE%80%E4%BD%93%E4%B8%AD%E6%96%87"
TxtBox_1 = "好友"
TxtBox_2 = "验证"
TxtBox_3 = "添加"
TxtBox_4 = " 条结果符合 "
TxtBox_5 = "添加"
TxtBox_6 = "验证已发送"
TxtBox_7 = " 条结果符合 "
TxtBox_8 = "添加"
TxtBox_9 = "验证已发送"
TxtBox_10 = "没有结果符合 "
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E7%B9%81%E9%AB%94%E4%B8%AD%E6%96%87"
TxtBox_1 = "好友"
TxtBox_2 = "驗證"
TxtBox_3 = "添加"
TxtBox_4 = " 個結果符合 "
TxtBox_5 = "添加"
TxtBox_6 = "驗證已發送"
TxtBox_7 = " 個結果符合 "
TxtBox_8 = "添加"
TxtBox_9 = "驗證已發送"
TxtBox_10 = "沒有結果符合 "
TxtBox_11 = ""
TxtBox_12 = ""
Case "%D1%80%D1%83%D1%81%D1%81%D0%BA%D0%B8%D0%B9"
TxtBox_1 = "друзья"
TxtBox_2 = "Проверяет"
TxtBox_3 = "Добавить"
TxtBox_4 = " результатов для "
TxtBox_5 = "Добавить"
TxtBox_6 = "Подтвердить отправлено."
TxtBox_7 = " результатов для "
TxtBox_8 = "Добавить"
TxtBox_9 = "Подтвердить отправлено."
TxtBox_10 = "Нет результатов по запросу "
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E6%97%A5%E6%9C%AC%E8%AA%9E"
TxtBox_1 = "フレンズ"
TxtBox_2 = "検証する"
TxtBox_3 = "追加"
TxtBox_4 = " 検索結果 for "
TxtBox_5 = "追加"
TxtBox_6 = "送信済みを確認します"
TxtBox_7 = " 検索結果 for "
TxtBox_8 = "追加"
TxtBox_9 = "送信済みを確認します"
TxtBox_10 = "0 検索結果 for "
TxtBox_11 = ""
TxtBox_12 = ""
Case "fran%C3%A7ais"
TxtBox_1 = "Copains"
TxtBox_2 = "Vérifie"
TxtBox_3 = "Ajouter"
TxtBox_4 = " résultats pour "
TxtBox_5 = "Ajouter"
TxtBox_6 = "Vérifiez envoyé."
TxtBox_7 = " résultats pour "
TxtBox_8 = "Ajouter"
TxtBox_9 = "Vérifiez envoyé."
TxtBox_10 = "Il n'y a aucun résultat pour "
TxtBox_11 = ""
TxtBox_12 = ""
Case "Deutsche"
TxtBox_1 = "Freunde"
TxtBox_2 = "Überprüft"
TxtBox_3 = "Hinzufügen"
TxtBox_4 = " Ergebnisse für "
TxtBox_5 = "Hinzufügen"
TxtBox_6 = "Überprüfen Sie Gesendet."
TxtBox_7 = " Ergebnisse für "
TxtBox_8 = "Hinzufügen"
TxtBox_9 = "Überprüfen Sie Gesendet."
TxtBox_10 = "Es gibt keine Ergebnisse für "
TxtBox_11 = ""
TxtBox_12 = ""
Case Else
TxtBox_1 = "Friends"
TxtBox_2 = "Verifies"
TxtBox_3 = "Add"
TxtBox_4 = " results for "
TxtBox_5 = "Add"
TxtBox_6 = "Verify Sended."
TxtBox_7 = " results for "
TxtBox_8 = "Add"
TxtBox_9 = "Verify Sended."
TxtBox_10 = "There is not any results for "
TxtBox_11 = ""
TxtBox_12 = ""
End Select
%>
<%
Unm = Server.URLEncode(Request.QueryString("Unm"))
Set Dbc = Server.CreateObject("Adodb.Connection")
Ctr = "provider=sqloledb;server=den1.mssql8.gear.host;database=cooles;uid=cooles;pwd=Rs5A9K~?JWw2;"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Set Rs2 = Server.CreateObject("Adodb.RecordSet")
Set Rs3 = Server.CreateObject("Adodb.RecordSet")
Show_Usr = "select * from Users where UserName like '"&Unm&"'"
Rs.Open Show_Usr,Ctr,1,1
Response.Write("<div><p style='font-size:1em;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b><a href='http://fscache20.cooles.top/login/friendlist.asp' style='color:blue;'>"&TxtBox_1&"</a></b>&nbsp;<b>|</b>&nbsp;<b><a href='http://fscache20.cooles.top/login/verifylist.asp' style='color:blue;'>"&TxtBox_2&"</a></b>&nbsp;<b>|</b>&nbsp;<b>"&TxtBox_3&"</b></p></div><iframe name='nofresh' style='display:none;width:0em;height:0em;'></iframe>")
If Not Rs.Eof Then
Form_Num = 0
Response.Write("<p style='font-size:1em;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&Rs.RecordCount&TxtBox_4&"'<script type='text/javascript'>document.write(decodeURIComponent('"&Unm&"'));</script>'</p>")
Do Until Rs.Eof
Form_Num = Form_Num + 1
Response.Write("<div style='float:left;width:100%;'><div style='float:left;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<script type='text/javascript'>document.write(decodeURIComponent('"&Rs(1)&"'));</script></div><div style='float:left;'><form action='sendverify.asp' style='float:left;' id='F_"&Form_Num&"' method='post' target='nofresh'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='hidden' style='width:0;height:0;' name='Num' value='"&Rs(0)&"' /><input type='hidden' style='width:0;height:0;' name='Unm' value='"&Rs(1)&"' /><a href='#' id='F_L_"&Form_Num&"' style='' onclick=""document.getElementById('F_"&Form_Num&"').submit();document.getElementById('F_L_"&Form_Num&"').onclick='function(){return false;}';document.getElementById('F_L_"&Form_Num&"').innerHTML='"&TxtBox_6&"';document.getElementById('F_L_"&Form_Num&"').style.color='gray';"">"&TxtBox_5&"</a></form></div></div>")
Clk_Frd = "select * from Friends_"&Session("UserNums")&" where UserNums = '"&Rs(0)&"'"
Rs2.Open Clk_Frd,Ctr,1,1
If Not Rs2.Eof Then
Response.Write("<script type=""text/javascript"">function disable(){return false;}document.getElementById('F_L_"&Form_Num&"').onclick=disable;document.getElementById('F_L_"&Form_Num&"').style.color='gray';</script>")
Else
End If
Rs2.Close
Rs.MoveNext
Loop
Rs.Close
Else
If IsNumeric(Unm) Then
Rs.Close
Show_Usr_s = "select * from Users where UserNums = '"&Unm&"'"
Rs3.Open Show_Usr_s,Ctr,1,1
If Not Rs3.Eof Then
Form_Num = 0
Response.Write("<p style='font-size:1em;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&Rs3.RecordCount&TxtBox_7&"'<script type='text/javascript'>document.write(decodeURIComponent('"&Unm&"'));</script>'</p>")
Do Until Rs3.Eof
Form_Num = Form_Num + 1
Response.Write("<div style='float:left;width:100%;'><div style='float:left;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<script type='text/javascript'>document.write(decodeURIComponent('"&Rs3(1)&"'));</script></div><div style='float:left;'><form action='sendverify.asp' style='float:left;' id='F_"&Form_Num&"' method='post' target='nofresh'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='hidden' style='width:0;height:0;' name='Num' value='"&Rs3(0)&"' /><input type='hidden' style='width:0;height:0;' name='Unm' value='"&Rs3(1)&"' /><a href='#' id='F_L_"&Form_Num&"' style='' onclick=""document.getElementById('F_"&Form_Num&"').submit();document.getElementById('F_L_"&Form_Num&"').onclick='function(){return false;}';document.getElementById('F_L_"&Form_Num&"').innerHTML='"&TxtBox_9&"';document.getElementById('F_L_"&Form_Num&"').style.color='gray';"">"&TxtBox_8&"</a></form></div></div>")
Clk_Frd = "select * from Friends_"&Session("UserNums")&" where UserNums = '"&Rs3(0)&"'"
Rs2.Open Clk_Frd,Ctr,1,1
If Not Rs2.Eof Then
Response.Write("<script type=""text/javascript"">function disable(){return false;}document.getElementById('F_L_"&Form_Num&"').onclick=disable;document.getElementById('F_L_"&Form_Num&"').style.color='gray';</script>")
Else
End If
Rs2.Close
Rs3.MoveNext
Loop
Rs3.Close
Else
Response.Write("<p style='font-size:1em;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&TxtBox_10&"'<script type='text/javascript'>document.write(decodeURIComponent('"&Unm&"'));</script>'.</p>")
End If
Else
Response.Write("<p style='font-size:1em;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&TxtBox_10&"'<script type='text/javascript'>document.write(decodeURIComponent('"&Unm&"'));</script>'.</p>")
End If
End If
Set Rs3 = Nothing
Set Rs2 = Nothing
Set Rs = Nothing
Set Dbc = Nothing
%>