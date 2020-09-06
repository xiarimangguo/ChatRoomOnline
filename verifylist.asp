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
<title>VerifyList</title>
<%
If Application(Session("UserName")&"_LG") = "" Then
Application(Session("UserName")&"_LG") = "English"
End If
Select Case Application(Session("UserName")&"_LG")
Case "English"
TxtBox_1 = "Friends"
TxtBox_2 = "Verifies"
TxtBox_3 = "Add"
TxtBox_4 = "Agree"
TxtBox_5 = "Refuse"
TxtBox_6 = "Verify"
TxtBox_7 = "Agreed"
TxtBox_8 = "Refused"
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E7%AE%80%E4%BD%93%E4%B8%AD%E6%96%87"
TxtBox_1 = "好友"
TxtBox_2 = "验证"
TxtBox_3 = "添加"
TxtBox_4 = "同意"
TxtBox_5 = "拒绝"
TxtBox_6 = "验证"
TxtBox_7 = "已同意"
TxtBox_8 = "已拒绝"
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E7%B9%81%E9%AB%94%E4%B8%AD%E6%96%87"
TxtBox_1 = "好友"
TxtBox_2 = "驗證"
TxtBox_3 = "添加"
TxtBox_4 = "同意"
TxtBox_5 = "拒絕"
TxtBox_6 = "驗證"
TxtBox_7 = "已同意"
TxtBox_8 = "已拒絕"
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%D1%80%D1%83%D1%81%D1%81%D0%BA%D0%B8%D0%B9"
TxtBox_1 = "друзья"
TxtBox_2 = "Проверяет"
TxtBox_3 = "Добавить"
TxtBox_4 = "Согласен"
TxtBox_5 = "Отказываться"
TxtBox_6 = "Проверяет"
TxtBox_7 = "Согласовано"
TxtBox_8 = "Отказалась"
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E6%97%A5%E6%9C%AC%E8%AA%9E"
TxtBox_1 = "フレンズ"
TxtBox_2 = "検証する"
TxtBox_3 = "追加"
TxtBox_4 = "同意する"
TxtBox_5 = "ごみ"
TxtBox_6 = "検証する"
TxtBox_7 = "同意した"
TxtBox_8 = "拒否"
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "fran%C3%A7ais"
TxtBox_1 = "Copains"
TxtBox_2 = "Vérifie"
TxtBox_3 = "Ajouter"
TxtBox_4 = "Se mettre d'accord"
TxtBox_5 = "Refuser"
TxtBox_6 = "Vérifie"
TxtBox_7 = "D'accord"
TxtBox_8 = "Refusé"
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "Deutsche"
TxtBox_1 = "Freunde"
TxtBox_2 = "Überprüft"
TxtBox_3 = "Hinzufügen"
TxtBox_4 = "Zustimmen"
TxtBox_5 = "Sich weigern"
TxtBox_6 = "Überprüft"
TxtBox_7 = "Einverstanden"
TxtBox_8 = "Verweigert"
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case Else
TxtBox_1 = "Friends"
TxtBox_2 = "Verifies"
TxtBox_3 = "Add"
TxtBox_4 = "Agree"
TxtBox_5 = "Refuse"
TxtBox_6 = "Verify"
TxtBox_7 = "Agreed"
TxtBox_8 = "Refused"
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
End Select
%>
<%
Set Dbc = Server.CreateObject("Adodb.Connection")
Ctr = "provider=sqloledb;server=den1.mssql8.gear.host;database=cooles;uid=cooles;pwd=Rs5A9K~?JWw2;"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Show_Avn = "SELECT COUNT(Orders) FROM Verifies_"&Session("UserNums")
Rs.Open Show_Avn,Ctr,1,1
Application("Cnt_Vrs_"&Session("UserName")) = Rs(0).Value
Rs.Close
Get_Lst = "select * from Verifies_"&Session("UserNums")&" order by Orders"
Rs.Open Get_Lst,Ctr,1,1
Response.Write("<div><p style='font-size:1em;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b><a href='http://fscache20.cooles.top/login/friendlist.asp' style='color:blue;'>"&TxtBox_1&"</a></b>&nbsp;<b>|</b>&nbsp;<b>"&TxtBox_2&"</b>&nbsp;<b>|</b>&nbsp;<b><a href='http://fscache20.cooles.top/login/friendadd.html?lge="&Lcd&"' style='color:blue;'>"&TxtBox_3&"</a></b></p></div><iframe name='nofresh' style='display:none;width:0em;height:0em;'></iframe>")
If Not Rs.Eof Then
Form_Num = 0
Do Until Rs.Eof
Form_Num = Form_Num + 1
Select Case Rs(4)
Case "Verify"
Vrf_ST = TxtBox_6
Color_ST = "Orange"
Case "Refused"
Vrf_ST = TxtBox_8
Color_ST = "Red"
Case "Agreed"
Vrf_ST = TxtBox_7
Color_ST = "Green"
End Select
Response.Write("<div style='float:left;width:100%;'><div style='float:left;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<script type='text/javascript'>document.write(decodeURIComponent('"&Rs(2)&"'));</script>&nbsp;&nbsp;<b>|</b>&nbsp;&nbsp;<b style='color:"&Color_ST&";'>"&Vrf_ST&"</b></div><div style='float:left;'><form action='verifyagree.asp' style='float:left;' id='Y_"&Form_Num&"' method='post' target='nofresh'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='hidden' style='width:0;height:0;' name='Orders' value='"&Rs(0)&"' /><a href='#' id='Y_L_"&Form_Num&"' style='' onclick=""document.getElementById('Y_"&Form_Num&"').submit();return false;"">"&TxtBox_4&"</a></form></div><div style='float:left;'><form action='verifyrefuse.asp' style='float:left;' id='N_"&Form_Num&"' method='post' target='nofresh'>&nbsp;&nbsp;&nbsp;&nbsp;<input type='hidden' style='width:0;height:0;' name='Orders' value='"&Rs(0)&"' /><a href='#' id='N_L_"&Form_Num&"' style='' onclick=""document.getElementById('N_"&Form_Num&"').submit();return false;"">"&TxtBox_5&"</a></form></div></div>")
If Rs(4) = "Verify" Then
Else
Response.Write("<script type=""text/javascript"">function disable(){return false;}document.getElementById('Y_L_"&Form_Num&"').onclick=disable;document.getElementById('Y_L_"&Form_Num&"').style.color='gray';document.getElementById('N_L_"&Form_Num&"').onclick=disable;document.getElementById('N_L_"&Form_Num&"').style.color='gray';</script>")
End If
Rs.MoveNext
Loop
Rs.Close
Else
Rs.Close
Response.End
End If
Set Rs = Nothing
Set Dbc = Nothing
%>