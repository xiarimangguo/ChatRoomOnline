<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<title>Language</title>
<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
If Application(Session("UserName")&"_LG") = "" Then
Application(Session("UserName")&"_LG") = "English"
End If
Select Case Application(Session("UserName")&"_LG")
Case "English"
TxtBox_1 = "Language"
TxtBox_2 = "Using"
TxtBox_3 = "Set"
TxtBox_4 = "Using"
TxtBox_5 = ""
TxtBox_6 = ""
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E7%AE%80%E4%BD%93%E4%B8%AD%E6%96%87"
TxtBox_1 = "语言"
TxtBox_2 = "当前"
TxtBox_3 = "设置"
TxtBox_4 = "已设置"
TxtBox_5 = ""
TxtBox_6 = ""
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E7%B9%81%E9%AB%94%E4%B8%AD%E6%96%87"
TxtBox_1 = "語言"
TxtBox_2 = "已設定"
TxtBox_3 = "設定"
TxtBox_4 = "已設定"
TxtBox_5 = ""
TxtBox_6 = ""
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%D1%80%D1%83%D1%81%D1%81%D0%BA%D0%B8%D0%B9"
TxtBox_1 = "язык"
TxtBox_2 = "Seted"
TxtBox_3 = "Устанавливать"
TxtBox_4 = "Seted"
TxtBox_5 = ""
TxtBox_6 = ""
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E6%97%A5%E6%9C%AC%E8%AA%9E"
TxtBox_1 = "言語"
TxtBox_2 = "設定済み"
TxtBox_3 = "セットする"
TxtBox_4 = "設定済み"
TxtBox_5 = ""
TxtBox_6 = ""
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "fran%C3%A7ais"
TxtBox_1 = "Langue"
TxtBox_2 = "Défini"
TxtBox_3 = "Ensemble"
TxtBox_4 = "Défini"
TxtBox_5 = ""
TxtBox_6 = ""
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "Deutsche"
TxtBox_1 = "Sprache"
TxtBox_2 = "Eingestellt"
TxtBox_3 = "Einstellen"
TxtBox_4 = "Eingestellt"
TxtBox_5 = ""
TxtBox_6 = ""
TxtBox_7 = ""
TxtBox_8 = ""
TxtBox_9 = ""
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case Else
TxtBox_1 = "Language"
TxtBox_2 = "Using"
TxtBox_3 = "Set"
TxtBox_4 = "Using"
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
Dim Language(7)
Language(0)="English"
Language(1)="简体中文"
Language(2)="繁體中文"
Language(3)="русский"
Language(4)="日本語"
Language(5)="français"
Language(6)="Deutsche"
Response.Write("<p style='font-size:1em;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>"&TxtBox_1&"</b>&nbsp;&nbsp;<b>("&TxtBox_2&":&nbsp;&nbsp;<b style='color:green;'><script type='text/javascript'>document.write(decodeURIComponent('"&Application(Session("UserName")&"_LG")&"'));</script></b>)</b></p><iframe name='nofresh' style='display:none;width:0em;height:0em;'></iframe>")
Num = -1
Form_Num = 0
Do Until Num > UBound(Language) - 2
Num = Num + 1
Form_Num = Form_Num + 1
If Application(Session("UserName")&"_LG") <> Server.URLEncode(Language(Num)) Then
Response.Write("<div style='float:left;width:100%;'><div style='float:left;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&Language(Num)&"</div><div style='float:left;'><form action='setlanguage.asp' style='float:left;' id='S_"&Form_Num&"' method='post' target='nofresh'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='hidden' style='width:0;height:0;' name='Lge' value='"&Language(Num)&"' /><a href='main.asp' target='_parent' id='S_L_"&Form_Num&"' style='' onclick=""document.getElementById('S_"&Form_Num&"').submit();"">"&TxtBox_3&"</a></form></div></div>")
Else
Response.Write("<div style='float:left;width:100%;'><div style='float:left;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&Language(Num)&"</div><div style='float:left;'><form action='setlanguage.asp' style='float:left;' id='S_"&Form_Num&"' method='post' target='nofresh'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='hidden' style='width:0;height:0;' name='Lge' value='"&Language(Num)&"' /><a href='main.asp' target='_parent' id='S_L_"&Form_Num&"' style='color:gray;' onclick=""return false;"">"&TxtBox_4&"</a></form></div></div>")
End If
Loop
%>