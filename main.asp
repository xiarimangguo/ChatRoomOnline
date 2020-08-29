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
<title>Chat Room Online!</title>
<%
If Application(Session("UserName")&"_LG") = "" Then
Application(Session("UserName")&"_LG") = "English"
End If
Select Case Application(Session("UserName")&"_LG")
Case "English"
TxtBox_1 = "Welcome to the online chat room!"
TxtBox_2 = "MemberList"
TxtBox_3 = "Language"
TxtBox_4 = "Hi,"
TxtBox_5 = "Logout"
TxtBox_6 = "MemberList"
TxtBox_7 = "Close"
TxtBox_8 = "Language"
TxtBox_9 = "Close"
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E7%AE%80%E4%BD%93%E4%B8%AD%E6%96%87"
TxtBox_1 = "欢迎来到在线聊天室!"
TxtBox_2 = "成员"
TxtBox_3 = "语言"
TxtBox_4 = "你好,"
TxtBox_5 = "退出"
TxtBox_6 = "成员"
TxtBox_7 = "关闭"
TxtBox_8 = "语言"
TxtBox_9 = "关闭"
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E7%B9%81%E9%AB%94%E4%B8%AD%E6%96%87"
TxtBox_1 = "歡迎來到在線聊天室!"
TxtBox_2 = "成員"
TxtBox_3 = "語言"
TxtBox_4 = "你好,"
TxtBox_5 = "登出"
TxtBox_6 = "成員"
TxtBox_7 = "關閉"
TxtBox_8 = "語言"
TxtBox_9 = "關閉"
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%D1%80%D1%83%D1%81%D1%81%D0%BA%D0%B8%D0%B9"
TxtBox_1 = "Добро пожаловать в онлайн-чат!"
TxtBox_2 = "члены"
TxtBox_3 = "язык"
TxtBox_4 = "Привет,"
TxtBox_5 = "выйти"
TxtBox_6 = "члены"
TxtBox_7 = "близко"
TxtBox_8 = "язык"
TxtBox_9 = "близко"
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "%E6%97%A5%E6%9C%AC%E8%AA%9E"
TxtBox_1 = "オンラインチャットルームへようこそ!"
TxtBox_2 = "会員"
TxtBox_3 = "言語"
TxtBox_4 = "こんにちは、"
TxtBox_5 = "ログアウト"
TxtBox_6 = "会員"
TxtBox_7 = "閉じる"
TxtBox_8 = "言語"
TxtBox_9 = "閉じる"
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "fran%C3%A7ais"
TxtBox_1 = "Bienvenue dans la salle de chat en ligne!"
TxtBox_2 = "Membres"
TxtBox_3 = "Langue"
TxtBox_4 = "Salut,"
TxtBox_5 = "Se déconnecter"
TxtBox_6 = "Membres"
TxtBox_7 = "Fermer"
TxtBox_8 = "Langue"
TxtBox_9 = "Fermer"
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case "Deutsche"
TxtBox_1 = "Willkommen im Online-Chatraum!"
TxtBox_2 = "Mitglieder"
TxtBox_3 = "Sprache"
TxtBox_4 = "Hello,"
TxtBox_5 = "Ausloggen"
TxtBox_6 = "Mitglieder"
TxtBox_7 = "Schließen"
TxtBox_8 = "Sprache"
TxtBox_9 = "Schließen"
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
Case Else
TxtBox_1 = "Welcome to the online chat room!"
TxtBox_2 = "MemberList"
TxtBox_3 = "Language"
TxtBox_4 = "Hi,"
TxtBox_5 = "Logout"
TxtBox_6 = "MemberList"
TxtBox_7 = "Close"
TxtBox_8 = "Language"
TxtBox_9 = "Close"
TxtBox_10 = ""
TxtBox_11 = ""
TxtBox_12 = ""
End Select
%>
<%
Session("Page") = "http://"&Request.ServerVariables("Server_Name")&Request.ServerVariables("Url")
If Session("Status") = "User logged in." Then
Response.Write "<p style='font-size:1em;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&TxtBox_1&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='javascript:void(0);' onclick='HiddenLanguage();ShowList()' style='color:blue;'>"&TxtBox_2&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='javascript:void(0);' onclick='HiddenList();ShowLanguage();' style='color:blue;'>"&TxtBox_3&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&TxtBox_4&" <script type='text/javascript'>document.write(decodeURIComponent('"&Session("UserName")&"'));</script>! <a href='http://fscache20.cooles.top/login/logout.asp' style='color:blue;'>"&TxtBox_5&"</a></p>"
Else
If Lcd <> "" Then
Response.Redirect("http://fscache20.cooles.top/login/login.html?lge="&Lcd)
Else
Response.Redirect("http://fscache20.cooles.top/login/login.html")
End If
End If
%>
<html>
<head>
<style>
*{padding:0 auto;margin:0 auto;}
</style>
<script>
function ShowList()
{document.getElementById('MemberList').style.display = '';}
function HiddenList()
{document.getElementById('MemberList').style.display = 'none';}
function ShowLanguage()
{document.getElementById('Language').style.display = '';}
function HiddenLanguage()
{document.getElementById('Language').style.display = 'none';}
</script>
</head>
<body>
<div style="float:left;display:inline;width:100%;height:87%;">
<iframe src="http://fscache20.cooles.top/login/chatbord.asp" style="width:100%;height:100%;">
</iframe>
<div style="float:left;display:inline;width:100%;height:11%;">
<iframe src="http://fscache20.cooles.top/login/sendmessage.html" style="width:100%;height:100%;"></iframe>
</div>
<div style="position:absolute;right:0em;top:1.25em;width:40%;height:87%;display:;" id="MemberList">
<div style="width:100%;height:1em;background-color:blue;"><p style="font-size:1em;color:white;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=TxtBox_6%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="javascript:void(0);" style="color:white;" onclick='HiddenList()'>[<%=TxtBox_7%>]</a></p></div>
<div style="width:100%;height:98.8%;">
<iframe src="http://fscache20.cooles.top/login/memberlist.asp" style="width:100%;height:100%;"></iframe>
</div>
</div>
</div>
<div style="position:absolute;right:0em;top:1.25em;width:40%;height:87%;display:none;" id="Language">
<div style="width:100%;height:1em;background-color:brown;"><p style="font-size:1em;color:white;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=TxtBox_8%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="javascript:void(0);" style="color:white;" onclick='HiddenLanguage()'>[<%=TxtBox_9%>]</a></p></div>
<div style="width:100%;height:98.8%;">
<iframe src="http://fscache20.cooles.top/login/language.asp" style="width:100%;height:100%;"></iframe>
</div>
</div>
</div>
</body>
</html>