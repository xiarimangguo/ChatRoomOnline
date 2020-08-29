<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
Dim Message
Inviter = Request.Form("Inviter")
Partner = Request.Form("Partner")
Message = Server.URLEncode(Request.Form("Message"))
If Instr(Message,"+") <> 0 Then
Message = Replace(Message,"+","%20")
Else
End If
Dim Color(8)
Color(0)="#8B0000"
Color(1)="#FF8C00"
Color(2)="#808000"
Color(3)="#008B8B"
Color(4)="#00BFFF"
Color(5)="#800080"
Color(6)="#FF1493"
Color(7)="#00FF7F"
Randomize
Random_Num = Int(Rnd*(7+1))
MsgText = "<p style='font-size:0.8em;color:#008000;'><script type='text/javascript'>document.write(decodeURIComponent('"&Session("UserName")&"'));</script>&nbsp;&nbsp;"&now()&"</p><p style='font-size:1em;color:"&Color(Random_Num)&";'><script type='text/javascript'>document.write(decodeURIComponent('"&Message&"'));</script></p>"
Set Fs = Server.CreateObject("Scripting.FileSystemObject")
If Fs.FileExists(Server.MapPath("/login/chats/Chat_"&Inviter&"_"&Partner&"_Bord.asp")) Then
SavePth = "/login/chats/Chat_"&Inviter&"_"&Partner&"_Bord.asp"
Else
If Fs.FileExists(Server.MapPath("/login/chats/Chat_"&Partner&"_"&Inviter&"_Bord.asp")) Then
SavePth = "/login/chats/Chat_"&Partner&"_"&Inviter&"_Bord.asp"
Else
Response.Write("Request Failed. <a href='javascript:void(0);'>Retry</a>")
Response.End
End If
End If
Set File = Fs.OpenTextFile(Server.MapPath(SavePth),8,True)
File.WriteLine(MsgText)
File.Close
Set File = Nothing
Set Fs = Nothing
%>