<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<title>Working...</title>
<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
Orders = Request.Form("Orders")
Set Dbc = Server.CreateObject("Adodb.Connection")
Ctr = "provider=sqloledb;server=den1.mssql8.gear.host;database=;uid=;pwd=;"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Get_Inf = "select UserNums from Friends_"&Session("UserNums")&" where Orders = "&Orders
Rs.Open Get_Inf,Ctr,1,1
Fr_Num = Rs(0)
Rs.Close
Del_Fr = "delete Friends_"&Session("UserNums")&" where Orders = "&Orders&";delete Friends_"&Fr_Num&" where UserNums = "&Session("UserNums")
Rs.Open Del_Fr,Ctr,1,3
Set Fs = Server.CreateObject("Scripting.FileSystemObject")
If Fs.FileExists(Server.MapPath("/login/chats/Chat_"&Session("UserNums")&"_"&Fr_Num&"_Bord.asp")) Then
Del = "/login/chats/Chat_"&Session("UserNums")&"_"&Fr_Num&"_Bord.asp"
Else
If Fs.FileExists(Server.MapPath("/login/chats/Chat_"&Fr_Num&"_"&Session("UserNums")&"_Bord.asp")) Then
Del = "/login/chats/Chat_"&Fr_Num&"_"&Session("UserNums")&"_Bord.asp"
Else
Del = ""
End If
End If
If Del <> "" Then
Fs.DeleteFile(Server.Mappath(Del))
End If
Response.Write("Successfully Deleted.")
Set Fs = Nothing
Set Rs = Nothing
Set Dbc = Nothing
%>
