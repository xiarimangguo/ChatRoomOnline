<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<title>Working...</title>
<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
Num = Request.Form("Num")
Unm = Request.Form("Unm")
Set Dbc = Server.CreateObject("Adodb.Connection")
Ctr = "provider=sqloledb;server=den1.mssql8.gear.host;database=cooles;uid=cooles;pwd=Rs5A9K~?JWw2;"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Snd_Vrf = "insert into Verifies_"&Num&" values('"&Session("UserNums")&"','"&Session("UserName")&"','','Verify')"
Rs.Open Snd_Vrf,Ctr,1,3
Response.Write("Verify Sended.")
Rs.Close
Set Rs = Nothing
Set Dbc = Nothing
%>