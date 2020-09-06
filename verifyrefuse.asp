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
Upd_ST = "update Verifies_"&Session("UserNums")&" set States = 'Refused' where Orders = "&Orders
Rs.Open Upd_ST,Ctr,1,3
Response.Write("Verify Refused.")
Set Rs = Nothing
Set Dbc = Nothing
%>
