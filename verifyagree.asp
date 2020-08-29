<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<title>Working...</title>
<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
Orders = Request.Form("Orders")
Set Dbc = Server.CreateObject("Adodb.Connection")
Ctr = "provider=sqloledb;server=den1.mssql8.gear.host;database=cooles;uid=cooles;pwd=Rs5A9K~?JWw2;"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Get_Inf = "select * from Verifies_"&Session("UserNums")&" where Orders = "&Orders
Rs.Open Get_Inf,Ctr,1,1
Vrf_Num = Rs(1)
Vrf_Unm = Rs(2)
Rs.Close
Upd_ST = "update Verifies_"&Session("UserNums")&" set States = 'Agreed' where Orders = "&Orders
Rs.Open Upd_ST,Ctr,1,3
Set Rs2 = Server.CreateObject("Adodb.RecordSet")
Add_Frd = "insert into Friends_"&Session("UserNums")&" values('"&Vrf_Num&"','"&Vrf_Unm&"','');insert into Friends_"&Vrf_Num&" values('"&Session("UserNums")&"','"&Session("UserName")&"','');"
Rs2.Open Add_Frd,Ctr,1,3
Set Stm = Server.CreateObject("ADODB.Stream")
With Stm
    .Type = 2
    .Mode = 3
    .Charset = "utf-8"
    .Open 
    .WriteText .LoadFromFile(Server.MapPath("/login/chats/chatbord.asp"))
    .Position = 3
End With
Set Stms = Server.CreateObject("ADODB.Stream")
With Stms
    .Type = 1
    .Mode = 3
    .Open
End With
Stm.CopyTo Stms
With Stms
    .SaveToFile Server.MapPath("/login/chats/Chat_"&Session("UserNums")&"_"&Vrf_Num&"_Bord.asp"),2
    .Close
End With
With Stm
    .Flush
    .Close
End With
Response.Write("Verify Agreed.")
Set Stms = Nothing
Set Stm = Nothing
Set Rs2 = Nothing
Set Rs = Nothing
Set Dbc = Nothing
%>