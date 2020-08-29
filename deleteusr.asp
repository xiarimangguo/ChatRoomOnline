<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<title>User Delete</title>
<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
Session("Page") = "http://"&Request.ServerVariables("Server_Name")&Request.ServerVariables("Url")
If Session("Status") = "User logged in." Then
Response.Write "Welcome, <script type='text/javascript'>document.write(decodeURIComponent('"&Session("UserName")&"'));</script>! "
Response.Write "<a href='http://fscache20.cooles.top/login/logout.asp'>Logout</a>"
Else
Response.Redirect("http://fscache20.cooles.top/login/login.html")
End If
%>
<p>User Delete</p>
<form method="POST" action="delete.asp">
CheckWd: <input type="text" name="Clw"><br />
PassWord: <input type="password" name="Pwd"><br />
<input type="submit" value="Submit">
<p>Type "<b style="color:red;">I am sure that I will delete myself from this server.</b>" in the "<b style="color:red;">CheckWd</b>" inputbox to <b style="color:red;">DELETE</b> yourself from this server. </p>
</form>