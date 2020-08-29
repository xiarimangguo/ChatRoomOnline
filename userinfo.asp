<title>User Information</title>
<%
Session("Page") = "http://"&Request.ServerVariables("Server_Name")&Request.ServerVariables("Url")
If Session("Status") = "User logged in." Then
Response.Write "Welcome, "&Session("UserName")&"! "
Response.Write "<a href='http://fscache20.cooles.top/login/logout.asp'>Logout</a>"
Else
Response.Redirect("http://fscache20.cooles.top/login/login.html")
End If
%>