<%@ CODEPAGE=65001 %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<%
Inviter = Request.QueryString("Inviter")
Partner = Request.QueryString("Partner")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<style>
*{padding:0 auto;margin:0 auto;}
</style>
<title>SendMessage</title>
<script>
function reset()
{
	document.getElementById('Message').value=''
}
</script>
</head>
<body>
<iframe name="nofresh" style="display:none;width:0em;height:0em;">
</iframe>
<form method="POST" action="sendmessages.asp" target="nofresh">
<div style="width:100%;height:100%;">
<textarea name="Message" id="Message" style="background-color:#F5F5F5;width:100%;height:100%;font-size:1em;"></textarea>
<input type="hidden" style="width:0;height:0;" name="Inviter" value="<%=Inviter%>" />
<input type="hidden" style="width:0;height:0;" name="Partner" value="<%=Partner%>" />
</div>
<div style="position:absolute;bottom:0em;right:0em;width:6%;height:2em;">
<input type="submit" value="Send" style="background-color:#A9A9A9;color:white;width:100%;height:100%;font-size:1em;" onclick="setTimeout('reset()',50)"/>
</div>
</form>
</body>
</html>