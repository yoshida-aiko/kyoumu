<%@ Language=VBScript %>
<%
	dim w_msg
	
	If request("txtmsg") = "" then 
		w_msg = C_BRANK_VIEW_MSG
	else
		w_msg = request("txtmsg")
	end If

%>

<!--#include file="../../Common/com_All.asp"-->
<html>
<head>
<link rel=stylesheet href=../../common/style.css type=text/css>
</head>
<body>
<center>
<br><br><br>
<span class="msg"><%=w_msg%></span>
</center>
</body>
</html>