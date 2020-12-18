<%@ Language=VBScript %>
<!--#include file="../../Common/com_All.asp"-->
<%call main

Sub main()
%>

	<html>
	<head>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	</head>

	<body>
	<center>
    <%call gs_title("Žö‹ÆoŒ‡“ü—Í","ˆê@——")%>
	<br><br><br><br><br>
			<span class="msg"><%=Request("txtMsg") %></span>
	</center>
	</body>

	</html>
<%
End Sub 
%>