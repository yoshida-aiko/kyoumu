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
    <br><br><br><br><br>
			<span class="msg"><%=request("txtMsg")%></span>
	</center>
	</body>

	</html>
<%
End Sub 
%>