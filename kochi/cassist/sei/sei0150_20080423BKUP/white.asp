<%@ Language=VBScript %>
<%call main

Sub main()
%>

	<html>
	<head>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	</head>

	<body>
	<center>
	<br><br><br>
			<span class="msg"><%=Request("txtMsg") %></span>
	</center>
	</body>

	</html>
<%
End Sub 
%>