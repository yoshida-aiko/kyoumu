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
    <%call gs_title("授業出欠入力","一　覧")%>
	<br><br><br><br><br>
			<span class="msg"><%=Request("txtMsg") %></span>
	</center>
	</body>

	</html>
<%
End Sub 
%>