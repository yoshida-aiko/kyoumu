<%@ Language=VBScript %>
<!--#include file="../../Common/com_All.asp"-->
<html>
<head>
<link rel=stylesheet href=../../common/style.css type=text/css>
</head>
<body>
<center>
<br><br><br>
	<span class="msg">
		<%
		If trim(Request("NoTanMsg")) <> "" Then

		Else
			response.write C_BRANK_VIEW_MSG
		End If
		%>
	</span>
</center>

</body>
</html>