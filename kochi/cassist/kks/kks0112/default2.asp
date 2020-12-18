<%@ Language=VBScript %>
<!--#include file="../../Common/com_All.asp"-->
<html>
<head>
<link rel=stylesheet href=../../common/style.css type=text/css>
</head>

<body>
<center>
<br><br><br>

	<%If Request("txtMsg")<>"" Then%>
		<span class="msg"><%=Request("txtMsg")%></span>
	<%Else%>
		<span class="msg"><%=C_BRANK_VIEW_MSG%></span>
	<%End If%>

</center>
</body>

</html>