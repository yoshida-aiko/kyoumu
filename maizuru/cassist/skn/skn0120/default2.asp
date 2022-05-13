<%@ Language=VBScript %>
<% 
 if Request("txtMode") = "no" then
	w_msg = "ŽŽŒ±€”õŠúŠÔ‚Å‚Í‚ ‚è‚Ü‚¹‚ñB"
 else
 	w_msg = C_BRANK_VIEW_MSG
 end if
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