<%@ Language=VBScript %>
<!--#include file="../../Common/com_All.asp"-->
<%call main

Sub main()
	Dim txtMsg
	
	if request("data_flg") <> "" then
		txtMsg = "項目を選択後、入力または、表示ボタンを押してください"
	else
		txtMsg = request("txtMsg")
	end if
%>

	<html>
	<head>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	</head>

	<body>
	<center>
    <br><br><br><br><br>
			<span class="msg"><%=txtMsg%></span>
	</center>
	</body>

	</html>
<%
End Sub 
%>