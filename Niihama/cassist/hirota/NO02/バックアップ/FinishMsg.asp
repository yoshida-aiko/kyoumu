<html>
<head>
	<title>�Ј��Ǘ�</title>
	<base target="Right">
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>
<%
	w_sFLG = Request.QueryString("FLG")
	Select Case w_sFLG
		Case "1"
			Response.Write "<h3 align=center>�� �V �K �o �^ �� �� ��</h3>"
			Response.Write "<hr>"
			Response.Write "<h3 align=center><font color=red>�Ј��f�[�^��o�^���܂����I</font></h3>"

		Case "2"
			Response.Write "<h3 align=center>�� �C �� �� �� ��</h3>"
			Response.Write "<hr>"
			Response.Write "<h3 align=center><font color=red>�Ј��f�[�^���C�����܂����I</font></h3>"

		Case "3"
			Response.Write "<h3 align=center>�� ��@�� �� �� ��</h3>"
			Response.Write "<hr>"
			Response.Write "<h3 align=center><font color=red>�Ј��f�[�^���폜���܂����I</font></h3>"
	end Select
 %>
<form action="INitiran.asp" target="Right" id=form1 name=form1>
	<p align=center><input type="submit" value="�� ��" id=submit1 name=submit1></p>
</form>

</body>
</html>
