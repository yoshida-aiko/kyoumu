<html>
<head>
	<title>�Ј��Ǘ�</title>
	<base target="Right">
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>
	<h3 align=center>�� EXCEL�o�� ��</h3>
<hr>
	<h2 align=center><font color=red>EXCEL�o�͂��������܂����I</font></h2>
<table align=center>
	<tr>
		<td>
			�o�͌���
		</td>
		<td>
			�F
		</td>
		<td>
			<%= Request.QueryString("Count") %> ��
		</td>
	</tr>
</table>
<br>
<table align=center width=20%>
	<tr>
		<form action="EXCEL.asp" id=form1 name=form1>
			<td align=center><p align=center><input type="submit" value="�� ��" id=submit1 name=submit1>	</td>
		</form>
		<form action="INitiran.asp" target="Right" id=form2 name=form2>
			<td align=center valign=bottom><input type="submit" value="�� ��" id=submit2 name=submit2></td>
		</form>
	<tr>
</table>
</body>
</html>
