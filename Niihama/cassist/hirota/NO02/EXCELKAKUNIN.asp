<html>
<head>
	<title>�Ј��Ǘ�</title>
	<base target="Right">
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->

<% w_sFLG = Request.QueryString("FLG")
   Select Case w_sFLG
   
 '----------------------------------------��������------------------------------------
	Case "1" %>
		<body>
			<h3 align=center>�� EXCEL�o�� ��</h3>
				<hr>
			<h3 align=center>�w�肵�������ŎЈ��f�[�^���o�͂��Ă���낵���ł����H</h3>

		<table align=center width=20%>
			<tr>
				<td align=center>
				<form action="EndCSV.asp" target="Right" method="Post" id=form1 name=form1>
					<p align=center><input type="submit" value="O K" id=submit1 name=submit1>
				</td>
				</form>
				<form action="CSV.asp" method="Post" target="Right" id=form2 name=form2>
				<td align=center>
					<input type="submit" value="�L�����Z��" id=submit2 name=submit2>
				</td>
				</form>
			<tr>
		</table>

<%
'----------------------------------------�����Ȃ�------------------------------------
	Case "2" %>
		<body>
			<h3 align=center>�� EXCEL�o�� ��</h3>
				<hr>
			<h3 align=center>�����͈͂̎w�肪����܂���B<br>���ׂĂ̎Ј��f�[�^���o�͂��Ă���낵���ł����H</h3>
		<br>
		<table align=center width=20%>
			<tr>
				<form action="EndCSV.asp" target="Right" method="Post" id=form3 name=form3>
					<td align=center>
						<p align=center><input type="submit" value="O K" id=submit1 name=submit1>
					</td>
				</form>
				<form action="CSV.asp" target="Right" id=form6 name=form6>
					<td align=center>
						<input type="submit" value="�� ��" id=submit4 name=submit4>
					</td>
				</form>
			<tr>
		</table>
		</body>
		</html>
<%
'----------------------------------------�Y���҂Ȃ�------------------------------------
	Case "3" %>
		<body>
			<h3 align=center>�� �o�̓G���[ ��</h3>
			<hr>
		<table align=center>
		<tr>
		<td>
			<h4 align=center><font color=red>�� �o�̓G���[���b�Z�[�W</font></h4>
		</td>
		</tr>
		</table>
		<table align=center>
		<tr>
			<td align=center>
				<ul>
				<li>�����ɊY������Ј��͂��܂���ł����B
				</ul>
			</td>
		</tr>
		</table>


		<table align="center" width=20%>
			<tr><form action="CSV.asp" target="Right" method="Post" id=form4 name=form4>
				<td align=center>
					<input type="submit" value="�� ��" id=submit3 name=submit3>
				</td>
				</form>
				<form action="INitiran.asp" target="Right" id=form5 name=form5>
				<td align=center>
						<input type="submit" value="�ꗗ" id=submit1 name=submit1>
				</td>
				</form>
			</tr>
		</table>

<% End Select %>
</body>
</html>
