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
			<% if Session.Contents("SELECT") = "CSV" then %>
				<h3 align=center>�� CSV�o�� ��</h3>
			<% else %>
				<h3 align=center>�� EXCEL�o�� ��</h3>
			<% end if %>
				<hr>
			<h3 align=center>�w�肵�������ŎЈ��f�[�^���o�͂��Ă���낵���ł����H</h3>

		<table align=center width=20%>
			<tr>
				<td align=center>
				<% if Session.Contents("SELECT") = "CSV" then %>
					<form action="EndCSV.asp" target="Right" method="Post" id=form1 name=form1>
				<% else %>
					<form action="EndEXCEL.asp" target="Right" method="Post">
				<% end if %>
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
			<% if session.Contents("SELECT") = "CSV" then %>
				<h3 align=center>�� CSV�o�� ��</h3>
			<% else %>
				<h3 align=center>�� EXCEL�o�� ��</h3>
			<% end if %>
				<hr>
			<h3 align=center>�����͈͂̎w�肪����܂���B<br>���ׂĂ̎Ј��f�[�^���o�͂��Ă���낵���ł����H</h3>
		<br>
		<table align=center width=20%>
			<tr>
				<% if session.Contents("SELECT") = "CSV" then %>
					<form action="EndCSV.asp" target="Right" method="Post" id=form3 name=form3>
				<% else %>
					<form action="EndEXCEL.asp" target="Right" method="Post">
				<% end if %>
					<td align=center>
						<p align=center><input type="submit" value="O K" id=submit1 name=submit1>
					</td>
				</form>
				<form action="CSV.asp" target="Right">
					<td align=center>
						<input type="submit" value="�� ��">
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
			<tr>
				<% if Session.Contents("SELECT") = "CSV" then %>
					<form action="CSV.asp" target="Right" method="Post" id=form4 name=form4>
				<% else %>
					<form action="EXCEL.asp" target="Right" method="Post">
				<% end if %>
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
