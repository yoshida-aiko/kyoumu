<html>
<head>
	<title>�Ј��Ǘ�</title>
	<base target="Right">
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>
	<h3 align=center>�� CSV�o�� ��</h3>
		<hr>
	<h2 align=center><font color=red>CSV�o�͂��������܂����I</font></h2>
<table align=center>
	<tr>
		<td>
			<table align=center>
				<tr>
					<td>
						�o�͏ꏊ
					</td>
					<td>
						�F
					</td>
					<td>
						<%= Session.Contents("Path") %>
					</td>
				</tr>
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
				<tr>
					<td>
						�o��CSV�t�@�C��
					</td>
					<td>
						�F
					</td>
					<td>
						<a href="\\WEBSVR_2\infogram\hirota\No02\Sample.csv">�_�E�����[�h</a>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<br>
<table align=center width=20%>
	<tr>
		<form action="CSV.asp" id=form1 name=form1>
			<td align=center><p align=center><input type="submit" value="�� ��"></td>
		</form>
		<form action="INitiran.asp" target="Right">
			<td align=center valign=bottom><input type="submit" value="�� ��"></td>
		</form>
	<tr>
</table>
</body>
</html>
