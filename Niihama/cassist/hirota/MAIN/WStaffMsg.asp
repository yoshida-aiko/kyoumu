<html>
<head>
	<title>�Ј��Ǘ�</title>
	<base target="Right">
</head>
<% ErrorCD = Request.QueryString("WStaff") %>
<% Select Case ErrorCD %>
<% Case "1" %>
	<body>
		<h3 align=center>�� �o�^�G���[ ��</h3>
			<hr><br>
		<h4 align=center><font color=red>���f�[�^�x�[�X�ɓo�^���邱�Ƃ��o���܂���ł����B</font></h4>
	<table align=center>
	<tr>
		<td>
			<ul>
			<li>�Ј�CD <%= Session.Contents("ErrorCD") %> �͂��łɎg���Ă��܂��B
			</ul>
		</td>
	</tr>
	</table>
	<table align=center>
		<tr>
			<td>
				�Ј�CD���d�����Ă���̂œo�^���邱�Ƃ͏o���܂���B<br>
				�Ⴄ�Ј�CD����͂��A�ēx�o�^���ĉ������B
			</td>
		</tr>
	</table>
	<p align=center><input type="button" value="�߂�" onClick=history.go(-1) id=button2 name=button2>

<% Case "2" %>

		<body>
			<h3 align=center>�� �d���f�[�^ ��</h3>
				<hr><br>
			<h4 align=center><font color=red>���f�[�^�x�[�X�ɏd���f�[�^������܂��B</font></h4>
		<table align=center>
		<tr>
			<td>
				<ul>
				<li>�ȑO�폜���ꂽ���R�[�h���f�[�^�x�[�X�ɋL������Ă��܂��B
				<p align=center>�Ј�CD <%= Session.Contents("�Ј�CD") %> ���㏑�����Ă���낵���ł����H</p>
				</ul>
			</td>
		</tr>
		</table>

		<form action="SQLexe.asp" method="Post" id=form1 name=form1>
			<input type="hidden" name="CD" value="<%= Session.Contents("�Ј�CD") %>">
			<input type="hidden" name="NAME" value="<%= Session.Contents("�Ј�����") %>">
			<input type="hidden" name="BIRTHDAY" value="<%= Session.Contents("���N����") %>">
			<input type="hidden" name="TELL1" value="<%= Session.Contents("�d�b�ԍ�1") %>">
			<input type="hidden" name="TELL2" value="<%= Session.Contents("�d�b�ԍ�2") %>">
			<input type="hidden" name="POST" value="<%= Session.Contents("�X��") %>">
			<input type="hidden" name="ADDRESS1" value="<%= Session.Contents("�Z��1") %>">
			<input type="hidden" name="ADDRESS2" value="<%= Session.Contents("�Z��2") %>">
			<input type="hidden" name="BIKOU" value="<%= Session.Contents("���l") %>">
			<% Session.Contents("FLG") = "UPDATE" %>
			<table align="center" width=30%>
				<tr>
					<td align=center><input type="submit" value="O K" id=submit2 name=submit2></td>
		</FORM>
			<form action="ADDNEW.asp" target="Right" method="Post" id=form2 name=form2>
					<td align=center><input type="submit" value="�o�^�L�����Z��" id=submit1 name=submit1></td>
			</form>
				</tr>
		</table>
<% End Select %>
</body>
</html>