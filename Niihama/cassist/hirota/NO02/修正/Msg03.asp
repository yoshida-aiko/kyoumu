<html>
<head>
	<title>�Ј��Ǘ�</title>
	<base target="Right">
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>
	<h3 align=center>�� �d���f�[�^ ��</h3>
		<hr><br>
	<h4 align=center><font color=red>���f�[�^�x�[�X�f�[�^�d�����b�Z�[�W</font></h4>
<table align=center>
<tr>
	<td>
		<ul>
		<li>�ȑO�폜���ꂽ���R�[�h���f�[�^�x�[�X�ɋL������Ă��܂��B
		<p align=center>�Ј�CD <%= Request.QueryString("CD") %> ���㏑�����Ă���낵���ł����H</p>
		</ul>
	</td>
</tr>
</table>

<form action="SQLexe.asp" method="Post">
	<input type="hidden" name="�Ј�CD" value="<%= Request.QueryString("CD") %>">
	<input type="hidden" name="�Ј�����" value="<%= Session.Contents("�Ј�����") %>">
	<input type="hidden" name="���N����" value="<%= Session.Contents("���N����") %>">
	<input type="hidden" name="�d�b�ԍ�1" value="<%= Session.Contents("�d�b�ԍ�1") %>">
	<input type="hidden" name="�d�b�ԍ�2" value="<%= Session.Contents("�d�b�ԍ�2") %>">
	<input type="hidden" name="�X��" value="<%= Session.Contents("�X��") %>">
	<input type="hidden" name="�Z��1" value="<%= Session.Contents("�Z��1") %>">
	<input type="hidden" name="�Z��2" value="<%= Session.Contents("�Z��2") %>">
	<input type="hidden" name="���l" value="<%= Session.Contents("���l") %>">
	<input type="hidden" name="FLG" value="2">
	<table align="center" width=30%>
		<tr>
			<td align=center><input type="submit" value="O K"></td>
</FORM>
	<form action="SHINKI.asp" target="Right" method="Post">
			<td align=center><input type="submit" value="�o�^�L�����Z��" id=submit1 name=submit1></td>
	</form>
		</tr>
</table>
</body>
</html>
