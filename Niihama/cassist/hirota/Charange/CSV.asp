<html>
<head>
<title>�Ј��Ǘ�</title>
	<base target="Right">
	
<!--#INCLUDE FILE="CheckVALUE.asp"-->

<SCRIPT LANGUAGE="VBS">
Function CheckFileName()
	CheckFileName=true
End Function
</SCRIPT>

</head>

<body>

<h3 align=center>�� �Ј��}�X�^CSV�o�� ��</h3>

<p><HR></p>
<br>
<br>
<br>
<form action="ExportCSV.asp" method="post" name="EXPORT">
<h4 align=center>�ȉ��̏����ŁA�Ј��}�X�^��CSV�o�͂��܂��B</h4>
<br>
<table CELLSPACING="0" CELLPADDING="12" ALIGN="CENTER">
<tr>
	<td>���@�Ј�CD</td>
	<td><input type=text name="txtStartCD"size=15 style="ime-mode:inactive" maxlength=4>
		�@�`�@<input type=text name=txtEndCD size=15 style="ime-mode:inactive" maxlength=4></td>
</tr>
<tr>
	<td>���@�Ј�����</td>
	<td>
		<input type=text name="txtName"size=42 maxlength="30" style="ime-mode:active">
	</td>
</tr>
<tr>
	<td>
	</td>
	<td>
		<font color=red>�������܂�����</font>
	</td>
</tr>

<tr>
	<td>���@�폜�t���O</td>
	<td><input type="checkbox" name="checkDel" value="1">�폜�ς݂̃f�[�^�͏o�͂��Ȃ��B
</tr>
</TABLE>
<input type="hidden" name="SQL">
<input type="hidden" name="txtFileName">
<br>
<br>
<br>
<br>
<br>
<hr>
<br>
<table align="center" width=20%>
	<tr>
		<td align=center>
			<INPUT TYPE="submit" VALUE="�o ��" name="Submit"></form>
		</td>
		<td align=center>
			<form action="INitiran.asp" target="Right"><INPUT TYPE="submit" VALUE="�� ��" id=submit2 name=submit2></form>
		</td>
	</tr>
</table>
	</td>
</tr>
</form>
</body>
</html>