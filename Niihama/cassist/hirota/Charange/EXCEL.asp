<html>
<head>
<title>�Ј��Ǘ�</title>
	<base target="Right">

<!--#INCLUDE FILE="CheckVALUE.asp"-->

<SCRIPT LANGUAGE="VBS">
' �t�@�C�����̓��̓`�F�b�N
Function CheckFileName()

	Dim CheckNameStr
	Dim NameStr
	Dim i
	CheckFileName=false
	CheckNameStr = "\/:;,*?<>|"
	if EXPORT.txtFileName.value = "" then
		Exit Function
	else
		For i = 1 to Len(EXPORT.txtFileName.value)
			NameStr=Mid(EXPORT.txtFileName.value,i,1)
			if Instr(CheckNameStr,NameStr) <> 0 then
				Exit Function
			end if
		Next
	end if
	CheckFileName=true
	
End Function
</SCRIPT>

</head>
<body>

<h3 align=center>�� �Ј��}�X�^EXCEL�o�� ��</h3>

<p><HR></p>


<br>
<h4 align=center>�ȉ��̏����ŁA�Ј��}�X�^��EXCEL�o�͂��܂��B</h4>
<br>
<form action="ExportEXCEL.asp" method="post" name="EXPORT">
<table CELLSPACING="0" CELLPADDING="12" ALIGN="CENTER">
	<tr>
		<td>���@�Ј�CD</td>
		<td><input type=text name="txtStartCD"size=15 style="ime-mode:inactive" maxlength=4>
			�@�`�@<input type=text name=txtEndCD size=15 style="ime-mode:inactive" maxlength=4></td>
	</tr>
	<tr>
		<td>���@�Ј�����</td>
		<td><input type=text name="txtName"size=42 maxlength="30" style="ime-mode:active"></td>
	</tr>
	<tr>
		<td></td>
		<td><font color=red>�������܂�����</font></td>
	</tr>
	<tr>
		<td>���@�폜�t���O</td>
		<td><input type="checkbox" name="checkDel" value="1">�폜�ς݂̃f�[�^�͏o�͂��Ȃ��B</td>
	</tr>
	<tr>
		<td>���@�o�͐�</td>
		<td><select name="cboName">
			<Option value="C:\WINDOWS\�޽�į��\">�޽�į��
			<Option value="Y:\�{��\">�{�䂳��<Option value="Y:\�A�c\">�A�c<Option value="Y:\���\">���
			<Option value="Y:\���c\">���c
			</select>
		</td>
	</tr>
	<tr>
		<td>���@�ۑ��t�@�C����</td>
		<td><input type="text" name="txtFileName" size=10 value="Sample" Maxlength="10" style="ime-mode:active">.xls<br></td>
	</tr>
</TABLE>
<input type="hidden" name="SQL">
<br>
<br>
<br>
<hr>
<br>
<table align="center" width=20%>
	<tr>
		<td align=center><INPUT TYPE="submit" VALUE="�o ��" name="Submit"></form></td>
		<td align=center><form action="INitiran.asp" target="Right">
				<INPUT TYPE="submit" VALUE="�� ��" id=submit2 name=submit2></form></td>
	</tr>
</table>
	</td>
</tr>
</form>
</body>
</html>