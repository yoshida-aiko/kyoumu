<html>
<head>
	<title>�Ј��Ǘ�</title>
	<base target="Right">
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<%
w_sFLG = Request.QueryString("FLG")
Select Case w_sFLG
	Case "1"
%>
		<body>
			<h3 align=center>�� �o�^�G���[ ��</h3>
				<hr><br>
			<h4 align=center><font color=red>���f�[�^�x�[�X�ɓo�^���邱�Ƃ��o���܂���ł����B</font></h4>
		<table align=center width=95%>
			<tr>
				<td valign=Top></td>
				<td>
					���͍��ڂɃG���[������܂����B���L�̏����𖞂������̂̓f�[�^�x�[�X�ɓo�^���邱�Ƃ��o���܂���B
					������x�悭�m���߂Ă����ēx���͂��ĉ������B
				</td>
			</tr>
		</table>
		<br>
		<table align=center width=85%>
			<ul>
			<li><font color=red>���N����</font>��I�ԏꍇ�͕K���N�A���A�������ׂđI�����ĉ������B
			<li><font color=red>���N����</font>�͑��݂��Ȃ����ɂ�����͂��Ȃ��ŉ������B
			<li><font color=red>�Ј�����</font>�A<font color=red>�Z��</font>�A<font color=red>���l</font>��HTML�^�O�Ȃǂ͓��͂��Ȃ��ŉ������B
			<li><font color=red>�d�b�ԍ�</font>����͂��鎞�́A�n�C�t��( - )��؂�œ��͂��ĉ������B
			<li><font color=red>�X�֔ԍ�</font>����͂��鎞�́A 3�� - 4���A�������� 3�� �œ��͂��ĉ������B
			</ul>
			�� �ȏ�̍��ڂ�������x�m���߂������ŁA�ēx�o�^���Ă��������B����ł��o�^����Ȃ��ꍇ��<a href="">�Ǘ���</a>�ɖ₢���킹�Ă��������B
		</table>

		<% Response.Write Request.Form("Msg") %>

		<p align=center><input type="button" value="�߂�" onclick=history.go(-1) id=button1 name=button1>

<% Case "2" %>
		<body>
			<h3 align=center>�� �o�^�G���[ ��</h3>
				<hr><br>
			<h4 align=center><font color=red>���f�[�^�x�[�X�ɓo�^���邱�Ƃ��o���܂���ł����B</font></h4>
		<table align=center>
		<tr>
			<td>
				<ul>
				<li>�Ј�CD <%= Session.Contents("w_sCD") %> �͂��łɎg���Ă��܂��B
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

<% Case "3" %>

		<body>
			<h3 align=center>�� �d���f�[�^ ��</h3>
				<hr><br>
			<h4 align=center><font color=red>���f�[�^�x�[�X�f�[�^�d�����b�Z�[�W</font></h4>
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
			<input type="hidden" name="�Ј�CD" value="<%= Session.Contents("�Ј�CD") %>">
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
					<td align=center><input type="submit" value="O K" id=submit2 name=submit2></td>
		</FORM>
			<form action="SHINKI.asp" target="Right" method="Post" id=form2 name=form2>
					<td align=center><input type="submit" value="�o�^�L�����Z��" id=submit1 name=submit1></td>
			</form>
				</tr>
		</table>
<% Case "4" %>
		<body>
			<h3 align=center>�� �o�^�G���[ ��</h3>
				<hr><br>
			<h4 align=center><font color=red>�����͍��ڂɃG���[������܂��B<br>
					�f�[�^�x�[�X�ɓo�^���邱�Ƃ��o���܂���ł����B</font></h4>
		<br>
		<table align=center width=85%>
			<ul>
			<li><font color=red>�Ј�CD</font>��<font color=red>�Ј�����</font>�͕K���L�����Ă��������B<br>
			<li><font color=red>�Ј�CD</font>�A<font color=red>�Ј�����</font>�ȊO�̍��ڂɊւ��Ă͋L�����Ȃ��Ă��\���܂���B
			</ul>
			�� �ēx�o�^�������s���Ă��������B����ł��o�^����Ȃ��ꍇ��<a href="">�Ǘ���</a>�ɖ₢���킹�Ă��������B
		</table>

		<% Response.Write Request.Form("Msg") %>

		<p align=center><input type="button" value="�߂�" onclick=history.go(-1) id=button1 name=button1>

<% End Select %>
</body>
</html>

