<html>
<head>
	<title>�Ј��Ǘ�</title>
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>

<!--#INCLUDE FILE="include01.asp"-->

<hr>
<br>
<h5 align=center><font color=Green>���̃f�[�^��o�^���Ă���낵���ł����H</font></h5>
<form action="SQLexe.asp" method="post" id=form1 name=form1>
	<input type="hidden" name="FLG" value="<%= g_sFLG %>">
	<table border=1 align="center" CELLPADDING="5" CELLSPACING="1">
		<tr>
			<td>
				�Ј�CD</td><td align=center><%= FixZero(w_sCD,4) %>
				<input type="hidden" name=�Ј�CD value=<%= w_sCD %>>
			</td>
		</tr>
		<tr>
			<td>
				�Ј�����</td><td align=center><%= w_sName %>
				<input type="hidden" name=�Ј����� value=<%= w_sName %>>
			</td>
		</tr>
		<tr>
			<td>
				���N����</td><td align=center><%= w_sBirthday %>
				<input type="hidden" name=���N���� value="<%= w_sBirth %>">
			</td>
		</tr>
		<tr>
			<td>
				�d�b�ԍ�1</td><td align=center><%= w_sTelphone1 %>
				<input type="hidden" name=�d�b�ԍ�1 value="<%= w_sTel1 %>">
			</td>
		</tr>
		<tr>
			<td>
				�d�b�ԍ�2</td><td align=center><%= w_sTelphone2 %>
				<input type="hidden" name=�d�b�ԍ�2 value="<%= w_sTel2 %>">
			</td>
		</tr>
		<tr>
			<td>
				�X��</td><td align=center><%= w_sPostPost %>
				<input type="hidden" name=�X�� value="<%= w_sPost %>">
			</td>
		</tr>
		<tr>
			<td>
				�Z��</td><td align=center><%= w_sAdd %>
				<input type="hidden" name=�Z��1 value="<%= w_sAddress1 %>">
				<input type="hidden" name=�Z��2 value="<%= w_sAddress2 %>">
			</td>
		</tr>
		<tr>
			<td>
				���l</td><td align=center><pre><%= w_sIndex %></pre>
				<input type="hidden" name=���l value="<%= w_sBikou %>">
			</td>
		</tr>
	</table>
	<h5 align=center><font color=Green>�悯���OK�{�^���������Ă��������B</font></h5>
	<table align="center" width=20%>
		<tr>
			<td align=center>
				<INPUT TYPE="submit" VALUE=" O K " onclick="Message()">
			</td>
</FORM>
			<td align=center>
				<INPUT TYPE="button" VALUE="�L�����Z��" onClick=history.go(-1) id=button1 name=button1>
			</td>
		</tr>
	</table>
</body>
</html>

<%

'*******************************************************************
'�@�@�[������
'*******************************************************************
function FixZero(n, l) 'as string
	FixZero = right(string(l, "0") & n, l)
end function

%>
