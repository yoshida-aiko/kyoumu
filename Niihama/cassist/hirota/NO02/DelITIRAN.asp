<%
	On Error Resume Next
	
	Dim w_cCn,w_rRs,w_SQL
	
' �I�u�W�F�N�g�̒�`
	Set w_cCn = Server.CreateObject("ADODB.Connection")
	Set w_rRs = Server.CreateObject("ADODB.Recordset")

    w_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
    w_rRs.Open "M_�Ј�",w_cCn,2,2
    
' ���N�G�X�g���ꂽ�Ј��f�[�^�𒊏o
	w_SQL="SELECT * FROM M_�Ј� WHERE �Ј�CD=" & Request.Form("�Ј�CD")    

	Set w_rRs = w_cCn.Execute(w_SQL)
	
' SQL���s���̃G���[����
	if Err then
		Session.Contents("ERROR")=Err.description
		Response.Redirect "MsgERROR.asp"
	end if
	On Error Goto 0

'********************************************************************************
'	�[������ : format()
'********************************************************************************
	function FixZero(n, l) 'as string
		FixZero = right(string(l, "0") & n, l)
	end function
%>

<HTML>

<HEAD>
<TITLE>�Ј��Ǘ�</TITLE>

	<base target="Right">

</HEAD>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>
	<h3 align=center>���@�폜��ʁ@��</h3>
	<HR>
	<h5 align=center><font color=Blue>�폜����ƌ��ɖ߂����Ƃ͏o���܂���B<br>�ȉ��̎Ј��f�[�^���폜���Ă��悻�����ł����H</h5>
<form action="SQLexe.asp" method="post" target="Right">
<input type="hidden" name="FLG" value="3">
	<table border=1 align="center" CELLPADDING="8" CELLSPACING="1">
		<tr>
			<td>
				�Ј�CD</td><td align=center><%= FixZero(w_rRs("�Ј�CD"),4) %>
				<input type="hidden" name=�Ј�CD value=<%= w_rRs("�Ј�CD") %>>
			</td>
		</tr>
		<tr>
			<td>
				�Ј�����</td><td align=center><%= w_rRs("�Ј�����") %>
			</td>
		</tr>
		<tr>
			<td>
			<% if Isnull(w_rRs("���N����"))=false then %>
				���N����</td><td align=center><%= w_rRs("���N����") %>
			<% else %>
				���N����</td><td align=center><font color="red">�L������</font>
			<% end if %>
			</td>
		</tr>
		<tr>
			<td>
			<% if Isnull(w_rRs("�d�b�ԍ�1"))=false then %>
				�d�b�ԍ�1</td><td align=center><%= w_rRs("�d�b�ԍ�1") %>
			<% else %>
				�d�b�ԍ�1</td><td align=center><font color="red">�L������</font>
			<% end if %>
			</td>
		</tr>
		<tr>
			<td>
			<% if Isnull(w_rRs("�d�b�ԍ�2"))=false then %>
				�d�b�ԍ�2</td><td align=center><%= w_rRs("�d�b�ԍ�2") %>
			<% else %>
				�d�b�ԍ�2</td><td align=center><font color="red">�L������</font>
			<% end if %>
			</td>
		</tr>
		<tr>
			<td>
			<% if Isnull(w_rRs("�X��"))=false then %>
				�X��</td><td align=center><%= w_rRs("�X��") %>
			<% else %>
				�X��</td><td align=center><font color="red">�L������</font>
			<% end if %>
			</td>
		</tr>
		<tr>
			<td>
			<% if Isnull(w_rRs("�Z��1"))=false then %>
				<% if Isnull(w_rRs("�Z��2"))=false then %>
					�Z��</td><td align=center><%= w_rRs("�Z��1") %><br><%= w_rRs("�Z��2") %>
				<% else %>
					�Z��</td><td align=center><%= w_rRs("�Z��1") %><br>
				<% end if %>
			<% elseif isnull(w_rRs("�Z��2"))=false then %>
					�Z��</td><td align=center><br><%= w_rRs("�Z��2") %>
			<% else %>
					�Z��</td><td align=center><font color="red">�L������</font><br>
			<% end if %>
			</td>
		</tr>
		<tr>
			<td>
			<% if Isnull(w_rRs("���l"))=false then %>
				���l</td><td align=center><pre><%= w_rRs("���l") %></pre>
			<% else %>
				���l</td><td align=center><font color="red">�L������</font>
			<% end if %>
			</td>
		</tr>
	</table>

<h5 align=center><font color=Blue size="2">�悯���OK�{�^���������ĉ������B</font></h5>
	<table align=center width=20%>
		<tr>
			<td align=center>
				<input type="submit" value=" O K ">
			</td>
</form>
		<form action="INitiran.asp" target="Right">
			<td align=center>
				<input type="submit" value="�L�����Z��">
			</td>
		</form>
		</tr>
	</table>

</BODY>
</HTML>

<%
	w_rRs.Close
	Set w_rRs = Nothing
	w_cCn.Close
	Set w_cCn = Nothing
%>
