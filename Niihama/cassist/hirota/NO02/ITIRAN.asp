<%
'----------------------------------------
' �Q�ƒ��̃y�[�W���L���b�V�������Ȃ�
'----------------------------------------
	Response.Expires = 0
	Response.AddHeader "Pragma", "No-Cache"
	Response.AddHeader "Cache-Control", "No-Cache"

	'Response.CacheControl = "No-Cache"
	'Response.AddHeader "Pragma", "No-Cache"
	'Response.Expires = -1

	On Error Resume Next
	Err.Clear
	Dim w_cCn,w_rRs,w_SQL

' �I�u�W�F�N�g�̒�`
	Set w_cCn = Server.CreateObject("ADODB.Connection")
	Set w_rRs = Server.CreateObject("ADODB.Recordset")

    w_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
    w_rRs.Open "M_�Ј�",w_cCn,2,2
    
	w_SQL="SELECT * FROM M_�Ј� WHERE �g�pFLG=1 ORDER BY 1 ASC"

	Set w_rRs = w_cCn.Execute(w_SQL)
	
' SQL���s���̃G���[����
	if Err then
		Session.Contents("ERROR")=Err.description
		Response.Redirect "MsgERROR.asp"
	end if
	On Error Goto 0
	
'********************************************************************************
'		 �[�����ߏ���
'********************************************************************************
	function FixZero(n, l) 'as string
		FixZero = right(string(l, "0") & n, l)
	end function
	
%>
<html>
<head>
	<title>�Ј��Ǘ�</title>
</head>
<!-- <BODY BGCOLOR=#F5F5F5> -->
<body>


<table border=1 width=89% align=center bordercolor=#C0C0C0>
<% Do While not w_rRs.EOF %>
	<tr>
		<td width=20% align=center><%= FixZero(w_rRs("�Ј�CD"),4) %></td>
		<td width=50%><%= w_rRs("�Ј�����") %></td>
	<form action="SYUUSEI.asp" method="post" target="Right">
		<td width=10% align=center><input type=hidden name=�Ј�CD value=<%= w_rRs("�Ј�CD") %>>
		<input type="submit" value="�C ��"></td>
	</form>
	<form action="DelITIRAN.asp" method="post" target="Right">
		<td width=10% align=center><input type=hidden name=�Ј�CD value=<%= w_rRs("�Ј�CD") %>>
		<input type="submit" value="�� ��"></td>
	</form>
	</tr>
<%
	w_rRs.MoveNext
	Loop

	w_rRs.Close
	w_cCn.Close
	Set w_rRs = Nothing
	Set w_cCn = Nothing
%>
</table>
</body>
</html>
