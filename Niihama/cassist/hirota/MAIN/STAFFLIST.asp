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
    
	w_SQL="SELECT * FROM M_�Ј� WHERE �g�pFLG=1 ORDER BY 1 ASC"

	w_rRs.Open w_SQL,w_cCn,3,3
	
' SQL���s���̃G���[����
	if Err then
		Session.Contents("ERROR")=Err.description
		Response.Redirect "SQLerror.asp"
	end if
	On Error Goto 0
	
	w_rRs.PageSize = 10

	w_rRs.AbsolutePage =1
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


<table border=1 width=60% align=center bordercolor=#C0C0C0>
<% For iLoop = 1 to w_rRs.PageSize %>
	<tr>
		<td width=20% align=center><%= FixZero(w_rRs("�Ј�CD"),4) %></td>
		<td width=50%><%= w_rRs("�Ј�����") %></td>
	<form action="UPDATE.asp" method="post" target="Right" id=form1 name=form1>
		<td width=10% align=center><input type=hidden name=�Ј�CD value=<%= w_rRs("�Ј�CD") %>>
		<input type="submit" value="�C ��" id=submit1 name=submit1></td>
	</form>
	<form action="DELETE.asp" method="post" target="Right" id=form2 name=form2>
		<td width=10% align=center><input type=hidden name=�Ј�CD value=<%= w_rRs("�Ј�CD") %>>
		<input type="submit" value="�� ��" id=submit2 name=submit2></td>
	</form>
	</tr>
<%
	w_rRs.MoveNext
	Next

	w_rRs.Close
	w_cCn.Close
	Set w_rRs = Nothing
	Set w_cCn = Nothing
%>
</table>
</body>
</html>