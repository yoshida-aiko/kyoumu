<%
	On Error Resume Next
	Err.Clear
	Dim w_cCn,w_rRs,SQL

' �I�u�W�F�N�g�̒�`
	Set w_cCn = Server.CreateObject("ADODB.Connection")
	Set w_rRs = Server.CreateObject("ADODB.Recordset")

    w_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
    w_rRs.Open "M_�Ј�",w_cCn,2,2
	
' �C��SQL
	SQL ="UPDATE M_�Ј� SET �Ј�����='" & Request.Form("�Ј�����") & "'"
	SQL = SQL & ",���N����=" & Request.Form("���N����")
	SQL = SQL & ",�d�b�ԍ�1=" & Request.Form("�d�b�ԍ�1")
	SQL = SQL & ",�d�b�ԍ�2=" & Request.Form("�d�b�ԍ�2")
	SQL = SQL & ",�X��=" & Request.Form("�X��")
	SQL = SQL & ",�Z��1=" & Request.Form("�Z��1")
	SQL = SQL & ",�Z��2=" & Request.Form("�Z��2")
	SQL = SQL & ",���l=" & Request.Form("���l")
	SQL = SQL & ",�g�pFLG=1 WHERE �Ј�CD=" & Request.Form("�Ј�CD")
	
	Set w_rRs = w_cCn.Execute(SQL)
	
' SQL���s���̃G���[����
	if Err then
		Session.Contents("ERROR")=Err.description
		Response.Redirect "MsgERROR.asp"
	end if
	
	On Error Goto 0

	'w_rRs.Close
	w_cCn.Close
	Set w_rRs = Nothing
	Set w_cCn = Nothing
	
' �I�����b�Z�[�W
	Response.Redirect "FinishSYUUSEI.asp"
	
%>
