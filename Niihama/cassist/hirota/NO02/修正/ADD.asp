<%
	On Error Resume Next
	Err.Clear
	Dim w_cCn,w_rRs
	Dim w_sSQL

' �I�u�W�F�N�g��`
	Set w_cCn = Server.CreateObject("ADODB.Connection")
	Set w_rRs = Server.CreateObject("ADODB.Recordset")

 	w_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
 	w_rRs.Open "M_�Ј�",w_cCn,2,2
 	
 	w_sSQL = "SELECT �Ј�CD FROM M_�Ј� WHERE �g�pFLG=0 AND �Ј�CD=" & Request.Form("�Ј�CD")

 	Set w_rRs = w_cCn.Execute(w_sSQL)
 	
 ' SQL���s���̃G���[����
 	if Err then
		Session.Contents("ERROR")=Err.description
		Response.Redirect "MsgERROR.asp"
	end if
	
	On Error Goto 0
	
' �Ј�CD�̏d���`�F�b�N(�g�pFLG=0�̏ꍇ)
 	if w_rRs.EOF=false then
		Session.Contents("�Ј�����")=Request.Form("�Ј�����")
		Session.Contents("���N����")=Request.Form("���N����")
		Session.Contents("�d�b�ԍ�1")=Request.Form("�d�b�ԍ�1")
		Session.Contents("�d�b�ԍ�2")=Request.Form("�d�b�ԍ�2")
		Session.Contents("�X��")=Request.Form("�X��")
		Session.Contents("�Z��1")=Request.Form("�Z��1")
		Session.Contents("�Z��2")=Request.Form("�Z��2")
		Session.Contents("���l")=Request.Form("���l")
		Response.Redirect "Msg03.asp?CD=" & Request.Form("�Ј�CD")
	end if
' �V�K�o�^��SQL���̍쐬	
	w_sSQL = "INSERT INTO M_�Ј� (�Ј�CD,�Ј�����,���N����,�d�b�ԍ�1,�d�b�ԍ�2,"
    w_sSQL = w_sSQL & "�X��,�Z��1,�Z��2,���l,�g�pFLG)"
    w_sSQL = w_sSQL & " VALUES (" & Request.Form("�Ј�CD") & ",'" & Request.Form("�Ј�����") & "'"
    w_sSQL = w_sSQL & "," & Request.Form("���N����")
    w_sSQL = w_sSQL & "," & Request.Form("�d�b�ԍ�1")
    w_sSQL = w_sSQL & "," & Request.Form("�d�b�ԍ�2")
    w_sSQL = w_sSQL & "," & Request.Form("�X��")
    w_sSQL = w_sSQL & "," & Request.Form("�Z��1")
    w_sSQL = w_sSQL & "," & Request.Form("�Z��2")
    w_sSQL = w_sSQL & "," & Request.Form("���l") & ",1)"

'�@�V�K�o�^����
	Set w_rRs = w_cCn.Execute(w_sSQL)
	
 ' SQL���s���̃G���[����
	if Err then
		Session.Contents("ERROR")=Err.description
		Response.Redirect "MsgERROR.asp"
	end if
	On Error Goto 0
	
' �������b�Z�[�W
	Response.Redirect "FinishSHINKI.asp"
		
	'w_rRs.Close
	w_cCn.Close
	Set w_rRs = Nothing
	Set w_cCn = Nothing
		
%>
