<%
	On Error Resume Next
	Err.Clear
	Dim w_cCn,w_rRs,SQL

' �I�u�W�F�N�g��`
	Set w_cCn = Server.CreateObject("ADODB.Connection")
	Set w_rRs = Server.CreateObject("ADODB.Recordset")

    w_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
    w_rRs.Open "M_�Ј�",w_cCn,2,2
    
' �Ј��f�[�^�폜��SQL��
    SQL = "UPDATE M_�Ј� SET �g�pFLG=0 WHERE �Ј�CD=" & Request.Form("�Ј�CD")
    
    Set w_rRs = w_cCn.Execute(SQL)
    
' SQL���s���̃G���[����
	if Err then
		Session.Contents("ERROR")=Err.description
		Response.Redirect "MsgERROR.asp"
	end if
	
	On Error Goto 0
	
' �������b�Z�[�W
	Response.Redirect "FinishDel.asp"
	
    w_rRs.Close
	Set w_rRs = Nothing
	w_cCn.Close
	Set w_cCn = Nothing

%>
