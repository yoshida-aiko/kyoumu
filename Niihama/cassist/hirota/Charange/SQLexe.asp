<%
	On Error Resume Next
	Err.Clear

	Dim w_cCn,w_rRs
	Dim w_sSQL
	
	w_CD=Request.Form("CD")
	w_NAME=Request.Form("NAME")
	w_BIRTHDAY=Request.Form("BIRTHDAY")
	w_TELL1=Request.Form("TELL1")
	w_TELL2=Request.Form("TELL2")
	w_POST=Request.Form("POST")
	w_ADDRESS1=Request.Form("ADDRESS1")
	w_ADDRESS2=Request.Form("ADDRESS2")
	w_BIKOU=Request.Form("BIKOU")

' �I�u�W�F�N�g��`
	Set w_cCn = Server.CreateObject("ADODB.Connection")
	Set w_rRs = Server.CreateObject("ADODB.Recordset")
 	w_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
 	w_rRs.Open "M_�Ј�",w_cCn,2,2

 
	if Session.Contents("FLG")="ADDNEW" then
	
	' �Ј�CD�d���`�F�b�N
		w_sSQL = "SELECT �Ј�CD FROM M_�Ј� WHERE �g�pFLG=1 AND �Ј�CD =" & w_CD
		if gf_SQLexe(w_sSQL)=false then
			Response.Redirect "SQLerror.asp"
		end if
	' �d���`�F�b�N
		if w_rRs.EOF=false then
			Session.Contents("ErrorCD")=w_CD
			Response.Redirect "WStaffMsg.asp?WStaff=1"
		end if
	' �g�pFLG=0�̎Ј��f�[�^���c���Ă���ꍇ
		w_sSQL = "SELECT �Ј�CD FROM M_�Ј� WHERE �g�pFLG=0 AND �Ј�CD=" & w_CD
		if gf_SQLexe(w_sSQL)=false then
			Response.Redirect "SQLerror.asp"
		end if
	' �Ј�CD�̏d���`�F�b�N(�g�pFLG=0�̏ꍇ)
	 	if w_rRs.EOF=false then
			Session.Contents("�Ј�CD")=w_CD
			Session.Contents("�Ј�����")=w_NAME
			Session.Contents("���N����")=w_BIRTHDAY
			Session.Contents("�d�b�ԍ�1")=w_TELL1
			Session.Contents("�d�b�ԍ�2")=w_TELL2
			Session.Contents("�X��")=w_POST
			Session.Contents("�Z��1")=w_ADDRESS1
			Session.Contents("�Z��2")=w_ADDRESS2
			Session.Contents("���l")=w_BIKOU
			Response.Redirect "WStaffMsg.asp?WStaff=2"
		end if
		if f_ADDNEW()=false then
			Response.Redirect "SQLerror.asp"
		end if
		
	elseif Session.Contents("FLG")="UPDATE" then
		if f_UPDATE()=false then
			Response.Redirect "SQLerror.asp"
		end if
	else
		if f_DELETE()=false then
			Response.Redirect "SQLerror.asp"
		end if
	end if
	
	Response.Redirect "INitiran.asp"

	w_rRs.Close
	w_cCn.Close
	Set w_rRs = Nothing
	Set w_cCn = Nothing



'**************************************************************
'			�V�K
'**************************************************************
Function f_ADDNEW()
	f_ADDNEW=false
' �V�K�o�^��SQL���̍쐬	
	w_sSQL = "INSERT INTO M_�Ј� (�Ј�CD,�Ј�����,���N����,�d�b�ԍ�1,�d�b�ԍ�2,"
    w_sSQL = w_sSQL & "�X��,�Z��1,�Z��2,���l,�g�pFLG)"
    w_sSQL = w_sSQL & " VALUES (" & Request.Form("CD")
    w_sSQL = w_sSQL & "," & Request.Form("NAME")
    w_sSQL = w_sSQL & "," & Request.Form("BIRTHDAY")
    w_sSQL = w_sSQL & "," & Request.Form("TELL1")
    w_sSQL = w_sSQL & "," & Request.Form("TELL2")
    w_sSQL = w_sSQL & "," & Request.Form("POST")
    w_sSQL = w_sSQL & "," & Request.Form("ADDRESS1")
    w_sSQL = w_sSQL & "," & Request.Form("ADDRESS2")
    w_sSQL = w_sSQL & "," & Request.Form("BIKOU")
    w_sSQL = w_sSQL & ",1)"

	if gf_SQLexe(w_sSQL)=false then
		Exit Function
	end if
	f_ADDNEW=true
End Function



'**************************************************************
'			�C��
'**************************************************************
Function f_UPDATE()
		f_UPDATE=false
	' �C��SQL
		w_sSQL ="UPDATE M_�Ј� SET �Ј�����=" & Request.Form("NAME")
		w_sSQL = w_sSQL & ",���N����=" & Request.Form("BIRTHDAY")
		w_sSQL = w_sSQL & ",�d�b�ԍ�1=" & Request.Form("TELL1")
		w_sSQL = w_sSQL & ",�d�b�ԍ�2=" & Request.Form("TELL2")
		w_sSQL = w_sSQL & ",�X��=" & Request.Form("POST")
		w_sSQL = w_sSQL & ",�Z��1=" & Request.Form("ADDRESS1")
		w_sSQL = w_sSQL & ",�Z��2=" & Request.Form("ADDRESS2")
		w_sSQL = w_sSQL & ",���l=" & Request.Form("BIKOU")
		w_sSQL = w_sSQL & ",�g�pFLG=1 WHERE �Ј�CD=" & Request.Form("CD")
		
		if gf_SQLexe(w_sSQL)=false then
			Exit Function
		end if
		f_UPDATE=true
End Function



'**************************************************************
'			�폜
'**************************************************************
Function f_DELETE()
	    f_DELETE=false
	' �Ј��f�[�^�폜��SQL��
	    w_sSQL = "UPDATE M_�Ј� SET �g�pFLG=0 WHERE �Ј�CD=" & Request.Form("CD")

		if gf_SQLexe(w_sSQL)=false then
			Exit Function
		end if
		f_DELETE=true
End Function



'**************************************************************
'			SQL���s�֐�
'**************************************************************

Function gf_SQLexe(p_sSQL)
	Set w_rRs = w_cCn.Execute(p_sSQL)
 ' SQL���s���̃G���[����
	if Err then
		Session.Contents("SQLerror")=Err.Description
		gf_SQLexe=false
	end if
	On Error Goto 0
	gf_SQLexe=true
End Function

%>