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
 	
 e_SELECT = Request.Form("FLG")
 Select case e_SELECT
	'**************************************************************
	'			�V�K
	'**************************************************************
	 case "1"
	 
	 ' �g�pFLG=0�̎Ј��f�[�^���c���Ă���ꍇ
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
			Session.Contents("�Ј�CD")=Request.Form("�Ј�CD")
			w_FLG="3"
			Response.Redirect "Msg.asp?FLG=" & w_FLG
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

		if gf_SQLexe(w_sSQL)=false then
			Session.Contents("ERROR")=Err.description
			Response.Redirect "MsgERROR.asp"
		end if
		w_sFLG = "1"
		
	'**************************************************************
	'			�C��
	'**************************************************************
	case "2"
		
	' �C��SQL
		w_sSQL ="UPDATE M_�Ј� SET �Ј�����='" & Request.Form("�Ј�����") & "'"
		w_sSQL = w_sSQL & ",���N����=" & Request.Form("���N����")
		w_sSQL = w_sSQL & ",�d�b�ԍ�1=" & Request.Form("�d�b�ԍ�1")
		w_sSQL = w_sSQL & ",�d�b�ԍ�2=" & Request.Form("�d�b�ԍ�2")
		w_sSQL = w_sSQL & ",�X��=" & Request.Form("�X��")
		w_sSQL = w_sSQL & ",�Z��1=" & Request.Form("�Z��1")
		w_sSQL = w_sSQL & ",�Z��2=" & Request.Form("�Z��2")
		w_sSQL = w_sSQL & ",���l=" & Request.Form("���l")
		w_sSQL = w_sSQL & ",�g�pFLG=1 WHERE �Ј�CD=" & Request.Form("�Ј�CD")
		
		if gf_SQLexe(w_sSQL)=false then
			Session.Contents("ERROR")=Err.description
			Response.Redirect "MsgERROR.asp"
		end if
		w_sFLG="2"	

	'**************************************************************
	'			�폜
	'**************************************************************
	case "3"
	    
	' �Ј��f�[�^�폜��SQL��
	    w_SQL = "UPDATE M_�Ј� SET �g�pFLG=0 WHERE �Ј�CD=" & Request.Form("�Ј�CD")

		if gf_SQLexe(w_SQL)=false then
			Session.Contents("ERROR")=Err.description
			Response.Redirect "MsgERROR.asp"
		end if
		w_sFLG="3"
	
end Select

Response.Redirect "FinishMsg.asp?FLG=" & w_sFLG

w_rRs.Close
w_cCn.Close
Set w_rRs = Nothing
Set w_cCn = Nothing

'**************************************************************
'			SQL���s�֐�
'**************************************************************
function gf_SQLexe(p_sSQL)
	Set w_rRs = w_cCn.Execute(p_sSQL)
 ' SQL���s���̃G���[����
	if Err then
		gf_SQLexe=false
	end if
	On Error Goto 0
	gf_SQLexe=true
end function

%>