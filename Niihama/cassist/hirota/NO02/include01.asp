<%
	On Error Resume Next
	Err.Clear
	Dim w_sCD,w_sName,w_sYEAR,w_sMONTH,w_sDAY,w_sTel1,w_sTel2,w_sTel3,w_sTel4,w_sTel5,w_sTel6
	Dim w_sPost1,w_sPost2,w_sAddress1,w_sAddress2,w_sBikou

' �e�L�X�g�ɓ��͂��ꂽ�f�[�^��ϐ��Ɋi�[����B
	w_sCD = Request.Form("w_sCD")
	w_sName = Request.Form("w_sName")
	w_sYEAR = Request.Form("w_sYEAR")
	w_sMONTH = Request.Form("w_sMONTH")
	w_sDAY = Request.Form("w_sDAY")
	w_sTel1 = Request.Form("w_sTel1")
	w_sTel2 = Request.Form("w_sTel2")
	w_sTel3 = Request.Form("w_sTel3")
	w_sTel4 = Request.Form("w_sTel4")
	w_sTel5 = Request.Form("w_sTel5")
	w_sTel6 = Request.Form("w_sTel6")
	w_sPost1 = Request.Form("w_sPost1")
	w_sPost2 = Request.Form("w_sPost2")
	w_sAddress1 = Request.Form("w_sAddress1")
	w_sAddress2 = Request.Form("w_sAddress2")
	w_sBikou = Request.Form("w_sBikou")	

Function SHINKI()
	'**************************************************************
	'			�V�@�K
	'**************************************************************
	
		Dim g_cCn,g_rRs,SQL
		
		g_sFLG="1"
		Response.Write "<h3 align=center>�� �V�K�m�F��� ��</h3>"
		
	' �Ј�CD�A�Ј����̂̓��̓`�F�b�N
		if w_sCD = "" or w_sName = "" then
			w_FLG = "4" '(���̓G���[�t���O : 4 )
			Response.Redirect "Msg.asp?FLG=" & w_FLG
		end if
		
	' �I�u�W�F�N�g��`
		Set g_cCn = Server.CreateObject("ADODB.Connection")
		Set g_rRs = Server.CreateObject("ADODB.Recordset")
		g_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
		                    & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
		g_rRs.Open "M_�Ј�",g_cCn,2,2
		
	' �Ј�CD�d���`�F�b�N
		SQL="SELECT �Ј�CD FROM M_�Ј� WHERE �g�pFLG=1 AND �Ј�CD=" & w_sCD
		Set g_rRs = g_cCn.Execute(SQL)
		
	' SQL���s���̃G���[����
		if Err then
			Session.Contents("ERROR")=Err.description
			Response.Redirect "MsgERROR.asp"
		end if	
		On Error Goto 0
		
	' �d���`�F�b�N
		if g_rRs.EOF=false then
			w_FLG="2" '(�d�����b�Z�[�W�t���O : 2 )
			Session.Contents("w_sCD")=w_sCD
			Response.Redirect "Msg.asp?FLG=" & w_FLG
		end if
End Function
	'**************************************************************
	'			�C�@��
	'**************************************************************
Function SYUUSEI()
		g_sFLG="2"
		w_sCD=Request.Form("CD")
		Response.Write "<h3 align=center>�� �C���m�F��� ��</h3>"
End Function

%>