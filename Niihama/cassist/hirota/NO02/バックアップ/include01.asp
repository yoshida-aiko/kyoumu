<%
	On Error Resume Next
	Err.Clear
	Dim w_sCD,w_sName,w_sYEAR,w_sMONTH,w_sDAY,w_sTel1,w_sTel2,w_sTel3,w_sTel4,w_sTel5,w_sTel6
	Dim w_sPost1,w_sPost2,w_sAddress1,w_sAddress2,w_sBikou

' �e�L�X�g�ɓ��͂��ꂽ�f�[�^��ϐ��Ɋi�[����B
	w_sCD = Request.Form("txtCD")
	w_sName = Request.Form("txtName")
	w_sYEAR = Request.Form("txtYEAR")
	w_sMONTH = Request.Form("txtMONTH")
	w_sDAY = Request.Form("txtDAY")
	w_sTel1 = Request.Form("txtTel1")
	w_sTel2 = Request.Form("txtTel2")
	w_sTel3 = Request.Form("txtTel3")
	w_sTel4 = Request.Form("txtTel4")
	w_sTel5 = Request.Form("txtTel5")
	w_sTel6 = Request.Form("txtTel6")
	w_sPost1 = Request.Form("txtPost1")
	w_sPost2 = Request.Form("txtPost2")
	w_sAddress1 = Request.Form("txtAddress1")
	w_sAddress2 = Request.Form("txtAddress2")
	w_sBikou = Request.Form("txtBikou")	

w_sFLG = Request.Form("FLG")
Select Case w_sFLG

	'**************************************************************
	'			�V�@�K
	'**************************************************************
	Case "1"
	
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
		
	'**************************************************************
	'			�C�@��
	'**************************************************************
	Case "2"
		g_sFLG="2"
		w_sCD=Request.Form("CD")
		Response.Write "<h3 align=center>�� �C���m�F��� ��</h3>"
end Select

	'**************************************************************
	'			���@�́@���@��
	'**************************************************************
	w_FLG="1"

' ���O�`�F�b�N�iHTML�^�O�𖄂ߍ��܂�Ă��Ȃ����j
	if f_CheckVALUE(w_sName)=false then
		Response.Redirect "Msg.asp?FLG=" & w_sFLG
	end if
	
' ���N�����`�F�b�N
	if w_sYEAR <> "" AND w_sMONTH <> "" AND w_sDAY <> "" then
		if IsDate(w_sYEAR & "/" & w_sMONTH & "/" & w_sDAY)=false then
			Response.Redirect "Msg.asp"
		else
			w_sBirthday = w_sYEAR & "�N" & w_sMONTH & "��" & w_sDAY & "��"
			w_sBirth = "'" & w_sYEAR & "/" & w_sMONTH & "/" & w_sDAY & "'"
		end if
	elseif w_sYEAR ="" AND w_sMONTH = "" AND w_sDAY = "" then
		w_sBirthday="<font color=red>�L������</font>"
		w_sBirth = "NULL"
	else
		Response.Redirect "Msg.asp?FLG=" & w_FLG
	end if

' �d�b�ԍ�1�`�F�b�N
	if w_sTel1 <> "" AND w_sTel2 <> "" AND w_sTel3 <> "" then
		w_sTelphone1= w_sTel1 & "-" & w_sTel2 & "-" & w_sTel3
		w_sTel1="'" & w_sTel1 & "-" & w_sTel2 & "-" & w_sTel3 & "'"
	elseif w_sTel1 = "" AND w_sTel2 = "" AND w_sTel3 = "" then
		w_sTelphone1="<font color=red>�L������</font>"
		w_sTel1 ="NULL"
	else
		Response.Redirect "Msg.asp?FLG=" & w_FLG
	end if

' �d�b�ԍ�2�`�F�b�N
	if w_sTel4 <> "" AND w_sTel5 <> "" AND w_sTel6 <> "" then
		w_sTelphone2=w_sTel4 & "-" & w_sTel5 & "-" & w_sTel6
		w_sTel2="'" & w_sTel4 & "-" & w_sTel5 & "-" & w_sTel6 & "'"
	elseif w_sTel4 = "" AND w_sTel5 = "" AND w_sTel6 = "" then
		w_sTelphone2="<font color=red>�L������</font>"
		w_sTel2 ="NULL"
	else
		Response.Redirect "Msg.asp?FLG=" & w_FLG
	end if

' �X�֔ԍ��`�F�b�N
	if w_sPost1 = "" then
		if w_sPost2 = "" then
			w_sPostPost="<font color=red>�L������</font>"
			w_sPost = "NULL"
		else
			Response.Redirect "Msg.asp?FLG=" & w_FLG
		end if
	elseif w_sPost2 = "" then
		if Len(w_sPost1) < 3 then
			Response.Redirect "Msg.asp?FLG=" & w_FLG
		end if
		w_sPostPost=w_sPost1
		w_sPost= "'" & w_sPost1 & "'"
	else
		if Len(w_sPost1) < 3 or Len(w_sPost2) < 4 then
			Response.Redirect "Msg.asp?FLG=" & w_FLG
		end if
		w_sPostPost=w_sPost1 & " - " & w_sPost2
		w_sPost= "'" & w_sPost1 & "-" & w_sPost2 & "'"
	end if

' �Z��1�A�Z��2�`�F�b�N
	if w_sAddress1 <> "" then
		if w_sAddress2 <> "" then
			if f_CheckVALUE(w_sAddress1)=false or f_CheckVALUE(w_sAddress2)=false then
				Response.Redirect "Msg.asp?FLG=" & w_FLG
			end if
			w_sAdd =w_sAddress1 & "<br>" & w_sAddress2
			w_sAddress1= "'" & w_sAddress1 & "'"
			w_sAddress2= "'" & w_sAddress2 & "'"
		else
			if f_CheckVALUE(w_sAddress1)=false then
					Response.Redirect "Msg.asp?FLG=" & w_FLG
			end if
			w_sAdd=w_sAddress1 & "<br>"
			w_sAddress1= "'" & w_sAddress1 & "'"
			w_sAddress2= "NULL"	
		end if
	elseif w_sAddress2 <> "" then
		if f_CheckVALUE(w_sAddress2)=false then
			Response.Redirect "Msg.asp?FLG=" & w_FLG
		end if
		w_sAdd= "<br>" & w_sAddress2
		w_sAddress1="NULL"
		w_sAddress2= "'" & w_sAddress2 & "'"
	else
		w_sAdd="<font color=red>�L������</font><br>"
		w_sAddress1="NULL"
		w_sAddress2="NULL"
	end if

' ���l�`�F�b�N
	if w_sBikou <> "" then
		if f_CheckVALUE(w_sBikou)=false then
			Response.Redirect "Msg.asp?FLG=" & w_FLG
		end if
		w_sIndex=w_sBikou
		w_sBikou= "'" & w_sBikou & "'"
	else
		w_sIndex="<font color=red>�L������</font>"
		w_sBikou="NULL"
	end if
	
'*******************************************************************
'�@�@�^�O�����͂��ꂽ���ǂ����𔻒�
'*******************************************************************
function f_CheckVALUE(p_VALUE)
	f_CheckVALUE = false
    If InStr(p_VALUE, "<") <> 0 Then
        Exit Function
    End If
    If InStr(p_sCD, ">") <> 0 Then
        Exit Function
    End If
    f_CheckVALUE = true
end function
%>