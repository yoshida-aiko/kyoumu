<%
' ��M�����e�L�X�g�̒l��ϐ��ɑ��
	w_sStartCD = Request.Form("txtStartCD")
	w_sEndCD = Request.Form("txtEndCD")
	w_sName = Request.Form("txtName")
	w_checkDel = Request.Form("checkDel")

	On Error Resume Next
    Err.Clear

	Dim g_cCn,g_rRs,w_sSQL,w_Index
	Dim w_sStartCD,w_sEndCD,w_sName,w_checkDel

' �I�u�W�F�N�g�̒�`   
	Set g_cCn = Server.CreateObject("ADODB.Connection")
	Set g_rRs = Server.CreateObject("ADODB.Recordset")
	
    g_cCn.Open "provider=Microsoft.Jet.OLEDB.4.0;" _
                        & "Data Source=\\WEBSVR_2\infogram\hirota\sample2000.mdb"
    g_rRs.Open "M_�Ј�",g_cCn,2,2
    
'--------------------�S�p�𔼊p�ɕϊ�----------------------------------

	Set bobj = Server.CreateObject("basp21")
	w_sStartCD = bobj.StrConv(w_sStartCD,8)	'�S�p�����p�ϊ�
	w_sEndCD = bobj.StrConv(w_sEndCD,8)	'�S�p�����p�ϊ�


'----------------CSV�̃e�L�X�g�̓��͕��������SQL�쐬-----------------------------

    w_sSQL = "SELECT * FROM M_�Ј� WHERE �Ј�CD >=0"
	
	w_sFLG="1"

    If w_sStartCD <> "" Then
        If w_sEndCD <> "" Then
            If gf_bSEICD(w_sStartCD) = False or gf_bSEICD(w_sEndCD) = false Then
                Response.Redirect "MsgCSV.asp"
            End If
            w_sSQL = w_sSQL & " AND �Ј�CD >=" & w_sStartCD & " AND �Ј�CD<=" & w_sEndCD
        Else
            If gf_bSEICD(w_sStartCD) = False Then
               Response.Redirect "MsgCSV.asp"
            End If
            w_sSQL = w_sSQL & " AND �Ј�CD >=" & w_sStartCD
        End If
    ElseIf w_sEndCD <> "" Then
        If gf_bSEICD(w_sEndCD) = False Then
            Response.Redirect "MsgCSV.asp"
        End If
        w_sSQL = w_sSQL & " AND �Ј�CD <=" & w_sEndCD
    End If
    If w_sName <> "" Then
        w_sSQL = w_sSQL & " AND �Ј����� LIKE '%" & w_sName & "%'"
    End If
    If w_checkDel = 1 Then
        w_sSQL = w_sSQL & " AND �g�pFLG=1"
    End If
    w_sSQL = w_sSQL & " ORDER BY 1 ASC"
    
    Set g_rRs = g_cCn.Execute(w_sSQL)
    
' SQL���s���̃G���[����
	if Err then
		Session.Contents("ERROR")=Err.description
		Response.Redirect "MsgERROR.asp"
	end if
	
	On Error Goto 0

'*****************************************************************
'	���̓`�F�b�N�����i�֐��j
'*****************************************************************
function gf_bSEICD(p_sCD)
    gf_bSEICD = false
' �Ј�CD�̓��͌^�������ɂȂ��Ă��邩�H
    If IsNumeric(p_sCD) = False Then
        Exit Function
    End If
' ���������i�J���}�A�����A�����_�A���}�[�N�͎󂯕t���Ȃ��j
    If InStr(p_sCD, ".") <> 0 Then
        Exit Function
    End If
    If InStr(p_sCD, "-") <> 0 Then
        Exit Function
    End If
    If InStr(p_sCD, "+") <> 0 Then
        Exit Function
    End If
    If InStr(p_sCD, ",") <> 0 Then
        Exit Function
    End If
    If InStr(p_sCD, "\") <> 0 Then
        Exit Function
    End If
	if p_sCD < 0 or p_sCD > 9999 then
		Exit Function
	End If
    gf_bSEICD = True
end function

%>

