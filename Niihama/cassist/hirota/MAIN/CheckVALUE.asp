<SCRIPT LANGUAGE="VBS">
<!--�N���C�A���g�T�C�h�̃X�N���v�g

Sub Submit_OnClick()

	w_StartCD = EXPORT.txtStartCD.value
	w_EndCD = EXPORT.txtEndCD.value
	w_Name = EXPORT.txtName.value
	w_CboCheck = EXPORT.checkDel.Checked
	
    SQL = "SELECT * FROM M_�Ј� WHERE �Ј�CD >=0"
    
    if EXPORT.txtFileName.value="" then
		EXPORT.txtFileName.value="Sample"
	end if

' �w��������Ȃ��ꍇ
	if w_StartCD = "" AND w_EndCD = "" AND w_Name = "" AND w_CboCheck = false then
		if CheckFileName()=false then
			Msgbox "�t�@�C�����ɕs���Ȗ��O���g���Ă��܂��B�����̕����͎g�����Ƃ��o���܂���B" _
					& vbcrlf + vbcrlf & "		\ ; : , * < > | ",16,"�Ј����̓��̓G���["
			window.event.returnValue=false
			EXPORT.txtFileName.select
			Exit Sub
		end if
		MsgStr = Msgbox("�w�����������܂���B���ׂẴf�[�^���o�͂��Ă���낵���ł����H",vbOkCancel + vbInformation,"�G�N�X�|�[�g")
			if MsgStr = vbCancel then
				window.event.returnValue = false
				EXPORT.txtStartCD.focus
				Exit Sub
			end if
			SQL = SQL & " ORDER BY 1 ASC"
			EXPORT.SQL.value = SQL
			Exit Sub
	End if
	
' �Ј�CD�̓��̓`�F�b�N
    If w_StartCD <> "" Then
        If w_EndCD <> "" Then
            If gf_bCheckCD(w_StartCD) = False or gf_bCheckCD(w_EndCD) = false Then
                Msgbox "�Ј�CD�ɕ������܂܂�Ă��܂��B��������͂��Ă��������B",16,"�Ј�CD���̓G���["
				window.event.returnValue=false
				EXPORT.txtStartCD.select
				Exit Sub
            End If
            w_StartCD = Cint(w_StartCD)
            w_EndCD = Cint(w_EndCD)
            SQL = SQL & " AND �Ј�CD >=" & w_StartCD & " AND �Ј�CD<=" & w_EndCD
        Else
            If gf_bCheckCD(w_StartCD) = False Then
               Msgbox "�Ј�CD�ɕ������܂܂�Ă��܂��B��������͂��Ă��������B",16,"�Ј�CD���̓G���["
				window.event.returnValue=false
				EXPORT.txtStartCD.select
				Exit Sub
            End If
            w_StartCD = Cint(w_StartCD)
            SQL = SQL & " AND �Ј�CD >=" & w_StartCD
        End If
    ElseIf w_EndCD <> "" Then
        If gf_bCheckCD(w_EndCD) = False Then
            Msgbox "�Ј�CD�ɕ������܂܂�Ă��܂��B��������͂��Ă��������B",16,"�Ј�CD���̓G���["
			window.event.returnValue=false
			EXPORT.txtEndCD.select
			Exit Sub
        End If
        w_EndCD = Cint(w_EndCD)
        SQL = SQL & " AND �Ј�CD <=" & Cint(w_EndCD)
    End If
    
' �Ј����̂̓��̓`�F�b�N
    If w_Name <> "" Then
		if gf_bCheckNAME(w_Name) = false then
			Msgbox "�Ј����̂�"<"�C�܂���">"���܂܂�Ă��܂��B",16,"�Ј����̓��̓G���["
			window.event.returnValue=false
			EXPORT.txtName.select
			Exit Sub
		end if
        SQL = SQL & " AND �Ј����� LIKE '%" & w_Name & "%'"
    End If
    If w_CboCheck = true Then
        SQL = SQL & " AND �g�pFLG=1"
    End If
    if CheckFileName()=false then
		Msgbox "�t�@�C�����ɕs���Ȗ��O���g���Ă��܂��B�����̕����͎g�����Ƃ��o���܂���B" _
				& vbcrlf + vbcrlf & "		\ ; : , * < > | ",16,"�Ј����̓��̓G���["
		window.event.returnValue=false
		EXPORT.txtFileName.select
		Exit Sub
	end if
	
' ���b�Z�[�W
    MsgStr = Msgbox("���̏����ŏo�͂��Ă���낵���ł����H",vbOkCancel + vbInformation,"�G�N�X�|�[�g")
		if MsgStr = vbCancel then
			window.event.returnValue = false
			EXPORT.txtStartCD.focus
			Exit Sub
		End if
	    SQL = SQL & " ORDER BY 1 ASC"
	    EXPORT.SQL.value = SQL
End Sub

'*****************************************************************
'	���̓`�F�b�N�����i�֐��j
'*****************************************************************

Function gf_bCheckCD(p_sCD)
    gf_bCheckCD = false
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
    gf_bCheckCD = True
End Function


'*******************************************************************
'�@�@�^�O�����͂��ꂽ���ǂ����𔻒�
'*******************************************************************
Function gf_bCheckNAME(p_sNAME)

	gf_bCheckNAME = false
    If InStr(p_sNAME, "<") <> 0 Then
        Exit Function
    End If
    If InStr(p_sNAME, ">") <> 0 Then
        Exit Function
    End If
    gf_bCheckNAME = true

End Function
//-->
</SCRIPT>