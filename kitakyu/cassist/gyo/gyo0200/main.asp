<%@ Language=VBScript %>
<%
'*************************************************************************
'* �V�X�e����: ���������V�X�e��
'* ��  ��  ��: �s�������ꗗ
'* ��۸���ID : gyo/gyo0200/main.asp
'* �@      �\: ���y�[�W �s�������}�X�^�̈ꗗ���X�g�\�����s��
'*-------------------------------------------------------------------------
'* ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'*           :�����N�x       ��      SESSION���i�ۗ��j
'*           cboGyojiDate      :�I�������s�����t
'*           chkGyojiCd      :�s���R�[�h
'*          txtMode             :���샂�[�h
'* ��      ��:�Ȃ�
'* ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'*           :�����N�x       ��      SESSION���i�ۗ��j
'* ��      ��:
'*           �������\��
'*               ���������ɂ��Ȃ��s�������ꗗ��\��
'*           ���s���̂ݕ\���`�F�b�N�{�b�N�XON��
'*               �s���̂ݕ\��
'*-------------------------------------------------------------------------
'* ��      ��: 2001/06/26 ���{ ����
'* ��      �X: 2001/07/27 �ɓ����q  M40_CALENDER�e�[�u���폜�ɑΉ�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    'Public  m_bErrMsg           '�װү����
    Public  m_sMsg              'ү����

    '�擾�����f�[�^�����ϐ�
    Public  m_iKyokanCd         ':�����R�[�h
    Public  m_iSyoriNen         ':�����N�x
    Public  m_iGyojiM          ':�s�������ꗗ��
    Public  m_iGyojiFlg         ':�\���p
    
    Public  m_iDate            ':�\������
    Public  m_iYear            ':�\���N
    Public  m_iMonth            ':�\����
    Public  m_iDay              ':�\����
    Public  m_sYobi             ':�\���j��
    Public  m_iYobiCd           ':�j���R�[�h
    Public  m_iKyujituFlg       ':�x���R�[�h�iDB�j
    Public  m_sColor            ':�\���F�i�e�[�u���w�i�p�j
    Public  m_iColorCd          ':�\���F�R�[�h�i�e�[�u���w�i�p�j
    Public  m_iGyojiCd          ':�s���R�[�h
    Public  m_sGyojiMei         ':�s����
    Public  m_sBiko             ':���l
    Public  m_iKaisibi          ':�s���J�n��
    Public  m_iSyuryobi         ':�s���I����
    Public  m_iHyojiFlg         ':�\���t���O
    Public  m_iNKaisiDate       ':�N�x�J�n��
    Public  m_iNKaisibi         ':�N�x�J�n��(��)

	'//�w���֘A���
    Public  m_sGakki,m_sZenki_Start,m_sKouki_Start,m_sKouki_End

    Public  m_Rs                'recordset

    '�y�[�W�֌W
    Public  m_iMax              ':�ő�y�[�W
    Public  m_iDsp              '// �ꗗ�\���s��
    Public  m_iCount            '//1������̍s����
    Public  m_iCountN           '//1������̍s�����iN�Ԗځj

    '�f�[�^�擾�p
'    Public  Const C_NENDO_KAISITUKI = 4             '�N�x�J�n��
    'Public  Const C_NENDO_KAISITUKI_MATUBI = 30     '�N�x�J�n������

'///////////////////////////���C������/////////////////////////////

    'Ҳ�ٰ�ݎ��s
    Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////

Sub Main()
'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_sWHERE            '// WHERE��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    Dim w_iRecCount         '//���R�[�h�J�E���g�p

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�s�������ꗗ"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iDsp = C_PAGE_LINE

    Do
        '// �ް��ް��ڑ�
        w_iRet = gf_OpenDatabase()
        If w_iRet <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

        '// �l�̏�����
        Call s_SetBlank()

        '// ���Ұ�SET
        Call s_SetParam()

        '// �N�x�J�n�����擾
        Call f_GetNendoKaisibi()

		'//�w�������擾
		w_iRet = gf_GetGakkiInfo(m_sGakki,m_sZenki_Start,m_sKouki_Start,m_sKouki_End)
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

		'//�s�����׃e�[�u�����A�I�����ꂽ���̃J�����_�[���擾
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT T32_GYOJI_M.T32_HIDUKE"
		w_sSQL = w_sSQL & vbCrLf & " FROM T32_GYOJI_M"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T32_GYOJI_M.T32_NENDO=" & cInt(m_iSyoriNen)
		w_sSQL = w_sSQL & vbCrLf & "  AND SUBSTR(T32_HIDUKE,6,2)='" & gf_fmtZero(m_iGyojiM,2) & "'"
		w_sSQL = w_sSQL & vbCrLf & " GROUP BY T32_GYOJI_M.T32_HIDUKE"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY SUBSTR(T32_HIDUKE,9,2)"

        Set m_Rs = Server.CreateObject("ADODB.Recordset")
        w_iRet = gf_GetRecordsetExt(m_Rs, w_sSQL, m_iDsp)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_bErrFlg = True
            m_sErrMsg = "���R�[�h�Z�b�g�̎擾�Ɏ��s���܂���"
            Exit Do
        End If

        '// �y�[�W��\��
        Call showPage()
        Exit Do

    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    gf_closeObject(m_Rs)
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂ��󔒂ɏ�����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetBlank()

    m_iKyokanCd = ""
    m_iSyoriNen = ""
    m_iGyojiM = ""
    m_iGyojiFlg = ""

    m_iGyojiCd = ""
    m_sGyojiMei = ""
    m_sBiko = ""
    m_iHyojiFlg = ""
    
    m_sYobi = ""
    m_iKyujituFlg = ""
    m_iYobiCd = ""

    m_iDay = ""
    m_iMonth = ""
    m_iYear = ""
    m_iDate = ""
    m_sColor = ""
    
    m_iKaisibi = ""
    m_iSyuryobi = ""
    m_iNKaisiDate = ""
    m_iNKaisibi = ""
    
    m_iCount = ""
    m_iCountN = ""
    
End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iKyokanCd = Session("KYOKAN_CD")         ':�����R�[�h
    m_iSyoriNen = Session("NENDO")     ':�����N�x
    m_iGyojiM = Request("cboGyojiDate")        ':�\����
    m_iGyojiFlg = 0                             ':�s���\���p

	if Request("chkGyojiCd") = "on" Then
		m_iGyojiFlg = 1
		else
	end if

End Sub

'********************************************************************************
'*  [�@�\]  �s�����̎擾
'*  [����]  
'*  [�ߒl]  0:���擾�����A1:���R�[�h�Ȃ��A99:���s
'*  [����]  
'********************************************************************************
Function f_GetGyojiMei()
    
    Dim w_Rs                '// ں��޾�ĵ�޼ު��
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    
    On Error Resume Next
    Err.Clear
    
    f_GetGyojiMei = 0

    Do

        m_iCount = 0
        m_iCountN = 0

        '// �s���w�b�_ں��޾�Ă��擾
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT "
        w_sSQL = w_sSQL & "T31_GYOJI_CD"
        w_sSQL = w_sSQL & ",T31_GYOJI_MEI"
        w_sSQL = w_sSQL & ",T31_BIKO"
        w_sSQL = w_sSQL & ",T31_KAISI_BI"
        w_sSQL = w_sSQL & ",T31_SYURYO_BI"
        w_sSQL = w_sSQL & ",T31_HYOJI_FLG"
        w_sSQL = w_sSQL & " FROM T31_GYOJI_H "
        w_sSQL = w_sSQL & " WHERE T31_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & " AND T31_KAISI_BI <= '" & m_iDate & "'"
        w_sSQL = w_sSQL & " AND T31_SYURYO_BI >= '" & m_iDate & "'"

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If w_iRet <> 0 Then
           'ں��޾�Ă̎擾���s
            m_sGyojiMei = "�@"
            m_sBiko = "�@"
            m_sErrMsg = "ں��޾�Ă̎擾���s"
            m_bErrFlg = True
            f_GetGyojiMei = 99
            Exit Do
        Else
        End If

        If w_Rs.EOF Then
            '�Ώ�ں��ނȂ�
            m_sGyojiMei = "�@"
            m_sBiko = "�@"
            f_GetGyojiMei = 1
            Exit Do
        End If

        Do Until w_Rs.EOF
            '// �擾�����l���i�[
            m_iHyojiFlg = w_Rs("T31_HYOJI_FLG")     '//�\���t���O���i�[
            m_iKaisibi = w_Rs("T31_KAISI_BI")       '//�s���J�n�����i�[
            m_iGyojiCd = w_Rs("T31_GYOJI_CD")       '//�s���R�[�h���i�[

                if f_ChkKaisibi = 0 Then            '//�J�n���̏ꍇ�\��
                    m_iCount = m_iCount + 1         '//�\���������J�E���g
                else
                    if f_ChkHiduke = 0 Then         '//�y���\���`�F�b�N
                        if f_ChkHyojibi = 0 Then        '//�J�n���ȊO�ł��S�\���w��̏ꍇ�͕\��
                            m_iCount = m_iCount + 1     '//�\���������J�E���g
                        else
                        end if
                    else
                        Exit Do
                    end if
                end if


            w_Rs.MoveNext
        Loop

        if m_iCount = 0 Then
            f_GetGyojiMei = 1
            Exit Do
        end if

        w_Rs.MoveFirst

        Do Until w_Rs.EOF
            '// �擾�����l���i�[

            m_iHyojiFlg = ""
            m_iKaisibi  = ""
            m_iGyojiCd  = ""
            m_iHyojiFlg = w_Rs("T31_HYOJI_FLG")     '//�\���t���O���i�[
            m_iKaisibi  = w_Rs("T31_KAISI_BI")       '//�s���J�n�����i�[
            m_iGyojiCd  = w_Rs("T31_GYOJI_CD")       '//�s���R�[�h���i�[
            m_sGyojiMei = w_Rs("T31_GYOJI_MEI")     '//�s����
            m_sBiko     = w_Rs("T31_BIKO")              '//���l

                if f_ChkKaisibi = 0 Then            '//�J�n���̏ꍇ�\��
                    m_iCountN = m_iCountN + 1       '//�\������(N�Ԗ�)���J�E���g
                    Call show_Gyoji()
                else
                    if f_ChkHiduke = 0 Then                '//�y���\���`�F�b�N
                        if f_ChkHyojibi = 0 Then          '//�J�n���ȊO�ł��S�\���w��̏ꍇ�͕\��
                            m_iCountN = m_iCountN + 1      '//�\������(N�Ԗ�)���J�E���g
                            Call show_Gyoji()
                        else
                            'Call Show_NoGyoji()
                        end if
                    else
                        Exit Do
                    end if
                end if
                
            w_Rs.MoveNext
        Loop
        '// ����I��
        Exit Do
    
    Loop
    
    gf_closeObject(w_Rs)

'// LABEL_f_GetGyojiMei_END
End Function

'********************************************************************************
'*  [�@�\]  ���tCD�`�F�b�N�i�y���j���\�����g�p�j
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾�����A1:ں��ނȂ��A99:���s
'*  [����]  
'********************************************************************************
Function f_ChkHiduke()

    Dim w_Rs2                '// ں��޾�ĵ�޼ު��
    Dim w_iRet2              '// �߂�l
    Dim w_sSQL2              '// SQL��
    
    On Error Resume Next
    Err.Clear
    
    f_ChkHiduke = 0

    Do
    
        '// �s������ں��޾�Ă��擾
        w_sSQL2 = ""
        w_sSQL2 = w_sSQL2 & "SELECT"
        w_sSQL2 = w_sSQL2 & " T32_GYOJI_CD"
        w_sSQL2 = w_sSQL2 & " FROM T32_GYOJI_M "
        w_sSQL2 = w_sSQL2 & " WHERE T32_NENDO = " & m_iSyoriNen
        w_sSQL2 = w_sSQL2 & " AND T32_HIDUKE = '" & m_iDate & "'"
        w_sSQL2 = w_sSQL2 & " AND T32_GYOJI_CD = " & m_iGyojiCd
        
        w_iRet2 = gf_GetRecordset(w_Rs2, w_sSQL2)
'response.write w_sSQL2 & "<br>"
        
        If w_iRet2 <> 0 Then
            'ں��޾�Ă̎擾���s
            'm_sErrMsg = "ں��޾�Ă̎擾�Ɏ��s���܂���"
            f_ChkHiduke = 99
            Exit Do 'GOTO LABEL_f_ChkHiduke_END
        Else
        End If
        
        If w_Rs2.EOF Then
            '�Ώ�ں��ނȂ�
            f_ChkHiduke = 1
            Exit Do 'GOTO LABEL_f_ChkHiduke_END
        End If

        '// ����I��
        Exit Do
    
    Loop
    
    gf_closeObject(w_Rs2)

'// LABEL_f_ChkHiduke_END

End Function

'********************************************************************************
'*  [�@�\]  �J�n���`�F�b�N
'*  [����]  �Ȃ�
'*  [�ߒl]  0:�J�n���A1:�J�n���ȊO
'*  [����]  
'********************************************************************************
Function f_ChkKaisibi()

    f_ChkKaisibi = 1
    
        if m_iDate = m_iKaisibi Then
            f_ChkKaisibi = 0
        end if

End Function

'********************************************************************************
'*  [�@�\]  �\�����`�F�b�N
'*  [����]  �Ȃ�
'*  [�ߒl]  0:�\�����A1:�\�����ȊO
'*  [����]  
'********************************************************************************
Function f_ChkHyojibi()

    f_ChkHyojibi = 1

    if m_iHyojiFlg = 0 Then
        f_ChkHyojibi = 0
    elseif m_iDay = 1 Then
        f_ChkHyojibi = 0
    end if

End Function

'********************************************************************************
'*  [�@�\]  DB����l���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParamD()

    m_sYobi = ""
    m_iKyujituFlg = ""
    m_iYobiCd = ""
    m_iDay   = ""
    m_iMonth = ""
    m_iYear  = ""
    m_iDate  = ""

    m_sYobi = left(WeekdayName(Weekday(CDate(m_Rs("T32_HIDUKE")))) ,1)
    'm_iKyujituFlg = m_Rs("M40_KYUJITU_FLG")
    m_iYobiCd = Weekday(m_Rs("T32_HIDUKE"))	'//�j��CD
    m_iDay = day(m_Rs("T32_HIDUKE"))		'//��
    m_iMonth = month(m_Rs("T32_HIDUKE"))	'//��
    m_iYear = year(m_Rs("T32_HIDUKE"))		'//�N
    m_iDate = m_Rs("T32_HIDUKE")			'//���t

End Sub

'********************************************************************************
'*  [�@�\]  �e�[�u���̔w�i�F��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetColor()

	'//���t���j�����x�ɂ��ǂ���
	w_bHoliday = f_GetdateInfo()

	'//�j��,�x��
	If w_bHoliday = True Then
        m_sColor = "Holiday"
	Else
		'//�j���łȂ��ꍇ

        'If m_iYobiCd = "1" Then
        If m_iYobiCd = vbSunday Then
			'//���j��
            m_sColor = "Holiday"

        'ElseIf  m_iYobiCd = "7" Then
        ElseIf  m_iYobiCd = vbSaturday Then
			'//�y�j��
            m_sColor = "Saturday"
		Else
			'//����
	        m_sColor = "Weekday"
        End If

	End If
    
End Sub

'********************************************************************************
'*  [�@�\]  ���t���j�����ǂ����𒲂ׂ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetDateInfo()
    Dim w_Rs                '// ں��޾�ĵ�޼ު��
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_bHoliday

    On Error Resume Next
    Err.Clear

	w_bHoliday = False

    Do

        '// �s������ں��޾��(�x���f�[�^)���擾
        w_sSQL = ""
        w_sSQL = w_sSQL & "SELECT"
        w_sSQL = w_sSQL & " T32_GYOJI_CD"
        w_sSQL = w_sSQL & " FROM T32_GYOJI_M "
        w_sSQL = w_sSQL & " WHERE T32_NENDO = " & m_iSyoriNen
        w_sSQL = w_sSQL & " AND T32_HIDUKE = '" & m_iDate & "'"
        w_sSQL = w_sSQL & " AND T32_KYUJITU_FLG = '" & C_SYUKUJITU & "'"	'//T32_KYUJITU_FLG = C_SYUKUJITU �c�x��

'response.write w_sSQL & "<br>"

        w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            m_sErrMsg = ""
            Exit Do
        End If

        If w_Rs.EOF Then
            '//�x���ł͂Ȃ�
			w_bHoliday = False
		Else
			'//�x��
			w_bHoliday =True
        End If

        Exit Do
    Loop

	'//�߂�l���
	f_GetDateInfo = w_bHoliday

	'//ں��޾��CLOSE
    gf_closeObject(w_Rs)

End Function 

Sub show_Gyoji()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �s���\�肪1������ꍇ�̏o��
'********************************************************************************
if m_iCount = 1 Then
%>
    <TR>
        <TD ALIGN="right" class="<%=m_sColor%>"><%=m_iDay%><BR></TD>
        <TD ALIGN="center" class="<%=m_sColor%>"><%=m_sYobi%><BR></TD>
        <TD ALIGN="left" class="<%=m_sColor%>"><%=m_sGyojiMei%><BR></TD>
        <TD ALIGN="left" class="<%=m_sColor%>">
<%

        '//4���n�\��
        'if CInt(m_iMonth) = CInt(C_NENDO_KAISITUKI) and CInt(m_iDay) < CInt(m_iNKaisibi) Then
'        if CInt(m_iMonth) = CInt(C_NENDO_KAISITUKI) and CInt(m_iDay) <= CInt(day(m_sKouki_End)) Then
'            response.write m_sBiko & "��" & m_iYear & "�N" & chr(13)
'        else
            response.write m_sBiko & chr(13)
'        end if

%><BR>
        </TD>
    </TR>
<%
else
    if m_iCountN = 1 Then
        Call show_GyojiS()
    else
        Call show_GyojiSTd()
    end if
end if

End Sub

Sub show_GyojiS()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �s���\�肪��������ꍇ�̏o��
'********************************************************************************
%>
    <TR>
        <TD ALIGN="right" class="<%=m_sColor%>" rowspan="<%=m_iCount%>"><%=m_iDay%><BR></TD>
        <TD ALIGN="center" class="<%=m_sColor%>" rowspan="<%=m_iCount%>"><%=m_sYobi%><BR></TD>
        <TD ALIGN="left" class="<%=m_sColor%>"><%=m_sGyojiMei%><BR></TD>
        <TD ALIGN="left" class="<%=m_sColor%>">
<%
        '//4���n�\��
        'if CInt(m_iMonth) = CInt(C_NENDO_KAISITUKI) and CInt(m_iDay) < CInt(m_iNKaisibi) Then
'        if CInt(m_iMonth) = CInt(C_NENDO_KAISITUKI) and CInt(m_iDay) <= CInt(day(m_sKouki_End)) Then
'            response.write m_sBiko & "��" & m_iYear & "�N" & chr(13)
'        else
            response.write m_sBiko & chr(13)
'        end if
%><BR>
        </TD>
    </TR>
<%
End Sub

Sub show_GyojiSTd()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �s���\�肪��������ꍇ�̏o��
'********************************************************************************
%>
    <TR>
        <TD ALIGN="left" class="<%=m_sColor%>"><%=m_sGyojiMei%><BR></TD>
        <TD ALIGN="left" class="<%=m_sColor%>">
<%
        '//4���n�\��
        'if CInt(m_iMonth) = CInt(C_NENDO_KAISITUKI) and CInt(m_iDay) < CInt(m_iNKaisibi) Then
'        if CInt(m_iMonth) = CInt(C_NENDO_KAISITUKI) and CInt(m_iDay) <= CInt(day(m_sKouki_End)) Then
'            response.write m_sBiko & "��" & m_iYear & "�N" & chr(13)
'        else
            response.write m_sBiko & chr(13)
'        end if
%><BR>
        </TD>
    </TR>
<%
End Sub

Sub show_NoGyoji()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  �s���\�肪�����ꍇ�̏o��
'********************************************************************************
%>
    <TR>
        <TD ALIGN="right" class="<%=m_sColor%>"><%=m_iDay%><BR></TD>
        <TD ALIGN="center" class="<%=m_sColor%>"><%=m_sYobi%><BR></TD>
        <TD ALIGN="left" class="<%=m_sColor%>">�@<BR></TD>
        <TD ALIGN="left" class="<%=m_sColor%>">
<%
        '//4���n�\��
        'if CInt(m_iMonth) = CInt(C_NENDO_KAISITUKI) and CInt(m_iDay) < CInt(m_iNKaisibi) Then
'        if CInt(m_iMonth) = CInt(C_NENDO_KAISITUKI) and CInt(m_iDay) <= CInt(day(m_sKouki_End)) Then
'            response.write "��" & m_iYear & "�N" & chr(13)
'        else
            response.write "�@" & chr(13)
'        end if
%><BR>
        </TD>
    </TR>
<%
End Sub

Sub showPage_NoData()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>

    <html>
    <head>
 <link rel=stylesheet href="../../common/style.css" type=text/css>
   </head>

    <body>

    <center>
		<br><br><br>
		<span class="msg">�Ώۃf�[�^�͑��݂��܂���B��������͂��Ȃ����Č������Ă��������B</span>
    </center>

    </body>

    </html>


<%
    '---------- HTML END   ----------
End Sub

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
    Dim w_bFlg              '// �ް��L��
    Dim w_bNxt              '// NEXT�\���L��
    Dim w_bBfr              '// BEFORE�\���L��
    Dim w_iNxt              '// NEXT�\���Ő�
    Dim w_iBfr              '// BEFORE�\���Ő�
    Dim w_iCnt              '// �ް��\������

    Dim w_iRecordCnt        '//���R�[�h�Z�b�g�J�E���g

    On Error Resume Next
    Err.Clear

    w_iCnt  = 1
    w_bFlg  = True

%>

<html>
<head>
<link rel=stylesheet href="../../common/style.css" type=text/css>
<!--#include file="../../Common/jsCommon.htm"-->
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    //************************************************************
    //  [�@�\]  �߂�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //  [�쐬��] 
    //************************************************************
    function f_BackClick(){
    
        document.frm.action="../../menu/sansyo.asp";
        document.frm.target="_parent"
        document.frm.submit();
        
        
    }
//-->
</SCRIPT>
</head>

<body>
<center>
<form name="frm" Method="POST">
    <input type="hidden" name="txtMode" value="<%=m_sMode%>">
    <input type="hidden" name="chkGyojiCd" value="<%=m_iGyojiFlg%>">
<table border=0 width="<%=C_TABLE_WIDTH%>">
<tr><td align="center">
    <table border="1" width="100%" CLASS="hyo">
        <colgroup width="10%" valign="top">
        <colgroup width="10%" valign="top">
        <colgroup width="40%" valign="top">
        <colgroup width="40%" valign="top">
        <TR>
            <TH CLASS="header">��</TH>
            <TH CLASS="header">�j��</TH>
            <TH CLASS="header">�s����</TH>
            <TH CLASS="header">���l</TH>
        </TR>

<%Do Until m_Rs.EOF%>

	<%
	'//�O���[�o���ϐ��Ɋi�[
	Call s_SetParamD()

	'//�e�[�u���w�i�F�ݒ�
	Call s_SetColor()

	'//�s���̂ݕ\�����I�����ꂽ�ꍇ
    if m_iGyojiFlg = 1 Then
        Call f_GetGyojiMei()
    else
        if f_GetGyojiMei() = 1 Then
            Call show_NoGyoji()
        end if
    end if

	m_Rs.MoveNext

    If m_Rs.EOF Then
        w_bFlg = False
    ElseIf w_iCnt >= m_iDsp Then
        w_iNxt = m_iPageT + 1
        w_bNxt = True
        w_bFlg = False
    Else
        w_iCnt = w_iCnt + 1
    End If
    if m_Rs.EOF Then
        Exit Do
    end if

Loop
%>
    </table>
</td>
</tr>
</table>

</form>
</center>
</body>

</html>
<%
    '---------- HTML END   ----------
End Sub

%>
