<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Əo������
' ��۸���ID : kks/kks0110/kks0110_main.asp
' �@      �\: ���y�[�W ���Əo�����͂̈ꗗ���X�g�\�����s��
'-------------------------------------------------------------------------
' ��      ��: NENDO          '//�����N
'             KYOKAN_CD      '//����CD
'             GAKUNEN        '//�w�N
'             CLASSNO        '//�׽No
'             TUKI           '//��
' ��      ��:
' ��      �n: NENDO          '//�����N
'             KYOKAN_CD      '//����CD
'             GAKUNEN        '//�w�N
'             CLASSNO        '//�׽No
'             TUKI           '//��
' ��      ��:
'           �������\��
'               ���������ɂ��Ȃ��s���o�����͂�\��
'           ���o�^�{�^���N���b�N��
'               ���͏���o�^����
'-------------------------------------------------------------------------
' ��      ��: 2001/07/02 �ɓ����q
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ��CONST /////////////////////////////
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_bDaigae            '��֗��w���擾�׸�

    '�擾�����f�[�^�����ϐ�
    Public m_iSyoriNen      '//�����N�x
    Public m_iKyokanCd      '//����CD
    Public m_sGakunen       '//�w�N
    Public m_sClassNo       '//�׽NO
    Public m_sTuki          '//��
    Public m_sZenki_Start   '//�O���J�n��
    Public m_sKouki_Start   '//����J�n��
    Public m_sKouki_End     '//����I����

    Public m_sGakki         '//�w��
    Public m_sGakki_Kbn     '//�w���敪
    Public m_sKamokuCd      '//�ۖ�CD
    Public m_sSyubetu       '//���Ǝ��(TUJO:�ʏ����,TOKU:���ʊ���,KBTU:�ʎ���)
    Public m_sHissenKbn     '//�K�I�敪
	Public m_iTani			'//�P�����̒P�ʐ�
	Public m_bEndFLG		'//���ׂēo�^�s�̏ꍇTRUE

    Public m_AryHead()      '//�w�b�_���i�[�z��
    Public m_iRsCnt         '//�w�b�_ں��ސ�
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
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="���Əo������"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

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

        '//�ϐ�������
        Call s_ClearParam()

        '// ���Ұ�SET
        Call s_SetParam()

'//�f�o�b�O
'Call s_DebugPrint()

        '// �w�b�_���X�g���擾
        w_iRet = f_Get_HeadData()
        If w_iRet <> 0 Then
            m_bErrFlg = True
            Exit Do
        End If

		'//���ƃf�[�^���Ȃ��ꍇ
        'If m_iRsCnt < 0 Then
        If trim(request("txtMsg")) <> "" Then
			'//�󔒃y�[�W�\��
			Call showWhitePage(trim(request("txtMsg")))
            Exit Do
		Else
	        '// �y�[�W��\��
	        Call showPage()
        End If

        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
    '// �I������
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [�@�\]  �ϐ�������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen = ""
    m_iKyokanCd = ""
    m_sGakunen  = ""
    m_sClassNo  = ""
    m_sTuki     = ""
    m_sGakki    = ""
    m_sKamokuCd = ""
    m_sSyubetu  = ""
	m_iTani		= ""
	m_bEndFLG	= true

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_sZenki_Start = trim(Request("Tuki_Zenki_Start"))
    m_sKouki_Start = trim(Request("Tuki_Kouki_Start"))
    m_sKouki_End   = trim(Request("Tuki_Kouki_End"))
	m_iTani = Session("JIKAN_TANI") '�P�����̒P�ʐ�

    m_iSyoriNen = trim(Request("NENDO"))
    m_iKyokanCd = trim(Request("KYOKAN_CD"))

    m_sTuki     = trim(Request("TUKI"))
    m_sGakki    = trim(Request("GAKKI"))

    m_sSyubetu  = trim(Request("SYUBETU"))
    m_sGakunen  = trim(Request("GAKUNEN"))
    m_sClassNo  = trim(Request("CLASSNO"))
    m_sKamokuCd = trim(Request("KAMOKU_CD"))

	m_bEndFLG	= Cbool(Request("EndFLG"))

    If m_sGakki = "ZENKI" Then
        m_sGakki_Kbn = cstr(C_GAKKI_ZENKI)
    Else
        m_sGakki_Kbn = cstr(C_GAKKI_KOUKI)
    End If

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
Exit Sub
    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    response.write "m_iKyokanCd = " & m_iKyokanCd & "<br>"
    response.write "m_sGakunen  = " & m_sGakunen  & "<br>"
    response.write "m_sClassNo  = " & m_sClassNo  & "<br>"
    response.write "m_sTuki     = " & m_sTuki     & "<br>"
    response.write "m_sGakki    = " & m_sGakki    & "<br>"
    response.write "m_sKamokuCd = " & m_sKamokuCd & "<br>"
    response.write "m_sSyubetu  = " & m_sSyubetu  & "<br>"
    response.write "m_sZenki_Start = " & m_sZenki_Start & "<br>"
    response.write "m_sKouki_Start = " & m_sKouki_Start & "<br>"
    response.write "m_sKouki_End   = " & m_sKouki_End   & "<br>"

End Sub

'********************************************************************************
'*  [�@�\]  ���t�E�j���E���Ԃ̃w�b�_���擾�������s��
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_Get_HeadData()

    Dim w_sSQL
    Dim w_Rs

    On Error Resume Next
    Err.Clear
    
    f_Get_HeadData = 1

    Do 

        '//���t�͈̔͂��Z�b�g
        Call f_GetTukiRange(w_sSDate,w_sEDate)
		
		'// ���Ɠ��t�A���ԃf�[�^

		'// ���Ǝ�ʂ��l���ƁiKBTU�j�̎��͑�֎��Ԋ�����擾����B
		'// 2001/12/18 add
		If m_sSyubetu <> "KBTU" Then 

        w_sSQL = ""
		'// �ʏ�A���ʎ��Ƃ̏ꍇ
        w_sSQL = w_sSQL & vbCrLf & " SELECT"
        w_sSQL = w_sSQL & vbCrLf & "  A.T32_HIDUKE,"
        w_sSQL = w_sSQL & vbCrLf & "  B.T20_JIGEN AS JIGEN,"
        w_sSQL = w_sSQL & vbCrLf & "  B.T20_YOUBI_CD AS YOUBI_CD"
        w_sSQL = w_sSQL & vbCrLf & " FROM"
        w_sSQL = w_sSQL & vbCrLf & " T32_GYOJI_M A"
        w_sSQL = w_sSQL & vbCrLf & " ,T20_JIKANWARI B"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & " B.T20_YOUBI_CD = A.T32_YOUBI_CD "

		'//���ʊ����̏ꍇ�́A�������̎��Ƃ̍s���������āA�s�����ǂ����𔻒f����
		If m_sSyubetu = "TOKU"Then
	        w_sSQL = w_sSQL & vbCrLf & " AND TRUNC(B.T20_JIGEN+0.5) = A.T32_JIGEN"
		Else
	        w_sSQL = w_sSQL & vbCrLf & " AND B.T20_JIGEN = A.T32_JIGEN"
		End If

        w_sSQL = w_sSQL & vbCrLf & " AND B.T20_NENDO = A.T32_NENDO"
        w_sSQL = w_sSQL & vbCrLf & " AND A.T32_HIDUKE>='" & w_sSDate & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND A.T32_HIDUKE<'"  & w_sEDate & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND B.T20_NENDO="      & cInt(m_iSyoriNen)
        w_sSQL = w_sSQL & vbCrLf & " AND B.T20_GAKKI_KBN='" & m_sGakki_Kbn & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND B.T20_GAKUNEN= "   & cInt(m_sGakunen)
        w_sSQL = w_sSQL & vbCrLf & " AND B.T20_CLASS= "     & cInt(m_sClassNo)
        w_sSQL = w_sSQL & vbCrLf & " AND B.T20_KAMOKU='"    & trim(m_sKamokuCd) & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND B.T20_KYOKAN='"    & m_iKyokanCd & "'"
        w_sSQL = w_sSQL & vbCrLf & " AND A.T32_GYOJI_CD=0"
        w_sSQL = w_sSQL & vbCrLf & " AND A.T32_KYUJITU_FLG='0' "
        w_sSQL = w_sSQL & vbCrLf & " GROUP BY A.T32_HIDUKE,B.T20_YOUBI_CD,B.T20_JIGEN "
        w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T32_HIDUKE,B.T20_JIGEN"
		Else
			'// ���w���̑�։Ȗڂ̏ꍇ
			w_sSQL = ""
			w_sSQL = w_sSQL & vbCrLf & " SELECT"
			w_sSQL = w_sSQL & vbCrLf & "  A.T32_HIDUKE,"
			w_sSQL = w_sSQL & vbCrLf & "  B.T23_JIGEN AS JIGEN,"
			w_sSQL = w_sSQL & vbCrLf & "  B.T23_YOUBI_CD AS YOUBI_CD"
			w_sSQL = w_sSQL & vbCrLf & " FROM"
			w_sSQL = w_sSQL & vbCrLf & " T32_GYOJI_M A"
			w_sSQL = w_sSQL & vbCrLf & " ,T23_DAIGAE_JIKAN B"
			w_sSQL = w_sSQL & vbCrLf & " WHERE "
			w_sSQL = w_sSQL & vbCrLf & " B.T23_YOUBI_CD = A.T32_YOUBI_CD "
	'		 w_sSQL = w_sSQL & vbCrLf & " AND B.T23_JIGEN = A.T32_JIGEN"
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_NENDO = A.T32_NENDO"
			w_sSQL = w_sSQL & vbCrLf & " AND A.T32_HIDUKE>='" & w_sSDate & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND A.T32_HIDUKE<'"  & w_sEDate & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_NENDO="		& cInt(m_iSyoriNen)
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_GAKKI_KBN=" & m_sGakki_Kbn & " "
'			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_GAKUNEN= "	& cInt(m_sGakunen)
'			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_CLASS= " 	& cInt(m_sClassNo)
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_KAMOKU='"	& trim(m_sKamokuCd) & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_KYOKAN='"	& m_iKyokanCd & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND A.T32_GYOJI_CD=0"
			w_sSQL = w_sSQL & vbCrLf & " AND A.T32_KYUJITU_FLG='0' "
			w_sSQL = w_sSQL & vbCrLf & " GROUP BY A.T32_HIDUKE,B.T23_YOUBI_CD,B.T23_JIGEN "
			w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T32_HIDUKE,B.T23_JIGEN"
		End If

'response.write w_sSQL & "<BR>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_Get_HeadData = 99
            Exit Do
        End If

        m_iRsCnt = 0

        '=======================
        '//���Ԋ���z��ɃZ�b�g
        '=======================
        If w_Rs.EOF = false Then

            i = 0
            w_sHi = ""
            w_Rs.MoveFirst
            Do Until w_Rs.EOF
				
                '//�擾�������t�̎������x���܂��́A�s���̏ꍇ(w_bGyoji=True)�͂͂���
                iRet = f_Get_DateInfo(w_Rs("T32_HIDUKE"),w_Rs("JIGEN"),w_bGyoji)
                If iRet <> 0 Then
                    msMsg = Err.description
                    f_Get_HeadData = 99
                    Exit Do
                End If

                '//�x���E�s���ȊO�̂݃f�[�^���Z�b�g
                If w_bGyoji <> True Then

                    '//�z���ݒ�
                    ReDim Preserve m_AryHead(4,i)

                    '//�f�[�^�i�[
                    If w_sHi = gf_SetNull2String(w_Rs("T32_HIDUKE")) Then
                        m_AryHead(0,i) = ""     '//��
                        m_AryHead(1,i) = ""     '//��
                        m_AryHead(2,i) = ""     '//�j��CD
                    Else
                        m_AryHead(0,i) = month(gf_SetNull2String(w_Rs("T32_HIDUKE")))     '//��
                        m_AryHead(1,i) = day(gf_SetNull2String(w_Rs("T32_HIDUKE")))       '//��
                        m_AryHead(2,i) = gf_SetNull2String(w_Rs("YOUBI_CD"))          '//�j��CD
                    End If

                    m_AryHead(3,i) = gf_SetNull2String(w_Rs("JIGEN"))    '//����
                    m_AryHead(4,i) = gf_SetNull2String(w_Rs("T32_HIDUKE"))   '//���t

                    w_sHi = gf_SetNull2String(w_Rs("T32_HIDUKE"))
                    i = i + 1

                End If

                w_Rs.MoveNext
            Loop

        End If

        '//�擾�����f�[�^�����Z�b�g
        m_iRsCnt = i-1

        '//����I��
        f_Get_HeadData = 0
        Exit Do
    Loop

    '//ں��޾��CLOSE
   Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �擾�������t�E�������A�x���܂��͍s���łȂ���
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  
'********************************************************************************
Function f_Get_DateInfo(p_Hiduke,p_Jigen,p_bGyoji)

    Dim w_sSQL
    Dim w_Rs
    Dim w_bGyoujiFlg

    On Error Resume Next
    Err.Clear
    
    f_Get_DateInfo = 1
    w_bGyojiFlg = False

    Do 

        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT"
        w_sSQL = w_sSQL & vbCrLf & " A.T32_GYOJI_CD"
        w_sSQL = w_sSQL & vbCrLf & " FROM T32_GYOJI_M A"
        w_sSQL = w_sSQL & vbCrLf & " WHERE "
        w_sSQL = w_sSQL & vbCrLf & "  A.T32_NENDO=2001 "
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T32_GAKUNEN IN (" & cInt(m_sGakunen) & "," & C_GAKUNEN_ALL & ")"
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T32_CLASS IN ("   & cInt(m_sClassNo) & "," & C_CLASS_ALL   & ")"
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T32_HIDUKE='" & p_Hiduke & "'"
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T32_JIGEN=" & p_Jigen
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T32_COUNT_KBN<>" & C_COUNT_KBN_JUGYO
        w_sSQL = w_sSQL & vbCrLf & "  AND A.T32_KYUJITU_FLG<>'" & C_HEIJITU & "'"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_Get_DateInfo = 99
            Exit Do
        End If

        If w_Rs.EOF = False Then
            '//ں��ނ�����ꍇ�͋x�����A�s���̓�
            w_bGyojiFlg = True
        End If

        f_Get_DateInfo = 0
        Exit Do
    Loop

        '//�߂�l���Z�b�g
        p_bGyoji = w_bGyojiFlg

        '//ں��޾��CLOSE
       Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  ���̌����������쐬(7���c�@"MONTH>=2001/07/01 AND MONTH<2001/08/01" �Ƃ��Ďg�p)
'*  [����]  �Ȃ�
'*  [�ߒl]  p_sSDate
'*          p_sEDate
'*  [����]  
'********************************************************************************
Function f_GetTukiRange(p_sSDate,p_sEDate)

    p_sSDate = ""
    p_sEDate = ""

    If m_sGakki = "ZENKI" Then
        w_iNen = cint(m_iSyoriNen)

	    '//�J�n��
		If cint(month(m_sZenki_Start)) = Cint(m_sTuki) Then
			p_sSDate = m_sZenki_Start
		Else
			p_sSDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_sTuki),2) & "/01"
		End If

	    '//�I����
		If cint(month(m_sKouki_Start)) = Cint(m_sTuki) Then
			p_sEDate = m_sKouki_Start
		Else 
		    If Cint(m_sTuki) = 12 Then
		        p_sEDate = cstr(w_iNen+1) & "/01/01"
		    Else
		        p_sEDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_sTuki+1),2) & "/01"
		    End If
		End If

    Else
		'//����̔N
        If cint(m_sTuki) <=4 Then
            w_iNen = cint(m_iSyoriNen) + 1
        Else
            w_iNen = cint(m_iSyoriNen)
        End If

	    '//�J�n��
		If cint(month(m_sKouki_Start)) = Cint(m_sTuki) Then
		    p_sSDate = m_sKouki_Start
		Else
		    p_sSDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_sTuki),2) & "/01"
		End If

	    '//�I����
		If cint(month(m_sKouki_End)) = Cint(m_sTuki) Then
			'p_sEDate = m_sKouki_End
			p_sEDate = DateAdd("d",1,m_sKouki_End)
		Else 
		    If Cint(m_sTuki) = 12 Then
		        p_sEDate = cstr(w_iNen+1) & "/01/01"
		    Else
		        p_sEDate = cstr(w_iNen) & "/" & gf_fmtZero(cstr(m_sTuki+1),2) & "/01"
		    End If
		End If

    End If

End Function

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
	dim w_str	'�\�����b�Z�[�W

    On Error Resume Next
    Err.Clear

	if m_bEndFLG = False then '�܂��o�^�\
		If m_iTani > 1 then '�P�����̒P�ʐ����Q�ȏ�̎�
			w_str = "<span class=CAUTION>�� �����̉��̋󔒗����N���b�N���āA�o���󋵂���͂��Ă��������B�i�����x�������P����(�o��)�̏��ŕ\������܂��j<BR></span>" & vbCrLf
			w_str = w_str & "�y ���F����(" & m_iTani & "���ە�)�@�P�F����(�P���ە�)�@�x�F�x���@���F���� �z" & vbCrLf
		Else				'�m�[�}�����
			w_str = "<span class=CAUTION>�� �����̉��̋󔒗����N���b�N���āA�o���󋵂���͂��Ă��������B�i�����x��������(�o��)�̏��ŕ\������܂��j</span>" & VbCrLf
			w_str = w_str & "�y ���F���ہ@�x�F�x���@���F���� �z" & vbCrLf
		End If

	Else '�o�^���Ԃ��߂��ēo�^�s�\�̏ꍇ�́A�Q�Ƃ̂�
		If m_iTani > 1 then 
			w_str = "�y ���F����(" & m_iTani & "���ە�)�@�P�F����(�P���ە�)�@�x�F�x���@���F���� �z" & vbCrLf
		Else 
			w_str = "�y ���F���ہ@�x�F�x���@���F���� �z" & vbCrLf
		End If 
	End If

%>
    <html>
    <head>
    <title>���Əo������</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {

		//�X�N���[����������
		parent.init();

    }

    //************************************************************
    //  [�@�\]  �o�^�{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Touroku(){

		parent.frames["main"].f_Touroku();
		return;
    }

    //************************************************************
    //  [�@�\]  �L�����Z���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Cancel(){
        //�󔒃y�[�W��\��
        parent.document.location.href="default.asp"
    }

    //-->
    </SCRIPT>

    </head>
    <body LANGUAGE=javascript onload="return window_onload()">
    <form name="frm" method="post">
    <center>
    <%call gs_title("���Əo������","��@��")%>
    <%Do %>
        <%If m_iRsCnt < 0 Then%>
            <br><br>
            <span class="msg">���Ɠ���������܂���</span>
            <%Exit Do%>
        <%End If%>

<!-------------------------�ǉ����w�b�_----------------------------------->

        <table>
		<tr><td>
	        <table class="hyo" border="1" width="545">
	            <tr>

					<%
					'//�ʎ��Ƃ̏ꍇ�͔N�A�N���X���Ȃ�
					If trim(m_sSyubetu) = "KBTU" Then
						w_sClass_Name = "-"
					Else
						w_sClass_Name = Request("CLASS_NAME")
					End If

					If m_sGakki_Kbn = cstr(C_GAKKI_ZENKI) Then
						w_sGakki = "�O��"
					Else
						w_sGakki = "���"
					End If

					%>
	                <th nowrap class="header" width="65"  align="center"><%=w_sGakki%></th>
	                <td nowrap class="detail" width="50"  align="center"><%=m_sTuki%>��</td>
	                <th nowrap class="header" width="65"  align="center">�N���X</th>
	                <td nowrap class="detail" width="150"  align="center"><%=m_sGakunen%>�N�@<%=w_sClass_Name%></td>
	                <th nowrap class="header" width="65" align="center">����</th>
	                <td nowrap class="detail" width="150" align="center"><%=Request("KAMOKU_NAME")%></td>

	            </tr>
	        </table>
		</td></tr><tr>
		<td align="center">
			<table>
				<tr>
<%	if m_bEndFLG = False then '�܂��o�^�\ %>
		        <td valign="bottom"align="center">
		            <input class="button" type="button" onclick="javascript:f_Touroku();" value="�@�o�@�^�@">
		            &nbsp;&nbsp;&nbsp;
		            <input class="button" type="button" onclick="javascript:f_Cancel();" value="�L�����Z��">
		        </td>
<% Else %>
		        <td valign="bottom"align="center">
		            <input class="button" type="button" onclick="javascript:f_Cancel();" value=" �߁@�� ">
		        </td>
<% End If %>
				</tr>
	        </table>
		</td></tr>
        </table>

<!-------------------------�ǉ����w�b�_----------------------------------->

		<table>
        <tr>
<% '�\�����b�Z�[�W�@�p�^�[���ɂ��\�������� %>
            <td align="center" colspan=3><font size="2" color="#222268"><%=w_str%></font>
            </td>
        </tr>

		</table>

        <!--���׃w�b�_��(���E�j���E��������\��)-->
        <table >
        <tr>
            <td align="center" valign="top">
            <table class="hyo"  border="1" >

            <tr>
                <th class="header" height="100" rowspan="4" width="50" align="center"  nowrap><font >
                    <table ><tr><th width="10" class="header" nowrap><%=gf_GetGakuNomei(m_iSyoriNen,C_K_KOJIN_1NEN)%></th></tr></table></font>
                </th>
                <th class="header" width="150" align="center" nowrap><font color="#ffffff">��</font></th>
                <%for i = 0 to m_iRsCnt%>
                    <th class="header" align="center"  nowrap><font ><%=m_AryHead(0,i)%><br></font></th>

                <%Next%>
            </tr>

            <tr>
                <th class="header" align="center" nowrap ><font color="#ffffff">��</font></th>
                <%for i = 0 to m_iRsCnt%>
                    <th class="header" align="center" width="30"  nowrap ><font color="#ffffff"><%=m_AryHead(1,i)%></font></th>
                <%Next%>
            </tr>

            <tr>
                <th class="header" align="center" nowrap><font color="#ffffff">�j��</font></th>
                <%for i = 0 to m_iRsCnt%>
                    <th class="header" align="center" width="30"  nowrap ><font color="#ffffff"><%=gf_GetYoubi(m_AryHead(2,i))%></font></th>
                <%Next%>
            </tr>

            <tr>
                <th class="header" align="center" nowrap><font color="#ffffff">����</font></th>
                <%for i = 0 to m_iRsCnt%>
					<%
					'//���ƊO�����̏ꍇ�́A������\�����Ȃ�
					If m_sSyubetu = "TOKU"Then
						w_iDispJigen = "-"
					Else
						w_iDispJigen = m_AryHead(3,i)
					End If
					%>
                    <th class="header" align="center" width="30"  nowrap ><font color="#ffffff"><%=w_iDispJigen%></font>
                    </th>
                <%Next%>
            </tr>
            </table>

        </td>
        <td width="10"><br></td>
        <td align="center" valign="top" width="120" nowrap>

            <!--���E�w���̌��ȋy�ђx�����݌v-->
            <table width="120" class="hyo" border="1">
            <tr>
                <th height="20" colspan="2" class="header" nowrap align="center" width="60"><font color="#ffffff">���v</font></th>
                <th height="20" colspan="2" class="header" nowrap align="center" width="60"><font color="#ffffff">�v</font></th>
            </tr>
            <tr>
                <th height="80" class="header" width="30" align="center" nowrap><font color="#ffffff">�x<br>��</font></th>
                <th height="80" class="header" width="30" align="center" nowrap><font color="#ffffff">��<br>��</font></th>
                <th height="80" class="header" width="30" align="center" nowrap><font color="#ffffff">�x<br>��</font></th>
                <th height="80" class="header" width="30" align="center" nowrap><font color="#ffffff">��<br>��</font></th>
            </tr>
            </table>

        </td>
        </tr>
        </table>

        <%Exit Do%>
    <%Loop%>

    </form>
    </center>
    </body>
    </html>
<%
End Sub

'********************************************************************************
'*  [�@�\]  ��HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showWhitePage(p_Msg)
%>
    <html>
    <head>
    <title>���Əo������</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {
		parent.document.location.href="default2.asp?txtMsg=<%=Server.URLEncode(p_Msg)%>"
		return;
    }
    //-->
    </SCRIPT>

    </head>
	<body LANGUAGE=javascript onload="return window_onload()">

    </body>
    </html>
<%
End Sub
%>

