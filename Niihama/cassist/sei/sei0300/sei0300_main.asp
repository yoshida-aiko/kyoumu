<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �l�ʐ��шꗗ
' ��۸���ID : sei/sei0300/sei0300_top.asp
' �@      �\: �t���[���y�[�W ���шꗗ�̓o�^���s��
'-------------------------------------------------------------------------
' ��      ��:NENDO			:�N�x
'            txtSikenKBN	:�����敪
'            txtGakuNo		:�w�N
'            txtClassNo		:�N���XNO
'            txtGakusei		:�w��NO
'            txtBeforGakuNo	:�O�̊w��NO
'            txtAfterGakuNo	:��̊w��NO
' ��      �n:
'            txtSikenKBN	:�����敪
'            txtGakuNo		:�w�N
'            txtClassNo		:�N���XNO
'            txtGakusei		:�w��NO
'            txtBeforGakuNo	:�O�̊w��NO
'            txtAfterGakuNo	:��̊w��NO
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/09/04 �ɓ����q
' ��      �X: 2003/10/24 ���c�F��������̏ꍇ�A���Ǝ��Ԑ��E���ێ����E�x���񐔂�ݐς��ĕ\��
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<!--#include file="./sei0300_com.asp"-->

<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////
	Public CONST C_KETTEN_LIMIT = 60
	Public CONST C_KENGEN_SEI0300_FULL = "FULL"	'//�A�N�Z�X����FULL
	Public CONST C_KENGEN_SEI0300_TAN = "TAN"	'//�A�N�Z�X�����S�C
	Public CONST C_KENGEN_SEI0300_GAK = "GAK"	'//�A�N�Z�X�����w��

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
	Public  m_bErrFlg           '�װ�׸�
	Public  m_iSyoriNen		'//�����N�x

	'//���C���̕ϐ�
	Public  m_sSikenKBN		'//�����敪
	Public  m_sGakunen		'//�w�N
	Public  m_sClassNo		'//�N���X
	Public  m_sGakkaNo		'//�w��NO
	Public  m_sGakusei		'//�w��NO

	Public  m_sName			'//���k����
	Public  m_sGakusekiNo	'//�w��NO
	Public  m_sBeforGakuNo	'//�w��NO���O�̊w��
	Public  m_sAfterGakuNo	'//�w��NO�����Ƃ̊w��

	Public  m_iMemberCnt	'//�N���X�l��
	Public  m_iSikiji		'//�Ȏ�
	Public  m_iSyoken		'//����
	Public  m_iAverage		'//���ϓ_
	
	Public  m_AryResult()	'//���уf�[�^�i�[�z��
	Public  m_iCnt			'//���уf�[�^����
	Public  m_sKengen
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
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message�p�̕ϐ��̏�����
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="���шꗗ"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"     
	w_sTarget="_parent"

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

        '//�l�̏�����
        Call s_ClearParam()

        '//�ϐ��Z�b�g
        Call s_SetParam()

		'//�����`�F�b�N
		w_iRet = f_CheckKengen(w_sKengen)
		If w_iRet <> 0 Then
            m_bErrFlg = True
			m_sErrMsg = "�Q�ƌ���������܂���B"
			Exit Do
		End If

		'//�������S�C�̏ꍇ�͒S�C�N���X�����擾����
		If w_sKengen = C_KENGEN_SEI0300_TAN Then

			'//�S�C�N���X���擾
			'//��񂪎擾�ł��Ȃ��ꍇ�͒S�C�N���X�������ׁA�Q�ƕs�Ƃ���
			w_iRet = f_GetClassInfo(m_sKengen)
			If w_iRet <> 0 Then
				m_bErrFlg = True
				m_sErrMsg = "�Q�ƌ���������܂���B"
				Exit Do
			End If

		ElseIf w_sKengen = C_KENGEN_SEI0300_GAK Then

			'//�w�ȏ��擾
			'//��񂪎擾�ł��Ȃ��ꍇ�͊w�Ȃ������ׁA�Q�ƕs�Ƃ���
			w_iRet = f_GetGakkaInfo(m_sKengen)
			If w_iRet <> 0 Then
				m_bErrFlg = True
				m_sErrMsg = "�Q�ƌ���������܂���B"
				Exit Do
			End If

		End If

		'//���k���擾
		w_iRet = f_GetGakuseiData()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			m_sErrMsg = "���k��񂪎擾�ł��܂���ł����B"
			Exit Do
		End If

		'//�Ȏ��A�������擾
		w_iRet = f_GetGakuseiInfo()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			m_sErrMsg = "���k��񂪎擾�ł��܂���ł����B"
			Exit Do
		End If

		'//���уf�[�^�擾
		w_iRet = f_GetResultData()
		If w_iRet <> 0 Then
			m_bErrFlg = True

			Exit Do
		End If

		'// �y�[�W��\��
		Call showPage()
	    Exit Do
	Loop

	'// �װ�̏ꍇ�ʹװ�߰�ނ�\��
	If m_bErrFlg = True Then
		'w_sMsg = gf_GetErrMsg()
		'Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
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
	m_sSikenKBN = ""
    m_sGakunen  = ""
    m_sClassNo  = ""
    m_sGakkaNo  = ""
    m_sGakusei  = ""

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

	m_iSyoriNen = Session("NENDO")
	m_sSikenKBN = Request("txtSikenKBN")
	m_sGakunen  = Request("txtGakuNo")
	m_sClassNo  = Request("txtClassNo")
	m_sGakkaNo  = Request("txtGakkaNo")
	m_sGakusei  = Request("txtGakusei")

End Sub

'********************************************************************************
'*  [�@�\]  �f�o�b�O�p
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_DebugPrint()
'Exit Sub

    response.write "m_iSyoriNen = " & m_iSyoriNen & "<br>"
    response.write "m_sSikenKBN = " & m_sSikenKBN & "<br>"
    response.write "m_sGakunen  = " & m_sGakunen  & "<br>"
    response.write "m_sClassNo  = " & m_sClassNo  & "<br>"
    response.write "m_sGakkaNo  = " & m_sGakkaNo  & "<br>"
    response.write "m_sGakusei  = " & m_sGakusei  & "<br>"

End Sub

'********************************************************************************
'*	[�@�\]	�����`�F�b�N
'*	[����]	�Ȃ�
'*	[�ߒl]	w_sKengen
'*	[����]	���O�C��USER�̏������x���ɂ��A�Q�Ɖs�̔��f������
'*			�@FULL�A�N�Z�X�����ێ��҂́A�S�Ă̐��k�̐��я����Q�Ƃł���
'*			�A�S�C�A�N�Z�X�����ێ��҂́A�󂯎����N���X���k�̐��я����Q�Ƃł���
'*			�B��L�ȊO��USER�͎Q�ƌ����Ȃ�
'********************************************************************************
Function f_CheckKengen(p_sKengen)
    Dim w_iRet
    Dim w_sSQL
	 Dim rs

	 On Error Resume Next
	 Err.Clear

	 f_CheckKengen = 1

	 Do

		'T51��茠�����擾
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  T51_SYORI_LEVEL.T51_ID "
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  T51_SYORI_LEVEL"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T51_SYORI_LEVEL.T51_ID IN ('SEI0300','SEI0301','SEI0302')"
		w_sSql = w_sSql & vbCrLf & "  AND T51_SYORI_LEVEL.T51_LEVEL" & Session("LEVEL") & " = 1"

		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			m_sErrMsg = Err.description
			f_CheckKengen = 99
			Exit Do
		End If

		If rs.EOF Then
			m_sErrMsg = "�Q�ƌ���������܂���B"
			Exit Do
		Else
			Select Case rs("T51_ID")
				Case "SEI0300"	'//�t���A�N�Z�X��������
					p_sKengen = C_KENGEN_SEI0300_FULL
				Case "SEI0301"	'//�S�C�����L��
					p_sKengen = C_KENGEN_SEI0300_TAN
				Case "SEI0302"	'//�w�Ȍ����L��
					p_sKengen = C_KENGEN_SEI0300_GAK
			End Select

		End If

		f_CheckKengen = 0
		Exit Do
	 Loop


	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �����`�F�b�N�i�S�C�N���X���擾�j
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  ���S�C�A�N�Z�X�������ݒ肳��Ă���USER�ł��A���ۂɒS�C�N���X��
'*			�󂯎����Ă��Ȃ��ꍇ�ɂ͎Q�ƕs�Ƃ���
'********************************************************************************
Function f_GetClassInfo(p_sKengen)

	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClassInfo = 1
	p_sKengen = ""

	Do 

		'// �S�C�N���X���
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & "  ,M05_CLASS.M05_CLASSNO "
		w_sSQL = w_sSQL & vbCrLf & " FROM M05_CLASS"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      M05_CLASS.M05_NENDO=" & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_TANNIN='" & session("KYOKAN_CD") & "'"

		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_GetClassInfo = 99
			Exit Do
		End If

		If rs.EOF Then
			'//�N���X��񂪎擾�ł��Ȃ��Ƃ�
            m_sErrMsg = "�Q�ƌ���������܂���B"
			Exit Do
		End If

		f_GetClassInfo = 0
		p_sKengen = C_KENGEN_SEI0300_TAN
		Exit Do
	Loop

	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �����`�F�b�N�i���[�U�w�ȏ��擾�j
'*  [����]  �Ȃ�
'*  [�ߒl]  0:���擾���� 99:���s
'*  [����]  ���S�C�A�N�Z�X�������ݒ肳��Ă���USER�ł��A���ۂɒS�C�N���X��
'*			�󂯎����Ă��Ȃ��ꍇ�ɂ͎Q�ƕs�Ƃ���
'********************************************************************************
Function f_GetGakkaInfo(p_sKengen)

	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetGakkaInfo = 1
	p_sKengen = ""

	Do 

		'// �S�C�N���X���
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M04_GAKKA_CD "
		w_sSQL = w_sSQL & vbCrLf & " FROM M04_KYOKAN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      M04_NENDO=" & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & "  AND M04_KYOKAN_CD='" & session("KYOKAN_CD") & "'"
		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_GetGakkaInfo = 99
			Exit Do
		End If
		If rs.EOF Then
			'//�N���X��񂪎擾�ł��Ȃ��Ƃ�
            m_sErrMsg = "�Q�ƌ���������܂���B"
			Exit Do
		Else
			p_sKengen = C_KENGEN_SEI0300_GAK 
'			m_sGakkaNo  = rs("M04_GAKKA_CD")
'			m_sGakkaMei = rs("M02_GAKKAMEI")

			'//�������S�C�̏ꍇ�́A�S�C�N���X�ȊO�͑I���ł��Ȃ�
'			m_sGakuNoOption = " DISABLED "
'			m_sClassNoOption = " DISABLED "
		End If

		f_GetGakkaInfo = 0
		Exit Do
	Loop

	Call gf_closeObject(rs)

End Function

Function f_GetGakuseiData()
'********************************************************************************
'*  [�@�\]  ���k�f�[�^���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Dim i
i = 1

	f_GetGakuseiData = 1

	Do 

		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT  "
		w_sSQL = w_sSQL & vbCrLf & "     A.T11_GAKUSEI_NO "
		w_sSQL = w_sSQL & vbCrLf & "    ,A.T11_SIMEI "
		w_sSQL = w_sSQL & vbCrLf & "    ,B.T13_GAKUSEKI_NO "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "     T11_GAKUSEKI A,T13_GAKU_NEN B "
		w_sSQL = w_sSQL & vbCrLf & " WHERE"
		w_sSQL = w_sSQL & vbCrLf & "     B.T13_NENDO = " & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & " AND B.T13_GAKUNEN = " & m_sGakunen
	If m_sKengen <> C_KENGEN_SEI0300_GAK then
		w_sSQL = w_sSQL & vbCrLf & " AND B.T13_CLASS = " & m_sClassNo
	Else
		w_sSQL = w_sSQL & vbCrLf & " AND B.T13_Gakka_CD = " & m_sGakkaNo
	End If
		w_sSQL = w_sSQL & vbCrLf & " AND A.T11_GAKUSEI_NO = B.T13_GAKUSEI_NO "
'		w_sSQL = w_sSQL & vbCrLf & " AND A.T11_NYUNENDO = B.T13_NENDO - B.T13_GAKUNEN + 1 "
		'//���ݍ݊w���̐��k�̂ݕ\���ΏۂƂ���
		w_sSQL = w_sSQL & vbCrLf & " AND B.T13_ZAISEKI_KBN < " & C_ZAI_SOTUGYO
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY B.T13_GAKUSEKI_NO "

		w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If w_iRet <> 0 Then
	        'ں��޾�Ă̎擾���s
			f_GetGakuseiData = 99
			Exit do 
	    End If

		w_rCnt=cint(gf_GetRsCount(w_Rs))

		'//�z��̍쐬
		w_Rs.MoveFirst
		Do Until w_Rs.EOF

			ReDim Preserve w_sGakuseiAry(i)
			w_sGakuseiAry(i) = w_Rs("T11_GAKUSEI_NO")
			i = i + 1

			If w_Rs("T11_GAKUSEI_NO") = m_sGakusei Then
				'//���k���̂��擾���Z�b�g
				m_sName = w_Rs("T11_SIMEI")
				m_sGakusekiNo = w_Rs("T13_GAKUSEKI_NO")
			End If

			w_Rs.MoveNext

		Loop

		For i = 1 to w_rCnt

			If w_sGakuseiAry(i) = m_sGakusei Then

				'//�w��NO���O�̐��k�A��̐��k���擾���Z�b�g
				If i <= 1 Then
					m_sAfterGakuNo = w_sGakuseiAry(i+1)
					Exit For
				End If

				If i = w_rCnt Then
					m_sBeforGakuNo = w_sGakuseiAry(i-1)
					Exit For
				End If

				m_sAfterGakuNo = w_sGakuseiAry(i+1)
				m_sBeforGakuNo = w_sGakuseiAry(i-1)

				Exit For
			End If

		Next

		f_GetGakuseiData = 0
		Exit Do
	Loop

	'//ں��޾��CLOSE
	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �������̎擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetResultData()

    Dim w_iRet
    Dim w_sSQL
    Dim rs
	Dim w_bSelect
    dim m_SchoolNo

    On Error Resume Next
    Err.Clear

    f_GetResultData = 1

		'�w�Z�ԍ����擾�@2003.10.24 ins
		if Not gf_GetGakkoNO(m_SchoolNo) then
	        m_bErrFlg = True
			m_sErrMsg = "�w�Z�ԍ��̎擾�Ɏ��s���܂����B"
			Exit function
		end if
        m_SchoolNo = CSTR(m_SchoolNo)
'response.write m_SchoolNo
    Do

		'==============================
		'//�����̊J�n���A�I�������擾
		'==============================
		w_bRet = gf_GetKaisiSyuryo(cint(gf_SetNull2Zero(m_sSikenKBN)),cint(m_sGakunen), w_sKaisibi, w_sSyuryobi)

		If w_bRet <> True Then
			'//�J�n���A�I�����̎擾���s
		    f_GetResultData = 99
			m_sErrMsg = "���������̓o�^���s���Ă��������B"
			Exit Do 
		End If

		'==============================
		'//���я��擾
		'==============================
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  T16_KAMOKU_CD "
		w_sSql = w_sSql & vbCrLf & "  ,T16_KAMOKUMEI "
		w_sSql = w_sSql & vbCrLf & "  ,T16_HAITOTANI "
		w_sSql = w_sSql & vbCrLf & "  ,T16_KAISETU "
		w_sSql = w_sSql & vbCrLf & "  ,T16_HISSEN_KBN "
		w_sSql = w_sSql & vbCrLf & "  ,T16_SELECT_FLG "
		w_sSql = w_sSql & vbCrLf & "  ,T16_KAISETU "    '2003.10.28ins
		w_sSql = w_sSql & vbCrLf & "  ,M100_SEISEKI_INP "    '2003.10.28ins

		'//INS 2008/03/10 �������͂ɑΉ�
        if m_SchoolNo <> "11" then  '��������ȊO�̏ꍇ
		'//�����敪�ɂ��ꍇ�킯
		Select Case cint(gf_SetNull2Zero(m_sSikenKBN))

			Case C_SIKEN_ZEN_TYU    '�O�����Ԏ���
				w_sSql = w_sSql & vbCrLf & "  ,T16_SEI_TYUKAN_Z   AS HTEN "
				'w_sSql = w_sSql & vbCrLf & "  ,T16_HTEN_TYUKAN_Z   AS HTEN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_KEKA_TYUKAN_Z   AS KEKA "
				w_sSql = w_sSql & vbCrLf & "  ,T16_CHIKAI_TYUKAN_Z AS CHIKAI "
				w_sSql = w_sSql & vbCrLf & "  ,T16_J_JUNJIKAN_TYUKAN_Z AS JIKAN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_HYOKA_TYUKAN_Z AS HYOKA "	'INS 2008/03/10
				
			Case C_SIKEN_ZEN_KIM    '�O����������
				w_sSql = w_sSql & vbCrLf & "  ,T16_SEI_KIMATU_Z   AS HTEN "
				'w_sSql = w_sSql & vbCrLf & "  ,T16_HTEN_KIMATU_Z   AS HTEN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_KEKA_KIMATU_Z   AS KEKA "
				w_sSql = w_sSql & vbCrLf & "  ,T16_CHIKAI_KIMATU_Z AS CHIKAI "
				w_sSql = w_sSql & vbCrLf & "  ,T16_J_JUNJIKAN_KIMATU_Z AS JIKAN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_HYOKA_KIMATU_Z AS HYOKA "	'INS 2008/03/10
				
			Case C_SIKEN_KOU_TYU    '������Ԏ���
				w_sSql = w_sSql & vbCrLf & "  ,T16_SEI_TYUKAN_K   AS HTEN "
				'w_sSql = w_sSql & vbCrLf & "  ,T16_HTEN_TYUKAN_K   AS HTEN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_KEKA_TYUKAN_K   AS KEKA "
				w_sSql = w_sSql & vbCrLf & "  ,T16_CHIKAI_TYUKAN_K AS CHIKAI "
				w_sSql = w_sSql & vbCrLf & "  ,T16_J_JUNJIKAN_TYUKAN_K AS JIKAN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_HYOKA_TYUKAN_K AS HYOKA "	'INS 2008/03/10
				
			Case C_SIKEN_KOU_KIM    '�����������
				w_sSql = w_sSql & vbCrLf & "  ,T16_SEI_KIMATU_K   AS HTEN "
				'w_sSql = w_sSql & vbCrLf & "  ,T16_HTEN_KIMATU_K   AS HTEN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_KEKA_KIMATU_K   AS KEKA "
				w_sSql = w_sSql & vbCrLf & "  ,T16_CHIKAI_KIMATU_K AS CHIKAI"
				w_sSql = w_sSql & vbCrLf & "  ,T16_J_JUNJIKAN_KIMATU_K AS JIKAN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_HYOKA_KIMATU_K AS HYOKA "	'INS 2008/03/10
			Case Else
				'//�V�X�e���G���[
	            m_sErrMsg = "������񂪂���܂���B"
				Exit Do
		End Select
        END IF
        '================================================================================
        if m_SchoolNo = "11" then  '��������̏ꍇ
		'//�����敪�ɂ��ꍇ�킯
		Select Case cint(gf_SetNull2Zero(m_sSikenKBN))

			Case C_SIKEN_ZEN_TYU    '�O�����Ԏ���
				w_sSql = w_sSql & vbCrLf & "  ,T16_SEI_TYUKAN_Z   AS HTEN "
				'w_sSql = w_sSql & vbCrLf & "  ,T16_HTEN_TYUKAN_Z   AS HTEN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_HYOKA_TYUKAN_Z AS HYOKA "	'INS 2008/03/10
				
			Case C_SIKEN_ZEN_KIM    '�O����������
				w_sSql = w_sSql & vbCrLf & "  ,T16_SEI_KIMATU_Z   AS HTEN "
				'w_sSql = w_sSql & vbCrLf & "  ,T16_HTEN_KIMATU_Z   AS HTEN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_HYOKA_KIMATU_Z AS HYOKA "	'INS 2008/03/10
				
			Case C_SIKEN_KOU_TYU    '������Ԏ���
				w_sSql = w_sSql & vbCrLf & "  ,T16_SEI_TYUKAN_K   AS HTEN "
				'w_sSql = w_sSql & vbCrLf & "  ,T16_HTEN_TYUKAN_K   AS HTEN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_HYOKA_TYUKAN_K AS HYOKA "	'INS 2008/03/10
				
			Case C_SIKEN_KOU_KIM    '�����������
				w_sSql = w_sSql & vbCrLf & "  ,T16_SEI_KIMATU_K   AS HTEN "
				'w_sSql = w_sSql & vbCrLf & "  ,T16_HTEN_KIMATU_K   AS HTEN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_HYOKA_KIMATU_K AS HYOKA "	'INS 2008/03/10

			Case Else
				'//�V�X�e���G���[
	            m_sErrMsg = "������񂪂���܂���B"
				Exit Do
		End Select
				w_sSql = w_sSql & vbCrLf & "  ,T16_KEKA_TYUKAN_Z   AS KEKA "
				w_sSql = w_sSql & vbCrLf & "  ,T16_CHIKAI_TYUKAN_Z AS CHIKAI "
				w_sSql = w_sSql & vbCrLf & "  ,T16_J_JUNJIKAN_TYUKAN_Z AS JIKAN "
				w_sSql = w_sSql & vbCrLf & "  ,T16_KEKA_KIMATU_Z   AS KEKA2 "
				w_sSql = w_sSql & vbCrLf & "  ,T16_CHIKAI_KIMATU_Z AS CHIKAI2 "
				w_sSql = w_sSql & vbCrLf & "  ,T16_J_JUNJIKAN_KIMATU_Z AS JIKAN2 "
				w_sSql = w_sSql & vbCrLf & "  ,T16_KEKA_TYUKAN_K   AS KEKA3 "
				w_sSql = w_sSql & vbCrLf & "  ,T16_CHIKAI_TYUKAN_K AS CHIKAI3 "
				w_sSql = w_sSql & vbCrLf & "  ,T16_J_JUNJIKAN_TYUKAN_K AS JIKAN3 "
				w_sSql = w_sSql & vbCrLf & "  ,T16_KEKA_KIMATU_K   AS KEKA4 "
				w_sSql = w_sSql & vbCrLf & "  ,T16_CHIKAI_KIMATU_K AS CHIKAI4 "
				w_sSql = w_sSql & vbCrLf & "  ,T16_J_JUNJIKAN_KIMATU_K AS JIKAN4 "
        END IF
        '================================================================================
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  T16_RISYU_KOJIN"

'INS 2008/03/10
		w_sSql = w_sSql & vbCrLf & "  ,M03_KAMOKU"
		w_sSql = w_sSql & vbCrLf & "  ,M100_KAMOKU_ZOKUSEI"
'INS END 2008/03/10

		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  T16_NENDO=" & m_iSyoriNen
		w_sSql = w_sSql & vbCrLf & "  AND T16_GAKUSEI_NO='" & m_sGakusei & "'"

'INS 2008/03/10
		w_sSql = w_sSql & vbCrLf & "  AND T16_NENDO = M03_NENDO "
		w_sSql = w_sSql & vbCrLf & "  AND T16_KAMOKU_CD = M03_KAMOKU_CD "
		w_sSql = w_sSql & vbCrLf & "  AND M03_NENDO = M100_NENDO "
		w_sSql = w_sSql & vbCrLf & "  AND M03_ZOKUSEI_CD = M100_ZOKUSEI_CD "
		w_sSql = w_sSql & vbCrLf & "  AND M100_KAMOKUBUNRUI = '01' "
'INS END 2008/03/10

		w_sSql = w_sSql & vbCrLf & " ORDER BY T16_SEQ_NO"

'response.write w_sSQL & "<br>"

        iRet = gf_GetRecordset(w_Rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
			m_sErrMsg = "���я�񂪎擾�ł��܂���ł����B"
            f_GetResultData = 99
            Exit Do
        End If

		'==============================
		'//���я���z��Ɋi�[����
		'==============================

		i = 0
		Do Until w_Rs.EOF

			'�J�݂��Ă��Ȃ��Ȗڂ͕\�����Ȃ�
			do 
                '2003.10.28 upd_s

				'If f_GetKaisetu(gf_SetNull2String(w_Rs("T16_KAMOKU_CD"))) = False Then
				'	Exit Do
				'End If
		        if m_SchoolNo = "11" or m_SchoolNo = "55" then  '��������̏ꍇ
					Select Case cint(gf_SetNull2Zero(m_sSikenKBN))
						Case C_SIKEN_ZEN_TYU ,C_SIKEN_ZEN_KIM   '�O�����Ԏ���,�O����������

			                IF  ((gf_SetNull2String(w_Rs("T16_KAISETU")) = "0") OR (gf_SetNull2String(w_Rs("T16_KAISETU")) = "1")) THEN
	                           Exit Do
							End If
						Case C_SIKEN_KOU_TYU    '������Ԏ���
			                IF  ((gf_SetNull2String(w_Rs("T16_KAISETU")) = "0") OR (gf_SetNull2String(w_Rs("T16_KAISETU") = "2"))) THEN
								Exit Do
							End If
						Case C_SIKEN_KOU_KIM	'�w�N���͑S�ĕ\������
			                IF  ((gf_SetNull2String(w_Rs("T16_KAISETU")) = "0") OR (gf_SetNull2String(w_Rs("T16_KAISETU") = "1")) OR (gf_SetNull2String(w_Rs("T16_KAISETU") = "2"))) THEN
								Exit Do
							End If
						Case Else
					End Select
				Else
	                '2003.10.28 upd_e
					Select Case cint(gf_SetNull2Zero(m_sSikenKBN))
						Case C_SIKEN_ZEN_TYU ,C_SIKEN_ZEN_KIM   '�O�����Ԏ���,�O����������

			                IF  ((gf_SetNull2String(w_Rs("T16_KAISETU")) = "0") OR (gf_SetNull2String(w_Rs("T16_KAISETU")) = "1")) THEN
	                           Exit Do
							End If
						Case C_SIKEN_KOU_TYU,C_SIKEN_KOU_KIM    '������Ԏ���,�����������
			                IF  ((gf_SetNull2String(w_Rs("T16_KAISETU")) = "0") OR (gf_SetNull2String(w_Rs("T16_KAISETU") = "2"))) THEN
								Exit Do
							End If
						Case Else
					End Select

				END IF

				w_Rs.MoveNext

				If w_Rs.EOF = True Then 
					Exit Do
				End IF
			Loop

			If w_Rs.EOF = True Then 
				Exit Do
			End IF

			'//�I���׸ޏ�����
			w_bSelect = False


			'//�Ȗڂ��K�C�̏ꍇ(C_HISSEN_HIS = 1 :�K�C)
			If cint(gf_SetNull2Zero(w_Rs("T16_HISSEN_KBN"))) = C_HISSEN_HIS Then
				w_bSelect = True
			Else

				'//�Ȗڂ��I���Ȗڂ̏ꍇ(C_HISSEN_SEN = 2 : �I��)
				If cint(gf_SetNull2Zero(w_Rs("T16_HISSEN_KBN"))) = C_HISSEN_SEN Then

					'//�I���Ȗڂ�I�����Ă���ꍇ(C_SENTAKU_YES = 1�F�I������)
					If cint(gf_SetNull2Zero(w_Rs("T16_SELECT_FLG"))) = C_SENTAKU_YES Then
						w_bSelect = True
					Else
						w_bSelect = False
					End If

				End If

			End If

			If w_bSelect = True Then
			'	Redim Preserve m_AryResult(5,i)
				Redim Preserve m_AryResult(6,i)	'UPDATE 2008/03/10

				'//������
				m_AryResult(0,i) = ""
				m_AryResult(1,i) = ""
				m_AryResult(2,i) = ""
				m_AryResult(3,i) = ""
				m_AryResult(4,i) = ""
				m_AryResult(5,i) = ""

				m_AryResult(6,i) = ""

				m_AryResult(0,i) = w_Rs("T16_KAMOKUMEI")	'//�Ȗږ���
				m_AryResult(1,i) = w_Rs("T16_HAITOTANI")	'//�z���P��

				'INS 2008/03/10
				m_AryResult(6,i) = w_Rs("M100_SEISEKI_INP")			'//�x����

				IF Cint(m_AryResult(6,i)) = 0 then
					m_AryResult(2,i) = w_Rs("HTEN")			'//�]���_
				else
					m_AryResult(2,i) = w_Rs("HYOKA")		'//�]��
				end if
				'INS END 2008/03/10


               m_AryResult(3,i) = w_Rs("JIKAN")			'//���Ǝ��Ԑ�


				'//���Ǝ��Ԑ����擾
'				w_bRet = gf_SouJugyo(w_lJikan,w_Rs("T16_KAMOKU_CD"),m_sGakunen,m_sClassNo,w_sKaisibi,w_sSyuryobi,m_iSyoriNen)
'				if w_bRet <> True Then
'					m_AryResult(3,i) = ""					'//���Ǝ��Ԑ��擾���s
'				Else
'					m_AryResult(3,i) = w_lJikan				'//���Ǝ��Ԑ�
'				End If

				m_AryResult(4,i) = w_Rs("KEKA")				'//���ې�
				m_AryResult(5,i) = w_Rs("CHIKAI")			'//�x����

                if m_SchoolNo = "11" then  '��������̏ꍇ

					Select Case cint(gf_SetNull2Zero(m_sSikenKBN))

					  		Case C_SIKEN_ZEN_TYU    '�O�����Ԏ���
				  				  m_AryResult(3,i) = w_Rs("JIKAN")			'//���Ǝ��Ԑ�
			     				  m_AryResult(4,i) = w_Rs("KEKA")	    	'//���ې�
			     	              m_AryResult(5,i) = w_Rs("CHIKAI") 		'//�x����

							Case C_SIKEN_ZEN_KIM    '�O����������

				  				  m_AryResult(3,i) = cint(gf_SetNull2Zero(w_Rs("JIKAN")))	+ cint(gf_SetNull2Zero(w_Rs("JIKAN2")))		'//���Ǝ��Ԑ�
			     				  m_AryResult(4,i) = cint(gf_SetNull2Zero(w_Rs("KEKA")))	+ cint(gf_SetNull2Zero(w_Rs("KEKA2")))     	'//���ې�
			     	              m_AryResult(5,i) = cint(gf_SetNull2Zero(w_Rs("CHIKAI"))) + cint(gf_SetNull2Zero(w_Rs("CHIKAI2")))		'//�x����

							Case C_SIKEN_KOU_TYU    '������Ԏ���
				  				  m_AryResult(3,i) = cint(gf_SetNull2Zero(w_Rs("JIKAN")))	+ cint(gf_SetNull2Zero(w_Rs("JIKAN2"))) + cint(gf_SetNull2Zero(w_Rs("JIKAN3")))						'//���Ǝ��Ԑ�
			     				  m_AryResult(4,i) = cint(gf_SetNull2Zero(w_Rs("KEKA")))	+ cint(gf_SetNull2Zero(w_Rs("KEKA2")))   + cint(gf_SetNull2Zero(w_Rs("KEKA3")))     					'//���ې�
			     	              m_AryResult(5,i) = cint(gf_SetNull2Zero(w_Rs("CHIKAI"))) + cint(gf_SetNull2Zero(w_Rs("CHIKAI2"))) + cint(gf_SetNull2Zero(w_Rs("CHIKAI3")))						'//�x����

							Case C_SIKEN_KOU_KIM    '�����������
				  				  m_AryResult(3,i) = cint(gf_SetNull2Zero(w_Rs("JIKAN")))	+ cint(gf_SetNull2Zero(w_Rs("JIKAN2"))) + cint(gf_SetNull2Zero(w_Rs("JIKAN3"))) + cint(gf_SetNull2Zero(w_Rs("JIKAN4")))		'//���Ǝ��Ԑ�
			     				  m_AryResult(4,i) = cint(gf_SetNull2Zero(w_Rs("KEKA")))	+ cint(gf_SetNull2Zero(w_Rs("KEKA2")))   + cint(gf_SetNull2Zero(w_Rs("KEKA3")))   + cint(gf_SetNull2Zero(w_Rs("KEKA4")))   	'//���ې�
			     	              m_AryResult(5,i) = cint(gf_SetNull2Zero(w_Rs("CHIKAI"))) + cint(gf_SetNull2Zero(w_Rs("CHIKAI2"))) + cint(gf_SetNull2Zero(w_Rs("CHIKAI3"))) + cint(gf_SetNull2Zero(w_Rs("CHIKAI4")))	'//�x����
					End Select
                END IF
				i = i + 1
			End If

			w_Rs.MoveNext
		Loop

		m_iCnt = i-1

        '//����I��
        f_GetResultData = 0


        Exit Do
    Loop

	'//ں��޾��CLOSE
	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �Ȗڂ̊J�ݎ������擾
'*  [����]  �Ȃ�
'*  [�ߒl]  True�F�J�݂���AFalse�F�J�݂Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetKaisetu(p_sKamokuCd)
    Dim w_sSQL              '// SQL��
    Dim w_iRet              '// �߂�l
	Dim rs

	ON ERROR RESUME NEXT
	ERR.CLEAR

	f_GetKaisetu = False

	Do

		w_sSQL =  ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T15_KAISETU" & m_sGakunen
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "        T15_KAMOKU_CD = '" & p_sKamokuCd & "'"
		w_sSQL = w_sSQL & vbCrLf & "    AND T15_NYUNENDO = " & m_iSyoriNen - m_sGakunen + 1
		w_sSQL = w_sSQL & vbCrLf & "    AND ( ("

		Select Case cint(m_sSikenKbn) '�I�񂾎����ɂ���āA�擾�Ȗڂ̊J�݊��Ԃ�ς���
			Case cint(C_SIKEN_ZEN_TYU)
				w_sSQL = w_sSQL & vbCrLf & " T15_KAISETU" & m_sGakunen & "=" & C_KAI_ZENKI & " "

			Case cint(C_SIKEN_ZEN_KIM)
				w_sSQL = w_sSQL & vbCrLf & " T15_KAISETU" & m_sGakunen & "=" & C_KAI_ZENKI & " "

			Case cint(C_SIKEN_KOU_TYU)
				'w_sSQL = w_sSQL & vbCrLf & " T15_KAISETU" & m_sGakunen & "=" & C_KAI_ZENKI & " OR "
				w_sSQL = w_sSQL & vbCrLf & " T15_KAISETU" & m_sGakunen & "=" & C_KAI_KOUKI & " "

			Case cint(C_SIKEN_KOU_KIM)
				w_sSQL = w_sSQL & vbCrLf & " T15_KAISETU" & m_sGakunen & "=" & C_KAI_ZENKI & " OR "
				w_sSQL = w_sSQL & vbCrLf & " T15_KAISETU" & m_sGakunen & "=" & C_KAI_KOUKI & " "

		End Select

		w_sSQL = w_sSQL & vbCrLf & "    )"
		w_sSQL = w_sSQL & vbCrLf & "    OR ("
		w_sSQL = w_sSQL & vbCrLf & "       T15_KAISETU" & m_sGakunen & "=" & C_KAI_TUNEN & " "
		w_sSQL = w_sSQL & vbCrLf & "    ) ) "

'response.write w_ssql & "<br>"

		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit function
		End If

		If rs.EOF= False Then
		    Call gf_closeObject(rs)
			Exit Function
		End If 

		Exit do 
	Loop

	'//�߂�l���Z�b�g
	f_GetKaisetu = True

	'//RS Close
    Call gf_closeObject(rs)

	ERR.CLEAR

End Function


'********************************************************************************
'*  [�@�\]  �w�Ȃ̗��̂��擾
'*  [����]  p_sGakkaCd : �w��CD
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetGakkaNm(p_iGakunen,p_iClass)
    Dim w_sSQL              '// SQL��
    Dim w_iRet              '// �߂�l
	Dim w_sName 
	Dim rs

	ON ERROR RESUME NEXT
	ERR.CLEAR

	f_GetGakkaNm = ""
	w_sName = ""

	Do

		w_sSQL =  ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_GAKKAMEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA "
		w_sSQL = w_sSQL & vbCrLf & "  ,M05_CLASS "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M02_GAKKA.M02_GAKKA_CD = M05_CLASS.M05_GAKKA_CD "
		w_sSQL = w_sSQL & vbCrLf & "  AND M02_GAKKA.M02_NENDO = M05_CLASS.M05_NENDO"
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_NENDO=" & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_GAKUNEN=" & p_iGakunen
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_CLASSNO=" & p_iClass

		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit function
		End If

		If rs.EOF= False Then
			w_sName = rs("M02_GAKKAMEI")
		End If 

		Exit do 
	Loop

	'//�߂�l���Z�b�g
	f_GetGakkaNm = w_sName

	'//RS Close
    Call gf_closeObject(rs)

	ERR.CLEAR

End Function

'********************************************************************************
'*  [�@�\]  �Ȏ��A�S�C�������擾
'*  [����]  p_sGakkaCd : �w��CD
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetGakuseiInfo()
    Dim w_sSQL              '// SQL��
    Dim w_iRet              '// �߂�l
	Dim rs

	ON ERROR RESUME NEXT
	ERR.CLEAR

	f_GetGakuseiInfo = 1
	m_iSikiji = ""
	m_iSyoken = ""
	m_iAverage = 0
	
	Do

		w_sSQL =  ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "

		'//�����敪�ɂ��ꍇ�킯
		Select Case cint(gf_SetNull2Zero(m_sSikenKBN))

			Case C_SIKEN_ZEN_TYU    '�O�����Ԏ���
				w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_SEKIJI_TYUKAN_Z AS SEKIJI"			'//�Ȏ�
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_NINZU_TYUKAN_Z AS NINZU"			'//�N���X�l��
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_SYOKEN_TYUKAN_Z AS SYOKEN "			'//����
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_HEIKIN_TYUKAN_Z AS HEIKIN "			'//���ϓ_
			Case C_SIKEN_ZEN_KIM    '�O����������
				w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_SEKIJI_KIMATU_Z AS SEKIJI"			'//�Ȏ�
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_NINZU_KIMATU_Z AS NINZU"			'//�N���X�l��
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_SYOKEN_KIMATU_Z AS SYOKEN "			'//����
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_HEIKIN_KIMATU_Z AS HEIKIN "			'//���ϓ_
			Case C_SIKEN_KOU_TYU    '������Ԏ���
				w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_SEKIJI_TYUKAN_K  AS SEKIJI"			'//�Ȏ�
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_NINZU_TYUKAN_K  AS NINZU"			'//�N���X�l��
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_SYOKEN_TYUKAN_K AS SYOKEN "			'//����
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_HEIKIN_TYUKAN_K AS HEIKIN "			'//���ϓ_
			Case C_SIKEN_KOU_KIM    '�����������
				w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_SEKIJI  AS SEKIJI"					'//�Ȏ�
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_CLASSNINZU  AS NINZU"				'//�N���X�l��
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_SYOKEN_KIMATU_K AS SYOKEN"			'//����
				w_sSQL = w_sSQL & vbCrLf & "  ,T13_GAKU_NEN.T13_HEIKIN_KIMATU_K AS HEIKIN "			'//���ϓ_
			Case Else
				'//�V�X�e���G���[
	            m_sErrMsg = "������񂪂���܂���B"
				Exit Do
		End Select

		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO=" & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & "  AND T13_GAKU_NEN.T13_GAKUSEI_NO='" & m_sGakusei & "'"

		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			f_GetGakuseiInfo = 99
			'ں��޾�Ă̎擾���s
			Exit function
		End If

		If rs.EOF= False Then
			m_iSikiji = rs("SEKIJI")
			m_iMemberCnt = rs("NINZU")
			m_iSyoken = rs("SYOKEN")
			m_iAverage = rs("HEIKIN")
		End If 

		f_GetGakuseiInfo = 0
		Exit do 
	Loop

	'//RS Close
    Call gf_closeObject(rs)

	ERR.CLEAR

End Function

'********************************************************************************
'*  [�@�\]  �\������(����)���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Function f_GetDisp_Data_Siken()
    Dim w_iRet
    Dim w_sSQL
    Dim rs
	Dim w_sSikenName

    On Error Resume Next
    Err.Clear

    f_GetDisp_Data_Siken = ""
	w_sSikenName = ""

    Do

        '�����}�X�^���f�[�^���擾
        w_sSql = ""
        w_sSql = w_sSql & vbCrLf & " SELECT "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI "
        w_sSql = w_sSql & vbCrLf & " FROM "
        w_sSql = w_sSql & vbCrLf & "  M01_KUBUN "
        w_sSql = w_sSql & vbCrLf & " WHERE "
        w_sSql = w_sSql & vbCrLf & "      M01_KUBUN.M01_NENDO=" & m_iSyoriNen
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_DAIBUNRUI_CD= " & C_SIKEN
        w_sSql = w_sSql & vbCrLf & "  AND M01_KUBUN.M01_SYOBUNRUI_CD=" & m_sSikenKBN

        iRet = gf_GetRecordset(rs, w_sSQL)
        If iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetDisp_Data_Siken = 99
            Exit Do
        End If

        If rs.EOF = False Then
            w_sSikenName = rs("M01_SYOBUNRUIMEI")
        End If

        Exit Do
    Loop

	'//�߂�l�Z�b�g
    f_GetDisp_Data_Siken = w_sSikenName

    Call gf_closeObject(rs)

End Function

'*******************************************************************************
' �@�@�@�\�F�Ȗڕ]���擾
' �ԁ@�@�l�F
' ���@�@���Fp_iTensu - �_��(IN)
'           
'
' �@�\�ڍׁF�_������]��NO��p_uData�ɕ]���A�]��A���_�Ȗڂ�ݒ肷��
' ���@�@�l�F�]��NO�����łɕ������Ă���ꍇ�ɂ͒���call����
'           �]��NO��������Ȃ��Ƃ��́Agf_GetKamokuTensuHyoka��call
'           2002.06.19 ����
'*******************************************************************************
Function f_GetTensuHyoka(p_iTensu)
    Dim w_oRecord
    Dim w_sSql
    
    Const C_HYOKA_FUKA = 1
    
    On Error Resume Next
    
    f_GetTensuHyoka = ""
	
	if gf_SetNull2String(p_iTensu) = "" then exit function
	
    w_sSql = ""
    w_sSql = w_sSql & " SELECT "
    w_sSql = w_sSql & " 	M08_HYOKA_SYOBUNRUI_RYAKU "
    w_sSql = w_sSql & " FROM "
    w_sSql = w_sSql & " 	M08_HYOKAKEISIKI "
    w_sSql = w_sSql & " WHERE "
    w_sSql = w_sSql & " 	M08_MIN <= " & p_iTensu								'�_��
    w_sSql = w_sSql & " AND M08_MAX >= " & p_iTensu
    w_sSql = w_sSql & " AND M08_NENDO = " & m_iSyoriNen							'�N�x
    w_sSql = w_sSql & " AND M08_HYOKA_TAISYO_KBN = " & C_HYOKA_TAISHO_IPPAN		'��ʊw��
    
    If gf_GetRecordset(w_oRecord,w_sSql) <> 0 Then : exit function
    
    '�Ȗ�M�Ȃ����G���[
    if w_oRecord.EOF Then Exit Function
    
    '�f�[�^�Z�b�g
    if cint(gf_SetNull2Zero(w_oRecord("M08_HYOKA_SYOBUNRUI_RYAKU"))) = C_HYOKA_FUKA then
    	f_GetTensuHyoka = "*"
    end if
    
    Call gf_closeObject(w_oRecord)
    
End Function


Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

	Dim w_iTaniKei
	Dim w_iHyokaKei
	Dim w_iJikanKei
	Dim w_iKekkaKei
	Dim w_iTikokuKei
	Dim w_iSeiAverage

%>

	<html>

	<head>
	<title>�l�ʐ��шꗗ</title>
	<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
	<SCRIPT language="JavaScript">
	<!--
    //************************************************************
    //  [�@�\]  �L�����Z���{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_Cansel(){

        document.frm.action="default.asp";
        document.frm.target="<%=C_MAIN_FRAME%>";
        document.frm.submit();
    
    }

    //************************************************************
    //  [�@�\]  �O��,���փ{�^���������ꂽ�Ƃ�
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //************************************************************
    function f_NextPage(p_FLG){

		if( p_FLG == 1){
			document.frm.txtGakusei.value = document.frm.txtBeforGakuNo.value;
		}else{
        	document.frm.txtGakusei.value = document.frm.txtAfterGakuNo.value;
		}

		document.frm.action="sei0300_main.asp";
		document.frm.target="_self";
		document.frm.submit();
    }

	//-->
	</SCRIPT>


	<link rel=stylesheet href="../../common/style.css" type=text/css>
	</head>
	<body>
	<center>
	<form name="frm" METHOD="post">
	<% call gs_title(" �l�ʐ��шꗗ "," ��@�� ") %>
	<BR>

	<%Do %>
		<table border="0" width="500" class=hyo align="center">
			<tr>
				<th width="500" class="header" colspan="4"><%=f_GetDisp_Data_Siken()%><br></th>
			</tr>
			<tr>
				<th width="50" class="header2">�N���X</th>
	<%If m_sKengen <> C_KENGEN_SEI0300_GAK then%>
				<td width="150" align="center" class="detail"><%=m_sGakunen%>-<%=m_sClassNo%> [<%=f_GetGakkaNm(m_sGakunen,m_sClassNo)%>]</td>
	<%Else%>
				<td width="150" align="center" class="detail"><%=m_sGakunen%>�N�@<%=gf_GetGakkaNm(m_iSyoriNen,m_sGakkaNo)%></td>
	<%End If%>
				<th width="50" class="header2">���@��</th>
				<td width="250" align="left" class="detail">�@( <%=m_sGakusekiNo%> )<%=m_sName%></td>
			</tr>
		</table>

		<!--�{�^��-->
		<BR>
		<table border="0" width="250">
		    <tr>
		        <td valign="top" align="center">
		            <input type="button" value="�@�O�@�ց@" class="button" <%If m_sBeforGakuNo = "" Then %> DISABLED <%End If%>  onclick="javascript:f_NextPage(1)">
		        </td>
		        <td valign="top" align="center">
		            <input type="button" value="�L�����Z��" class="button" onclick="javascript:f_Cansel()">
		        </td>
		        <td valign="top" align="center">
		            <input type="button" value="�@���@�ց@" class="button" <%If m_sAfterGakuNo = "" Then%>  DISABLED <%End If%>  onclick="javascript:f_NextPage(2)">
		        </td>
		    </tr>
		</table>
		<br>

		<!--�S�C����-->
		<table class="hyo" border="1" align="center" width="70%">
		<tr>
			<th class="header" colspan="6">�S�@�C�@���@��</th>
		</tr>
		<tr>
			<td class="detail" colspan="6" height="35" valign="top"><%=m_iSyoken%><br></td>
		</tr>

		<!--���ו�-->
		<tr>
			<th class="header" rowspan="2" width="31%">�ȁ@�ځ@��</th>
			<th class="header" rowspan="2" width="13%">�P�ʐ�</th>
			<th class="header" rowspan="2" width="13%">����<br>�]��</th>
			<th class="header" rowspan="2" width="13%">����<br>���Ԑ�</th>
			<th class="header" colspan="2" width="20%">�o�@�ȁ@��@��</th>
		</tr>
		<tr>
			<th class="header">����<br>����</th>
			<th class="header">�x��<br>��</th>
		</tr>

		<% 
		'//���v������
		w_iTaniKei   = 0
		w_iHyokaKei  = 0
		w_iJikanKei  = 0
		w_iKekkaKei  = 0
		w_iTikokuKei = 0

		For i = 0 To m_iCnt

			call gs_cellPtn(w_cell) %>
			<tr>
				<td class="<%=w_cell%>" align="left" ><%=m_AryResult(0,i)%><br></td>
				<td class="<%=w_cell%>" align="right"><%=FormatNumber(cint(gf_SetNull2Zero(m_AryResult(1,i))),1)%><br></td>

				<% If Cint(m_AryResult(6,i)) = 0 Then %>

					<td class="<%=w_cell%>" align="right"><%=f_GetTensuHyoka(m_AryResult(2,i))%>�@<%=m_AryResult(2,i)%></td>

				<% ELSE %>

					<td class="<%=w_cell%>" align="right">�@<%=m_AryResult(2,i)%></td>

				<% END IF%>
				<td class="<%=w_cell%>" align="right"><%=gf_SetNull2Zero(m_AryResult(3,i))%><br></td>
				<td class="<%=w_cell%>" align="right"><%=gf_SetNull2Zero(m_AryResult(4,i))%><br></td>
				<td class="<%=w_cell%>" align="right"><%=gf_SetNull2Zero(m_AryResult(5,i))%><br></td>
			</tr>

			<%
			'//�P�ʐ����v
			w_iTaniKei = w_iTaniKei + cint(gf_SetNull2Zero(m_AryResult(1,i)))

			'//���ѕ]�����v
			If Cint(m_AryResult(6,i)) = 0 Then
				w_iHyokaKei = w_iHyokaKei + cint(gf_SetNull2Zero(m_AryResult(2,i)))
			End if

			'//���Ǝ��Ԑ����v
			w_iJikanKei = w_iJikanKei + cint(gf_SetNull2Zero(m_AryResult(3,i)))

			'//���ې����v
			w_iKekkaKei = w_iKekkaKei + cint(gf_SetNull2Zero(m_AryResult(4,i)))

			'//�x���񐔍��v
			w_iTikokuKei = w_iTikokuKei + cint(gf_SetNull2Zero(m_AryResult(5,i)))
			%>

		<% next %>

		<!--���v-->
		<tr>
			<td class="NOCHANGE">���v</td>
			<td class="NOCHANGE" align="right"><%=FormatNumber(w_iTaniKei,1)%></td>
			<td class="NOCHANGE" align="right"><%=w_iHyokaKei%></td>
			<td class="NOCHANGE" align="right"><%=w_iJikanKei%></td>
			<td class="NOCHANGE" align="right"><%=w_iKekkaKei%></td>
			<td class="NOCHANGE" align="right"><%=w_iTikokuKei%></td>
		</tr>
		
		<!--����-->
		<tr>
			<td class="CELL2">����</td>
			<td class="CELL2" align="right">�\</td>
			<td class="CELL2" align="right"><%=m_iAverage%></td>
			<td class="CELL2" colspan="3" align="right">�Ȏ��@<%=m_iSikiji%>�ʁ^<%=m_iMemberCnt%>�l��</td>
		</tr>
		</table>

		<!--�{�^��-->
		<BR>
		<table border="0" width="250">
		    <tr>
		        <td valign="top" align="center">
		            <input type="button" value="�@�O�@�ց@" class="button" <%If m_sBeforGakuNo = "" Then %> DISABLED <%End If%>  onclick="javascript:f_NextPage(1)">
		        </td>
		        <td valign="top" align="center">
		            <input type="button" value="�L�����Z��" class="button" onclick="javascript:f_Cansel()">
		        </td>
		        <td valign="top" align="center">
		            <input type="button" value="�@���@�ց@" class="button" <%If m_sAfterGakuNo = "" Then%>  DISABLED <%End If%>  onclick="javascript:f_NextPage(2)">
		        </td>
		    </tr>
		</table>

		<%Exit Do%>
	<%Loop%>

	<input type="hidden" name="txtSikenKBN" value="<%=m_sSikenKBN%>">
	<input type="hidden" name="txtGakuNo"   value="<%=m_sGakunen%>">
	<input type="hidden" name="txtClassNo"  value="<%=m_sClassNo%>">
	<input type="hidden" name="txtGakkaNo"  value="<%=m_sGakkaNo%>">
	<input type="hidden" name="txtGakusei"  value="<%=m_sGakusei%>">

	<input type="hidden" name="txtBeforGakuNo" value="<%=m_sBeforGakuNo%>">
	<input type="hidden" name="txtAfterGakuNo" value="<%=m_sAfterGakuNo%>">

	</form>
	</center>
	</body>
	</html>
<%
    '---------- HTML END   ----------
End Sub
%>
