<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���Əo������
' ��۸���ID : kks/kks0110/kks0110_main.asp
' �@	  �\: ���y�[�W ���Əo�����͂̈ꗗ���X�g�\�����s��
'-------------------------------------------------------------------------
' ��	  ��: NENDO 		 '//�����N
'			  KYOKAN_CD 	 '//����CD
'			  GAKUNEN		 '//�w�N
'			  CLASSNO		 '//�׽No
'			  TUKI			 '//��
' ��	  ��:
' ��	  �n: NENDO 		 '//�����N
'			  KYOKAN_CD 	 '//����CD
'			  GAKUNEN		 '//�w�N
'			  CLASSNO		 '//�׽No
'			  TUKI			 '//��
' ��	  ��:
'			�������\��
'				���������ɂ��Ȃ��s���o�����͂�\��
'			���o�^�{�^���N���b�N��
'				���͏���o�^����
'-------------------------------------------------------------------------
' ��	  ��: 2001/07/02 �ɓ����q
' ��	  �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ��CONST /////////////////////////////
	Const C_SYOBUNRUICD_IPPAN = 4	'//���ȋ敪(0:�o��,1:����,2:�x��,3:����,4:����,�c)
	Const C_IDO_MAX_CNT = 8			'//�ő�ړ���(T13�ړ����o�^�p�t�B�[���h��)
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
	Public	m_bErrFlg			'�װ�׸�
	Public	m_bDaigae			 '��֗��w���擾�׸�

	'�擾�����f�[�^�����ϐ�
	Public m_iSyoriNen		'//�����N�x
	Public m_iKyokanCd		'//����CD
	Public m_sGakunen		'//�w�N
	Public m_sClassNo		'//�׽NO
	Public m_sTuki			'//��
	Public m_sZenki_Start	'//�O���J�n��
	Public m_sKouki_Start	'//����J�n��
	Public m_sKouki_End 	'//����I����
	Public m_sEndDay		'//���͂ł��Ȃ��Ȃ��

	Public m_sGakki 		'//�w��
	Public m_sGakki_Kbn 	'//�w���敪
	Public m_sKamokuCd		'//�ۖ�CD
	Public m_sSyubetu		'//���Ǝ��(TUJO:�ʏ����,TOKU:���ʊ���,KBTU:�ʎ���)
	Public m_sHissenKbn 	'//�K�I�敪
	Public m_iTani			'//�P�����̒P�ʐ�
	
	'ں��ރZ�b�g
	Public m_Rs_M			'//recordset���׏��
	Public m_Rs_D			'//recordset��֗��w��
	Public m_Rs_G			'//recordset�s���o�����

	Public m_AryHead()		'//�w�b�_���i�[�z��
	Public m_iRsCnt 		'//�w�b�_ں��ސ�
	Public m_iRuiKeiCnt 	'//�݌v�J�E���g
	Public m_AryRuiKei()	'//�݌v�i�[�z��

	Public m_iTukiKeiCnt	'//���v�J�E���g
	Public m_AryTukiKei()	'//���v�i�[�z��

	Public m_AryKesseki
	Public m_iSyubetu
	Public m_iSikenKbn

	Public m_sLevelFlg
	
	Public m_iShikenInsertType	'//�������ѓo�^����
								'C_SIKEN_ZEN_TYU = 1 '�O�����Ԏ���
								'C_SIKEN_ZEN_KIM = 2 '�O����������
								'C_SIKEN_KOU_TYU = 3 '������Ԏ���
								'C_SIKEN_KOU_KIM = 4 '�����������
		
	
'///////////////////////////���C������/////////////////////////////

	'Ҳ�ٰ�ݎ��s
	Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////

Sub Main()
'********************************************************************************
'*	[�@�\]	�{ASP��Ҳ�ٰ��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************

	Dim w_iRet				'// �߂�l
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

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
		
		'// �w�b�_���X�g���擾
		w_iRet = f_Get_HeadData()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'// ���k���X�g���擾
		w_iRet = f_Get_DetailData()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'//���k��񂪂Ȃ��ꍇ
		If m_Rs_M.EOF Then
			'//�󔒃y�[�W�\��
			Call showWhitePage("���k��񂪂���܂���")
			Exit Do
		End If
		
		'//�o�����׏��擾
		w_iRet = f_Get_AbsInfo()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'//�o���w���݌v�擾
		w_iRet = f_Get_AbsInfo_RuiKei()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'// �Ǘ��}�X�^���A�o�����ۂ̎������擾
		w_iRet = gf_GetKanriInfo(m_iSyoriNen,m_iSyubetu)
		If w_iRet <> 0 Then 
			m_bErrFlg = True
			Exit Do
		End If
		
		'//�o�����v�擾
		w_iRet = f_Get_AbsInfo_TukiKei()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If
		
		'//���Ɠ������Ȃ��ꍇ
		If m_iRsCnt < 0 Then
			'//�󔒃y�[�W�\��
			Call showWhitePage("���Ɠ���������܂���")
		   Exit Do
		End If
		
		'// �f�[�^�\���y�[�W��\��
		Call showPage()

		Exit Do
	Loop

	'// �װ�̏ꍇ�ʹװ�߰�ނ�\��
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If
	
	'// �I������
	Call gf_closeObject(m_Rs_M)
	Call gf_closeObject(m_Rs_D)
	Call gf_closeObject(m_Rs_G)
	Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*	[�@�\]	�ϐ�������
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub s_ClearParam()
	m_iSyoriNen = ""
	m_iKyokanCd = ""
	m_sGakunen	= ""
	m_sClassNo	= ""
	m_sTuki 	= ""
	m_sGakki	= ""
	m_sKamokuCd = ""
	m_sSyubetu	= ""
	m_iTani		= 0
	m_iShikenInsertType = 0
	
End Sub

'********************************************************************************
'*	[�@�\]	�S���ڂɈ����n����Ă����l��ݒ�
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub s_SetParam()

	m_sZenki_Start = trim(Request("Tuki_Zenki_Start"))
	m_sKouki_Start = trim(Request("Tuki_Kouki_Start"))
	m_sKouki_End   = trim(Request("Tuki_Kouki_End"))
	m_iTani = Session("JIKAN_TANI") '�P�����̒P�ʐ�
	
	m_iSyoriNen = trim(Request("NENDO"))
	m_iKyokanCd = trim(Request("KYOKAN_CD"))

	m_sTuki 	= trim(Request("TUKI"))
	m_sGakki	= trim(Request("GAKKI"))

	m_sSyubetu	= trim(Request("SYUBETU"))
	m_sGakunen	= trim(Request("GAKUNEN"))
	m_sClassNo	= trim(Request("CLASSNO"))
	m_sKamokuCd = trim(Request("KAMOKU_CD"))

	If m_sGakki = "ZENKI" Then
		m_sGakki_Kbn = cstr(C_GAKKI_ZENKI)
	Else
		m_sGakki_Kbn = cstr(C_GAKKI_KOUKI)
	End If

	call gf_Get_SyuketuEnd(cint(m_sGakunen),m_sEndDay)

End Sub

'********************************************************************************
'*	[�@�\]	���t�E�j���E���Ԃ̃w�b�_���擾�������s��
'*	[����]	�Ȃ�
'*	[�ߒl]	0:���擾���� 99:���s
'*	[����]	
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

			'// �ʏ�A���ʎ��Ƃ̏ꍇ
			w_sSQL = ""
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
			w_sSQL = w_sSQL & vbCrLf & " AND to_date(A.T32_HIDUKE,'YYYY/MM/DD')>='" & w_sSDate & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND to_date(A.T32_HIDUKE,'YYYY/MM/DD')<'"  & w_sEDate & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND B.T20_NENDO="		& cInt(m_iSyoriNen)
			w_sSQL = w_sSQL & vbCrLf & " AND B.T20_GAKKI_KBN='" & m_sGakki_Kbn & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND B.T20_GAKUNEN= "	& cInt(m_sGakunen)
			w_sSQL = w_sSQL & vbCrLf & " AND B.T20_CLASS= " 	& cInt(m_sClassNo)
			w_sSQL = w_sSQL & vbCrLf & " AND B.T20_KAMOKU='"	& trim(m_sKamokuCd) & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND B.T20_KYOKAN='"	& m_iKyokanCd & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND A.T32_GYOJI_CD=0"
			w_sSQL = w_sSQL & vbCrLf & " AND A.T32_KYUJITU_FLG='0' "
			w_sSQL = w_sSQL & vbCrLf & " GROUP BY A.T32_HIDUKE,B.T20_YOUBI_CD,B.T20_JIGEN "
			w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T32_HIDUKE,B.T20_JIGEN"

		Else
			'// �ʏ�A���ʎ��Ƃ̏ꍇ
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
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_NENDO = A.T32_NENDO"
			w_sSQL = w_sSQL & vbCrLf & " AND to_date(A.T32_HIDUKE,'YYYY/MM/DD')>='" & w_sSDate & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND to_date(A.T32_HIDUKE,'YYYY/MM/DD')<'"  & w_sEDate & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_NENDO="		& cInt(m_iSyoriNen)
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_GAKKI_KBN=" & m_sGakki_Kbn & " "
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_KAMOKU='"	& trim(m_sKamokuCd) & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND B.T23_KYOKAN='"	& m_iKyokanCd & "'"
			w_sSQL = w_sSQL & vbCrLf & " AND A.T32_GYOJI_CD=0"
			w_sSQL = w_sSQL & vbCrLf & " AND A.T32_KYUJITU_FLG='0' "
			w_sSQL = w_sSQL & vbCrLf & " GROUP BY A.T32_HIDUKE,B.T23_YOUBI_CD,B.T23_JIGEN "
			w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T32_HIDUKE,B.T23_JIGEN"
		End If
		
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
				iRet = f_Get_DateInfo(w_Rs("T32_HIDUKE"),cint(w_Rs("JIGEN")),w_bGyoji)
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
						m_AryHead(0,i) = "" 	'//��
						m_AryHead(1,i) = "" 	'//��
						m_AryHead(2,i) = "" 	'//�j��CD
					Else
						m_AryHead(0,i) = month(gf_SetNull2String(w_Rs("T32_HIDUKE")))	  '//��
						m_AryHead(1,i) = day(gf_SetNull2String(w_Rs("T32_HIDUKE"))) 	  '//��
						m_AryHead(2,i) = gf_SetNull2String(w_Rs("YOUBI_CD"))		  '//�j��CD
					End If

					m_AryHead(3,i) = replace(gf_SetNull2String(w_Rs("JIGEN")),".","$")	'//����
					m_AryHead(4,i) = gf_SetNull2String(w_Rs("T32_HIDUKE"))					'//���t

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
'*	[�@�\]	�擾�������t�E�������A�x���܂��͍s���łȂ���
'*	[����]	�Ȃ�
'*	[�ߒl]	0:���擾���� 99:���s
'*	[����]	
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
		w_sSQL = w_sSQL & vbCrLf & "  AND to_date(A.T32_HIDUKE,'YYYY/MM/DD')='" & p_Hiduke & "'"
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
'*	[�@�\]	���׏����擾����
'*	[����]	�Ȃ�
'*	[�ߒl]	0:���擾���� 99:���s
'*	[����]	
'********************************************************************************
Function f_Get_DetailData()

	Dim w_iRet

	On Error Resume Next
	Err.Clear
	
	f_Get_DetailData = 1

	Do 

		'//���Ǝ�ʂɂ�菈���𕪊�(TUJO:�ʏ����,TOKU:���ʊ���,KBTU:�ʎ���)
		Select Case trim(m_sSyubetu)
		  Case "TUJO" ':�ʏ����

			'//�ʏ���Ǝ擾��
			w_iRet = f_Get_Data_TUJO()
			If w_iRet <> 0 then
				Exit Do
			End If

		  Case "TOKU" ':���ʊ���

			'//���ʎ��Ǝ擾��(�S�C�׽�ꗗ)
			w_iRet = f_Get_Data_TOKU()
			If w_iRet <> 0 then
				Exit Do
			End If

		  Case "KBTU" ':�ʎ���

			'//�ʎ��Ǝ擾��(�ۖڎ󎝂����w���ꗗ)
			w_iRet = f_Get_Data_KOBETU()
			If w_iRet <> 0 then
				Exit Do
			End If

		  Case Else
			'//�V�X�e���G���[
			m_sErrMsg = "�p�����[�^���s�����Ă��܂��B(�V�X�e���G���[)"
		End Select

		f_Get_DetailData = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[�@�\]	�ʏ���ƑI�����N���X�ꗗ���擾
'*	[����]	�Ȃ�
'*	[�ߒl]	0:���擾���� 99:���s
'*	[����]	
'********************************************************************************
Function f_Get_Data_TUJO()

	Dim w_sSQL
	Dim w_Rs
	Dim w_iRet
	Dim w_sLevelFlg

	On Error Resume Next
	Err.Clear
	
	f_Get_Data_TUJO = 1
	w_sLevelFlg = ""

	Do 
		'//���w�N�x(=�����N�x-�w�N+1)
		w_NyuNen = cInt(m_iSyoriNen) - cInt(m_sGakunen) + 1

		'================
		'//���Ə��擾
		'================
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT DISTINCT "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_NENDO, "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKUNEN, "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_CLASSNO, "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKKA_CD, "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_NYUNENDO, "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_KAMOKU_CD, "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_HISSEN_KBN, "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_LEVEL_FLG"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS,T15_RISYU"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKKA_CD = T15_RISYU.T15_GAKKA_CD AND"
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_NENDO=" 	 & cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKUNEN="	 & cInt(m_sGakunen)  & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_CLASSNO="	 & cInt(m_sClassNo)  & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_NYUNENDO="	 & w_NyuNen 		 & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T15_RISYU.T15_KAMOKU_CD='" & trim(m_sKamokuCd) & "'"
		
		w_iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_Get_Data_TUJO = 99
			Exit Do
		End If
		
		If w_Rs.EOF = False Then
			'//���x���ۖ��׸ނ��擾
			w_sLevelFlg = w_Rs("T15_LEVEL_FLG")
			m_sLevelFlg = w_Rs("T15_LEVEL_FLG")
			'//�K�I�敪���擾
			m_sHissenKbn =w_Rs("T15_HISSEN_KBN")
		End If
		
		'//�ʏ�ۖڐ��k�ꗗ�擾
		w_iRet = f_Get_TUJO_Tujyo()
		If w_iRet <> 0 Then
			Exit Do
		End If
		
		'//�ʏ���ƑI�����A��֗��w���ꗗ�擾
		w_iRet = f_Get_TUJO_DaigeRyugak()
		If w_iRet <> 0 Then
			Exit Do
		End If
		
		f_Get_Data_TUJO = 0
		Exit Do
	Loop
	
	'//ں��޾��CLOSE
	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*	[�@�\]	���x���ʉۖڐ��k�ꗗ�擾
'*	[����]	�Ȃ�
'*	[�ߒl]	0:���擾���� 99:���s
'*	[����]	
'********************************************************************************
Function f_Get_TUJO_LevelBetu()

	Dim w_sSQL

	On Error Resume Next
	Err.Clear
	
	f_Get_TUJO_LevelBetu = 1

	Do 

		'// ���x���ʉۖڑI�𐶓k�ꗗ�擾
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT DISTINCT "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_GAKUSEKI_NO AS GAKUSEKI, "
		w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI.T11_SIMEI  AS SIMEI, "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_SELECT_FLG, "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_OKIKAE_FLG, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_NUM,"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO AS GAKUSEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI ,T16_RISYU_KOJIN,T13_GAKU_NEN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI.T11_GAKUSEI_NO = T16_RISYU_KOJIN.T16_GAKUSEI_NO AND"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO AND"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO="			& cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_NENDO="		   & cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_HAITOGAKUNEN="   & cInt(m_sGakunen)  & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_KAMOKU_CD='"	   & m_sKamokuCd	   & "' AND "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_LEVEL_KYOUKAN='" & m_iKyokanCd	   & "'"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY GAKUSEKI"

		iRet = gf_GetRecordset(m_Rs_M, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_Get_TUJO_LevelBetu = 99
			Exit Do
		End If

		f_Get_TUJO_LevelBetu = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[�@�\]	�ʏ�ۖڐ��k�ꗗ�擾
'*	[����]	�Ȃ�
'*	[�ߒl]	0:���擾���� 99:���s
'*	[����]	
'********************************************************************************
Function f_Get_TUJO_Tujyo()

	Dim w_sSQL

	On Error Resume Next
	Err.Clear
	
	f_Get_TUJO_Tujyo = 1

	Do 

		'// �ʏ�ۖڐ��k�ꗗ
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUNEN, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLASS, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEKI_NO AS GAKUSEKI, "
		w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI.T11_SIMEI AS SIMEI, "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_KAMOKU_CD, "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_SELECT_FLG, "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_OKIKAE_FLG,"
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_LEVEL_KYOUKAN,"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_NUM,"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO AS GAKUSEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN ,T16_RISYU_KOJIN,T11_GAKUSEKI"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO = T16_RISYU_KOJIN.T16_GAKUSEI_NO AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO = T16_RISYU_KOJIN.T16_NENDO  AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO="	 & cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUNEN=" & cInt(m_sGakunen)  & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_CLASS="	 & cInt(m_sClassNo)  & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T16_RISYU_KOJIN.T16_KAMOKU_CD='" & m_sKamokuCd & "'"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY GAKUSEKI "
		
		iRet = gf_GetRecordset(m_Rs_M, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_Get_TUJO_Tujyo = 99
			Exit Do
		End If
		
		f_Get_TUJO_Tujyo = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[�@�\]	���ʎ��ƑI�����A�S�C�N���X�ꗗ�擾
'*	[����]	�Ȃ�
'*	[�ߒl]	0:���擾���� 99:���s
'*	[����]	
'********************************************************************************
Function f_Get_Data_TOKU()

	Dim w_sSQL

	On Error Resume Next
	Err.Clear
	
	f_Get_Data_TOKU = 1

	Do 

		'// �S�C�N���X�ꗗ
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_NENDO, "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUNEN," 
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_CLASS, "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUSEKI_NO AS GAKUSEKI, "
		w_sSQL = w_sSQL & vbCrLf & "   T11_GAKUSEKI.T11_SIMEI AS SIMEI, "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_IDOU_NUM,"
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUSEI_NO AS GAKUSEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN,T11_GAKUSEKI "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO AND "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_NENDO=" & cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_GAKUNEN=" & cInt(m_sGakunen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "   T13_GAKU_NEN.T13_CLASS=" & cInt(m_sClassNo)
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY GAKUSEKI "
		
		iRet = gf_GetRecordset(m_Rs_M, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_Get_Data_TOKU = 99
			Exit Do
		End If
		
		f_Get_Data_TOKU = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[�@�\]	�ʎ��ƑI�����A���w���ꗗ�擾
'*	[����]	�Ȃ�
'*	[�ߒl]	0:���擾���� 99:���s
'*	[����]	
'********************************************************************************
Function f_Get_Data_KOBETU()

	Dim w_sSQL

	On Error Resume Next
	Err.Clear
	
	f_Get_Data_KOBETU = 1

	Do 

		'// ���w���ꗗ�擾
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKUSEKI_NO AS GAKUSEKI, "
		w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI.T11_SIMEI AS SIMEI, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_YOUBI_CD, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_JIGEN, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_NUM,"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO  AS GAKUSEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN ,T11_GAKUSEKI ,T13_GAKU_NEN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI.T11_GAKUSEI_NO = T13_GAKU_NEN.T13_GAKUSEI_NO AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO = T23_DAIGAE_JIKAN.T23_NENDO AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKUSEKI_NO = T13_GAKU_NEN.T13_GAKUSEKI_NO AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_NENDO=" & cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKKI_KBN='" & m_sGakki_Kbn	  & "' AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_KAMOKU='" & m_sKamokuCd & "' AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_KYOKAN='" & m_iKyokanCd & "'"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY GAKUSEKI"
		
		'response.write "w_sSQL =" & w_sSQL & "<BR>"
		
		iRet = gf_GetRecordset(m_Rs_M, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_Get_Data_KOBETU = 99
			Exit Do
		End If
		
		f_Get_Data_KOBETU = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[�@�\]	�ʏ���ƑI�����A��֗��w���ꗗ�擾
'*	[����]	�Ȃ�
'*	[�ߒl]	0:���擾���� 99:���s
'*	[����]	
'********************************************************************************
Function f_Get_TUJO_DaigeRyugak()

	Dim w_sSQL

	On Error Resume Next
	Err.Clear
	
	f_Get_TUJO_DaigeRyugak = 1

	Do 

		'//��֎擾�׸�
		m_bDaigae = True

		'// ���w���ꗗ�擾
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_NENDO, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKKI_KBN, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKUSEKI_NO AS GAKUSEKI, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_YOUBI_CD, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_JIGEN, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKUNEN, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_CLASS, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_KAMOKU, "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_KYOKAN, "
		w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI.T11_SIMEI AS SIMEI, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_NUM,"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO  AS GAKUSEI"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN,"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN,"
		w_sSQL = w_sSQL & vbCrLf & "  T11_GAKUSEKI"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_NENDO = T13_GAKU_NEN.T13_NENDO AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKUNEN = T13_GAKU_NEN.T13_GAKUNEN AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_CLASS = T13_GAKU_NEN.T13_CLASS AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKUSEKI_NO = T13_GAKU_NEN.T13_GAKUSEKI_NO  AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO = T11_GAKUSEKI.T11_GAKUSEI_NO AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_NENDO="		& cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_GAKKI_KBN='" & m_sGakki_Kbn		& "' AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_KAMOKU='"	& m_sKamokuCd		& "' AND "
		w_sSQL = w_sSQL & vbCrLf & "  T23_DAIGAE_JIKAN.T23_KYOKAN='"	& m_iKyokanCd		& "'"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY GAKUSEKI"
		
		iRet = gf_GetRecordset(m_Rs_D, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_Get_TUJO_DaigeRyugak = 99
			Exit Do
		End If

		f_Get_TUJO_DaigeRyugak = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[�@�\]	�ړ�����̏ꍇ�ړ��󋵂̎擾
'*	[����]	p_Gakusei_No:�w��NO
'*			p_Date		:���Ǝ��{��
'*	[�ߒl]	0:���擾���� 99:���s
'*	[����]	
'********************************************************************************
Function f_Get_IdouInfo(p_Gakusei_No,p_Date)

	Dim w_sSQL
	Dim w_Rs
	Dim w_IdoFlg
	Dim w_sKubunName

	On Error Resume Next
	Err.Clear

	w_IdoFlg = False

	Do

		'// ���׃f�[�^
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_1, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_1, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_2, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_2, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_3, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_3, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_4, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_4, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_5, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_5, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_6, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_6, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_7, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_7, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_KBN_8, "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_BI_8"
		w_sSQL = w_sSQL & vbCrLf & " FROM T13_GAKU_NEN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO=" & cint(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEI_NO='" & p_Gakusei_No & "' AND"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_NUM>0"

		iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			Exit Do
		End If

		If w_Rs.EOF = false Then

			i = 1
			Do Until i> cint(C_IDO_MAX_CNT)    '//C_IDO_MAX_CNT = 8�c�ő�ړ���

				If gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) = "" Then
					Exit Do
				End If

				If gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) > p_Date  Then
					Exit Do
				End If
				i = i + 1
			Loop

			If i = 1 then
				'//�ŏ��̈ړ��������Ɠ���薢���̏ꍇ�A���Ɠ��Ɉړ���Ԃł͂Ȃ�
				w_sKubunName = ""
			Else

				Select Case Trim(w_Rs("T13_IDOU_KBN_" & i-1))
					Case cstr(C_IDO_FUKUGAKU),cstr(C_IDO_TEI_KAIJO)  '//C_IDO_FUKUGAKU=3:���w�AC_IDO_TEI_KAIJO=5:��w����
						w_sKubunName = ""
					Case Else
						'//�ړ����R���擾(�敪�}�X�^�A�啪��=C_IDO)
						w_bRet = gf_GetKubunName_R(C_IDO,Trim(w_Rs("T13_IDOU_KBN_" & i-1)),m_iSyoriNen,w_sKubunName)
						If w_bRet<> True Then
							Exit Do
						End If
				End Select
			End If

		End If

		Exit Do
	Loop

	f_Get_IdouInfo = w_sKubunName

	Call gf_closeObject(w_Rs)

	Err.Clear

End Function

'********************************************************************************
'*	[�@�\]	���ȕʏo���f�[�^���擾
'*	[����]	�Ȃ�
'*	[�ߒl]	0:���擾���� 99:���s
'*	[����]	
'********************************************************************************
Function f_Get_AbsInfo()

	Dim w_sSQL

	On Error Resume Next
	Err.Clear
	
	f_Get_AbsInfo = 1

	Do 

		'//���͈̔͂��Z�b�g
		Call f_GetTukiRange(w_sSDate,w_sEDate)

		'// �o���f�[�^
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_HIDUKE, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_YOUBI_CD, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_GAKUSEKI_NO, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_JIGEN, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_SYUKKETU_KBN, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_JIMU_FLG, "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI_R"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU,M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_NENDO = M01_KUBUN.M01_NENDO(+) AND  "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_SYUKKETU_KBN = M01_KUBUN.M01_SYOBUNRUI_CD(+) AND "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_NENDO="	 & cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T21_HIDUKE>='"  & w_sSDate & "' AND "
		w_sSQL = w_sSQL & vbCrLf & "  T21_HIDUKE< '"  & w_sEDate & "' AND "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_KAMOKU='" & m_sKamokuCd		 & "' AND "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_KYOKAN='" & m_iKyokanCd		 & "' AND"
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_DAIBUNRUI_CD=" & C_KESSEKI	'//C_KESSEKI = 19 �啪��(���ȋ敪)
		w_sSQL = w_sSQL & vbCrLf & " GROUP BY "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_HIDUKE, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_YOUBI_CD, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_GAKUSEKI_NO, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_JIGEN, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_SYUKKETU_KBN, "
		w_sSQL = w_sSQL & vbCrLf & "  T21_SYUKKETU.T21_JIMU_FLG, "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI_R"
		
		iRet = gf_GetRecordset(m_Rs_G, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_Get_AbsInfo = 99
			Exit Do
		End If

		f_Get_AbsInfo = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*	[�@�\]	�l�̍s���o���f�[�^��Ԃ�
'*	[����]	p_Date		:���t
'*			p_Gakuseki	:�w��
'*			p_Jigen 	:����
'*	[�ߒl]	p_sSyuketu	:�o������(�f�[�^�Ȃ��̏ꍇ��0(�o��)��Ԃ�)
'*			p_sSyuketu_R:�o������
'*			p_bJim		:True=�������� False=���ƒS����������
'*	[����]	
'********************************************************************************
Function f_Get_Syuketu(p_Date,p_Gakuseki,p_Jigen,p_bJim,p_sSyuketu,p_sSyuketu_R)

	Dim w_sSyuketu

	On Error Resume Next
	Err.Clear

	p_sSyuketu = ""
	p_sSyuketu_R = "�@"
	p_bJim = False
	Do
		If m_Rs_G.EOF = False Then
			m_Rs_G.MoveFirst
			Do Until m_Rs_G.EOF 
				If p_Date = m_Rs_G("T21_HIDUKE") Then
					If trim(p_Gakuseki) = trim(m_Rs_G("T21_GAKUSEKI_NO")) Then

						If cstr(replace(p_Jigen,"$",".")) = cstr(m_Rs_G("T21_JIGEN")) Then

							'//�o������
							p_sSyuketu = m_Rs_G("T21_SYUKKETU_KBN")

							If cstr(p_sSyuketu) = cstr(C_KETU_SYUSSEKI) Then
								p_sSyuketu_R = "�@"
							Else
								p_sSyuketu_R = m_Rs_G("M01_SYOBUNRUIMEI_R")
							End If 

							'//�������͂��ꂽ�f�[�^�́A�ύX�s�Ƃ���
							'//�����׸�(0:���� 1:����)
							If cstr(gf_SetNull2String(m_Rs_G("T21_JIMU_FLG"))) = cstr(C_JIMU_FLG_JIMU) then
								p_bJim = True
							End If

							Exit Do
						End If
					End If
				End If
				m_Rs_G.MoveNext
			Loop
			m_Rs_G.MoveFirst
		End If

		Exit Do

	Loop

	Err.Clear

End Function

'********************************************************************************
'*	[�@�\]	��ԋ߂������̎����敪���擾
'*	[����]	�Ȃ�
'*	[�ߒl]	p_iSikenKbn �����敪
'*	[����]	
'********************************************************************************
Function f_GetSikenKbn(p_iSikenKbn)
	Dim w_sSQL
	Dim w_Rs
	Dim w_iRet

	Dim w_sDate

	On Error Resume Next
	Err.Clear
	
	f_GetSikenKbn = 1
	p_iSikenKbn = ""

	Do 
		'2001/12/17 Add >
		if Cint(m_sTuki) < 4 then
			w_sDate = gf_YYYY_MM_DD((Cint(m_iSyoriNen) + 1) & "/" & m_sTuki & "/01","/")
		Else
			w_sDate = gf_YYYY_MM_DD(m_iSyoriNen & "/" & m_sTuki & "/01","/")
		end if
		if m_sKouki_Start > w_sDate then
			w_sDate = m_sKouki_Start
		End if

		'//�����Ǘ��}�X�^��莎���敪���擾
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI.T24_SIKEN_KBN "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  T24_SIKEN_NITTEI.T24_NENDO=" & m_iSyoriNen
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_SIKEN_CD='0' "
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_GAKUNEN=" & m_sGakunen
		w_sSQL = w_sSQL & vbCrLf & "  AND T24_SIKEN_NITTEI.T24_JISSI_SYURYO>='" & w_sDate & "'"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY T24_SIKEN_NITTEI.T24_SIKEN_KBN"

		iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_GetSikenKbn = 99
			Exit Do
		End If

		'//�߂�l���
		If w_Rs.EOF = False Then
			p_iSikenKbn = w_Rs("T24_SIKEN_KBN")
		End If

		'//�f�[�^���擾�ł��Ȃ��Ƃ�(���������)�́A0�ɂ���
		If p_iSikenKbn = "" Then
			p_iSikenKbn = 0
		End If

		f_GetSikenKbn = 0

		Exit Do
	Loop

	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*	[�@�\]	�o���݌v�f�[�^���擾
'*	[����]	�Ȃ�
'*	[�ߒl]	0:���擾���� 99:���s
'*	[����]	
'********************************************************************************
Function f_Get_AbsInfo_RuiKei()

	Dim w_sSQL
	Dim rs
	Dim w_GakusekiNo
	
	On Error Resume Next
	Err.Clear
	
	f_Get_AbsInfo_RuiKei = 1

	Do 
		
		'//��ԋ߂������̎����敪���擾����
		w_iRet = f_GetSikenKbn(m_iSikenKbn)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			f_Get_AbsInfo_RuiKei = 99
			Exit Do
		End If
		
		'//��������ȍ~�̏ꍇ
		If cint(m_iSikenKbn) = 0 Then
			'w_iSiken = 4
			w_iSiken = 5
		Else
			w_iSiken = m_iSikenKbn
		End If
		
		'//�ŏ��̐��k�̊w�Дԍ����擾
		if not m_Rs_M.EOF then
			w_GakusekiNo = m_Rs_M("GAKUSEKI")
			m_Rs_M.movefirst
		end if
		
		'//�O�̎����Ǝ��̎����Ԃ̊J�n���A�I�������擾
		w_bRtn = gf_GetStartEnd("kks",m_iSyoriNen,m_sSyubetu,cint(w_iSiken),m_sGakunen,m_sClassNo,m_sKamokuCd,w_sKaisibi,w_sSyuryobi,m_iShikenInsertType)
		If w_bRtn <> True Then
			Exit Function
		End If
		
		'// �o���f�[�^
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_GAKUSEKI_NO, "
		w_sSQL = w_sSQL & vbCrLf & "    Count(A.T21_SYUKKETU_KBN) AS KAISU,"
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_SYUKKETU_KBN"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "    T21_SYUKKETU A "
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_NENDO="	& cInt(m_iSyoriNen) & " AND "
		
		w_sSQL = w_sSQL & vbCrLf & "    to_date(A.T21_HIDUKE,'YYYY/MM/DD') >= '" & w_sKaisibi & "' AND "
		w_sSQL = w_sSQL & vbCrLf & "    to_date(A.T21_HIDUKE,'YYYY/MM/DD') <=  '" & w_sSyuryobi & "' AND "
		
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_KAMOKU='" & m_sKamokuCd & "' AND "
		'w_sSQL = w_sSQL & vbCrLf & "    A.T21_KYOKAN='" & m_iKyokanCd & "' AND"
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_SYUKKETU_KBN IN (" & C_KETU_KEKKA & "," & C_KETU_TIKOKU & "," & C_KETU_SOTAI & "," & C_KETU_KEKKA_1 & ")"
		w_sSQL = w_sSQL & vbCrLf & " GROUP BY "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_GAKUSEKI_NO, "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_SYUKKETU_KBN"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T21_GAKUSEKI_NO"
		
		If gf_GetRecordset(rs, w_sSQL) <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_Get_AbsInfo_RuiKei = 99
			Exit Do
		End If

		If rs.EOF = false  Then
			m_iRuiKeiCnt = gf_GetRsCount(rs) - 1
			ReDim Preserve m_AryRuiKei(2,m_iRuiKeiCnt)
			
			'//������
			For i=0 to m_iRuiKeiCnt
				For j=0 to 2
					m_AryRuiKei(j,i)=0
				Next
			Next


			i = 0
			Do Until rs.EOF
				If w_GakuNo <> trim(rs("T21_GAKUSEKI_NO")) Then
					If w_GakuNo <> "" Then
						i = i + 1
					End If
					w_GakuNo = trim(rs("T21_GAKUSEKI_NO"))
					m_AryRuiKei(0,i) = w_GakuNo
				End If

				Select case cstr(rs("T21_SYUKKETU_KBN"))
					case cstr(C_KETU_KEKKA) 	'//���ې�
						m_AryRuiKei(1,i) = m_AryRuiKei(1,i) + cint(rs("KAISU")) * m_iTani
					case cstr(C_KETU_TIKOKU)	'//�x����
						m_AryRuiKei(2,i) = m_AryRuiKei(2,i) + cint(rs("KAISU"))
					case cstr(C_KETU_SOTAI)		'//���ސ�
						m_AryRuiKei(2,i) = m_AryRuiKei(2,i) + cint(rs("KAISU"))
					case cstr(C_KETU_KEKKA_1) 	'//���ې��i�P���ە��j
						m_AryRuiKei(1,i) = m_AryRuiKei(1,i) + cint(rs("KAISU"))
				End Select
				
				If w_GakuNo <> trim(rs("T21_GAKUSEKI_NO")) Then
					i = i + 1
					m_AryRuiKei(0,i) = trim(rs("T21_GAKUSEKI_NO"))
				End If
				
				rs.MoveNext
			Loop
		End If
		
		f_Get_AbsInfo_RuiKei = 0
		Exit Do
	Loop

	Call gf_closeObject(rs)

End Function



'********************************************************************************
'*	[�@�\]	�o�����v�f�[�^���擾
'*	[����]	�Ȃ�
'*	[�ߒl]	0:���擾���� 99:���s
'*	[����]	
'********************************************************************************
Function f_Get_AbsInfo_TukiKei()

	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear
	
	f_Get_AbsInfo_TukiKei = 1

	Do 
		'//���͈̔͂��Z�b�g
		Call f_GetTukiRange(w_sSDate,w_sEDate)
		
		'// �o���f�[�^
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_GAKUSEKI_NO, "
		w_sSQL = w_sSQL & vbCrLf & "    Count(A.T21_SYUKKETU_KBN) AS CNT,"
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_SYUKKETU_KBN"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "    T21_SYUKKETU A"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_NENDO="	 & cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_HIDUKE >= '" & w_sSDate	 & "' AND"
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_HIDUKE < '"  & w_sEDate	 & "' AND"
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_KAMOKU='"    & m_sKamokuCd & "' AND"
		'w_sSQL = w_sSQL & vbCrLf & "    A.T21_KYOKAN='"    & m_iKyokanCd & "' AND"
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_SYUKKETU_KBN IN ('" & cstr(C_KETU_KEKKA) & "','" & cstr(C_KETU_TIKOKU) & "','" & cstr(C_KETU_SOTAI) & "','" & cstr(C_KETU_KEKKA_1) & "')"
		w_sSQL = w_sSQL & vbCrLf & " GROUP BY "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_GAKUSEKI_NO, "
		w_sSQL = w_sSQL & vbCrLf & "    A.T21_SYUKKETU_KBN"
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY A.T21_GAKUSEKI_NO"
		
		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			msMsg = Err.description
			f_Get_AbsInfo_TukiKei = 99
			Exit Do
		End If

		If rs.EOF= false  Then

			'//ں��ރJ�E���g�擾
			m_iTukiKeiCnt = gf_GetRsCount(rs) - 1

			ReDim Preserve m_AryTukiKei(2,cInt(m_iTukiKeiCnt))

			'//������
			For j=0 to 2
				For i=0 to m_iTukiKeiCnt
					m_AryTukiKei(j,i)=0
				Next
			Next

			i = 0
			Do Until rs.EOF

				If w_GakuNo <> trim(rs("T21_GAKUSEKI_NO")) Then
					If w_GakuNo <> "" Then
						i = i + 1
					End If
					w_GakuNo = trim(rs("T21_GAKUSEKI_NO"))
					m_AryTukiKei(0,i) = w_GakuNo
				End If

				Select case cstr(rs("T21_SYUKKETU_KBN"))
					case cstr(C_KETU_KEKKA) 	'//���ې�
						m_AryTukiKei(1,i) = m_AryTukiKei(1,i) + cint(rs("CNT")) * m_iTani
					case cstr(C_KETU_TIKOKU)	'//�x����
						m_AryTukiKei(2,i) = m_AryTukiKei(2,i) + cint(rs("CNT"))
					case cstr(C_KETU_SOTAI)		'//���ސ�
						m_AryTukiKei(2,i) = m_AryTukiKei(2,i) + cint(rs("CNT"))
					case cstr(C_KETU_KEKKA_1) 	'//���ې��i�P���ە��j
						m_AryTukiKei(1,i) = m_AryTukiKei(1,i) + cint(rs("CNT"))
				End Select
				
				rs.MoveNext
			Loop
			
		End If
		
		f_Get_AbsInfo_TukiKei = 0
		Exit Do
	Loop
	
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*	[�@�\]	�o���敪�Ɩ��̂��擾
'*	[����]	�Ȃ�
'*	[�ߒl]	0:���擾���� 99:���s
'*	[����]	�o�����͂�JAVASCRIPT�쐬
'********************************************************************************
Function f_Get_SYUKETU_KBN(p_MaxNo)

	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear
	
	f_Get_SYUKETU_KBN = 1

	Do 
		'// ���׃f�[�^
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD, "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUIMEI_R"
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_NENDO=" & cInt(m_iSyoriNen) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_DAIBUNRUI_CD=" & cint(C_KESSEKI) & " AND "
		'//C_SYOBUNRUICD_IPPAN = 4	'//���ȋ敪(0:�o��,1:����,2:�x��,3:����,4:����,�c)
		w_sSQL = w_sSQL & vbCrLf & "  M01_KUBUN.M01_SYOBUNRUI_CD IN ("
		w_sSQL = w_sSQL & vbCrLf & " " & C_KETU_SYUSSEKI & " "
		w_sSQL = w_sSQL & vbCrLf & " ," & C_KETU_KEKKA & " "
		w_sSQL = w_sSQL & vbCrLf & " ," & C_KETU_TIKOKU & " "
		w_sSQL = w_sSQL & vbCrLf & " ," & C_KETU_SOTAI & " "
		
		If m_iTani > 1 then '�P�������P���ۂ��傫���ꍇ�́A�P���ۂ̋敪���o��
			w_sSQL = w_sSQL & vbCrLf & " ," & C_KETU_KEKKA_1 & " "
		End If

		w_sSQL = w_sSQL & vbCrLf & " ) "
		w_sSQL = w_sSQL & vbCrLf & " ORDER BY M01_KUBUN.M01_SYOBUNRUI_CD"

		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_Get_SYUKETU_KBN = 99
			Exit Do
		End If

		i=0
		If rs.EOF = True Then
			response.write ("var ary = new Array(0);")
			response.write ("var aryCD = new Array(0);")

			response.write ("aryCD[0] = '';")
			response.write ("ary[0] = '';")
		Else

			'//ں��ރJ�E���g�擾
			w_iCnt = gf_GetRsCount(rs) - 1
			response.write ("var aryCD = new Array(" & w_iCnt & ");") & vbCrLf
			response.write ("var ary = new Array(" & w_iCnt & ");") & vbCrLf

			Do Until rs.EOF
'			 response.write ("var ary[" & i & "] = new Array(1);") & vbCrLf
				If i = 0 Then
					response.write ("aryCD[0] = 0;") & vbCrLf
					response.write ("ary[0] = '�@';") & vbCrLf
				Else
'					 response.write ("ary[" & rs("M01_SYOBUNRUI_CD") &	"] = '" & rs("M01_SYOBUNRUIMEI_R") & "';") & vbCrLf

					'���̕��ɏC�� 2001/10/29
					'aryCD=�����ރR�[�h
					'ary=�����ޗ���
					response.write ("aryCD[" & i &	"] = '" & rs("M01_SYOBUNRUI_CD") & "';") & vbCrLf
					response.write ("ary[" & i &  "] = '" & rs("M01_SYOBUNRUIMEI_R") & "';") & vbCrLf
				End If

				i=i+1
				rs.MoveNext
			Loop

		End If

		p_MaxNo = w_iCnt

		f_Get_SYUKETU_KBN = 0
		Exit Do
	Loop

	Call gf_closeObject(rs)
	Err.Clear

End Function

'********************************************************************************
'*	[�@�\]	�z�񏉊���
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub s_AryInit(p_iRecCount)

	For j=0 to 4
		For i=0 to p_iRecCount
			m_AryHead(j,i) = ""
		Next
	Next

End Sub

'********************************************************************************
'*	[�@�\]	���̌����������쐬(7���c�@"MONTH>=2001/07/01 AND MONTH<2001/08/01" �Ƃ��Ďg�p)
'*	[����]	�Ȃ�
'*	[�ߒl]	p_sSDate
'*			p_sEDate
'*	[����]	
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
'*	[�@�\]	HTML���o��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub showPage()
Dim w_sIduoRiyu
Dim w_bJimInsData	'//���������׸�
Dim w_bNoSelect		'//���Ȕ�I���׸�
Dim w_bNoChange		'//�ύX�s���׸�
Dim w_bEndFLG		'//���ׂĕύX�s�̏ꍇTRUE

	On Error Resume Next
	Err.Clear

	w_bEndFLG = True
%>
	<html>
	<head>
	<title>�s���p�o������</title>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	<!--#include file="../../Common/jsCommon.htm"-->

	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--

	//************************************************************
	//	[�@�\]	�y�[�W���[�h������
	//	[����]
	//	[�ߒl]
	//	[����]
	//************************************************************
	function window_onload() {

		//�X�N���[����������
		parent.init();
	if(location.href.indexOf('#')==-1)
 	{
		//�w�b�_����\��submit
		//document.frm.target = "middle";
		document.frm.target = "topFrame";
		document.frm.action = "kks0110_middle.asp"
		document.frm.submit();
	}

		return;

	}

	//************************************************************
	//	[�@�\]	�o������
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//************************************************************
	function chg(chgInp) {

		no = 0;
		<%
		'//�o���敪���擾
		Call f_Get_SYUKETU_KBN(w_MaxNo)
		%>

		str = chgInp.value;
		for(i=0; i<<%=w_MaxNo+1%>; i++){
			if (ary[i]==str){
				break;
			}
		};

		no = i + 1;
		if (no > <%=w_MaxNo%>) no = 0;
		chgInp.value = ary[no];

		//�B���t�B�[���h�Ƀf�[�^���Z�b�g
		var obj=eval("document.frm.hid"+chgInp.name);
		obj.value=aryCD[no];
		return;
	}
	//************************************************************
	//	[�@�\]	�o�^�{�^���������ꂽ�Ƃ�
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_Touroku(){

		if (!confirm("<%=C_TOUROKU_KAKUNIN%>")) {
		   return ;
		}

		//�w�b�_���󔒕\��
		parent.topFrame.document.location.href="white.asp?txtMsg=<%=Server.URLEncode("�o�^���Ă��܂��E�E�E�E�@�@���΂炭���҂���������")%>"

		//���X�g����submit
		document.frm.target = "main";
		document.frm.action = "./kks0110_edt.asp"
		document.frm.submit();
		return;
	}

	//************************************************************
	//	[�@�\]	�L�����Z���{�^���������ꂽ�Ƃ�
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_Cancel(){
		//�����y�[�W��\��
		parent.document.location.href="default.asp"
	}

	//-->
	</SCRIPT>

	</head>
	<body LANGUAGE=javascript onload="window_onload()">
	<form name="frm" method="post" onClick="return false;">

	<center>
	<%Do %>
		<%If m_iRsCnt < 0 Then%>
			<br><br>
			<span class="msg">���Ɠ���������܂���</span>
			<%Exit Do%>
		<%End If%>

		<%If m_Rs_M.EOF Then%>
			<br><br>
			<span class="msg">���C�f�[�^������܂���</span>
			<%Exit Do%>
		<%End If%>

		<%'//�������B�����ڂɾ��%>
		<%for i = 0 to m_iRsCnt%>
			<input type="hidden" name="JIKANWARI" value="<%=m_AryHead(4,i) & "_" & m_AryHead(3,i)%>">
		<%Next%>

		<table >
		<tr>
			<td align="center" valign="top">
			<table class="hyo"	border="1" >

			<%

			Dim w_sSentaku

			'//���ד��͗��\��
			If m_Rs_M.EOF = False Then

				Do Until m_Rs_M.EOF

					'//���ټ�Ă̸׽���Z�b�g
					Call gs_cellPtn(w_Class) 
					%>
					<tr>
						<td class="<%=w_Class%>" align="center" height="28" nowrap width="50"><%=m_Rs_M("GAKUSEKI")%>
							<input type="hidden" name="GAKUSEI" value=<%=m_Rs_M("GAKUSEI")%>>
						</td>
						<td class="<%=w_Class%>" align="left" height="28" nowrap width="150"><%=m_Rs_M("SIMEI")%></td>
					<%
					'//���Ȕ�I���׸�
					w_bNoSelect = False

					'//�ʏ���ƑI����(���ʎ��ƁA�ʎ��ƈȊO�̂�)
					If m_sSyubetu = "TUJO" then

						'//�K�I�敪��2�̉Ȗ�(�I���Ȗ�)�̂Ƃ��A�I���ۂ𔻕�(�I������=1�A�I�����Ȃ�=0)�I�����Ȃ��ꍇ�́A�o�����͕s��
						If cstr(m_sHissenKbn) = cstr(C_HISSEN_SEN) Then

							'//�I���ۂ𔻕�
							If cstr(gf_SetNull2Zero(m_Rs_M("T16_SELECT_FLG"))) = cstr(gf_SetNull2Zero(C_SENTAKU_NO)) Then
								w_bNoSelect = True
							End If
						Else
							'//�u���Ȗ��׸ނ�1(0:�ʏ�,1:�u����,2:�u����)�̏ꍇ�́A�o�����͕s��
							If cstr(gf_SetNull2Zero(m_Rs_M("T16_OKIKAE_FLG"))) = cstr(gf_SetNull2Zero(C_TIKAN_KAMOKU_MOTO)) Then
								w_bNoSelect = True
							End If

	
							'���ٕʉȖ�
							If cstr(trim(m_sLevelFlg)) = cstr(C_LEVEL_YES) Then

								'//���ٕʉȖ� ���ٕʋ����R�[�h��NULL�Ȃ�A�o�����͕s�� ito
								If isNull(m_Rs_M("T16_LEVEL_KYOUKAN")) = True Then
									w_bNoSelect = True
								Else
									If m_Rs_M("T16_LEVEL_KYOUKAN") <> m_iKyokanCd Then
										w_bNoSelect = True
									End If
								End If
							End If
					
						End If

					End If

					For i = 0 to m_iRsCnt

						'//�ύX�s���׸ޏ�����
						w_bNoChange = False

						'//���k�����Ȃ�I�����ĂȂ��ꍇ
						If w_bNoSelect = True Then
							w_bNoChange = True
						End If

						'//�ړ��󋵂̍l��(T13_IDOU_NUM��1�ȏ�̏ꍇ�͈ړ��󋵂𔻕ʂ���)�ړ����̏ꍇ�́A�o�����͕s��
						w_sIduoRiyu = ""
						If gf_SetNull2Zero(m_Rs_M("T13_IDOU_NUM")) > 0 Then 
							w_sIduoRiyu = f_Get_IdouInfo(m_Rs_M("GAKUSEI"),m_AryHead(4,i))
						End If

						'//�ړ����łȂ��ꍇ�o���f�[�^���擾
						If Trim(w_sIduoRiyu) = "" Then
							'//���t�A�w��NO(5��),����,��������(���������׸ނ�1�̏ꍇ�͓��͕s��)���A�o����񓙂��擾
							Call f_Get_Syuketu(m_AryHead(4,i),m_Rs_M("GAKUSEKI"),m_AryHead(3,i),w_bJimInsData,w_Syuketu,w_Syuketu_R)
							'//���������׸ނ�1�̏ꍇ�͕ύX�s�Ƃ���
							If w_bJimInsData = True Then
								w_bNoChange = True
							End If
						End If
						'//�ړ���(���͕s��)
						If w_sIduoRiyu <> "" Then%>
							<td align="center" class="NOCHANGE" height="28" nowrap width="30" ><%=w_sIduoRiyu%><br>
							<input type="hidden" name="hidKBN<%=m_Rs_M("GAKUSEI") & "_" & replace(m_AryHead(4,i),"/","") & "_" & m_AryHead(3,i)%>" size="2" value="---"></td>
						<%
						'//�������� OR �I�����Ă��Ȃ�(���͕s��) ito
						Else

							If w_bNoChange = True Then%>
								<td align="center" class="NOCHANGE" height="28" nowrap	width="30" ><%=w_Syuketu_R%><br>
								<input type="hidden" name="hidKBN<%=m_Rs_M("GAKUSEI") & "_" & replace(m_AryHead(4,i),"/","") & "_" & m_AryHead(3,i)%>" size="2" value="---"></td>
							<%
							'//�ύX�E���͉f�[�^
							Else%>
								
								<% '�ύX�\���Ԃ̏ꍇ
									w_bEndFLG = False %>
									<td class="<%=w_Class%>" align="center" width="30" height="28" nowrap>
									<input type="button" class="<%=w_Class%>" name="KBN<%=m_Rs_M("GAKUSEI") & "_" & replace(m_AryHead(4,i),"/","") & "_" & m_AryHead(3,i)%>" size="2" value="<%=w_Syuketu_R%>"	style="border-style:none" style="text-align:center" tabindex="-1" onclick="return chg(this)">
									<input type="hidden" name="hidKBN<%=m_Rs_M("GAKUSEI") & "_" & replace(m_AryHead(4,i),"/","") & "_" & m_AryHead(3,i)%>" size="2" value="<%=w_Syuketu%>"></td>

							
							<%End If%>
					 <%End If%>

					<%Next%>
					</tr>
					<%m_Rs_M.MoveNext%>
				<%Loop%>

					<%m_Rs_M.MoveFirst%>
			<%End If%>

			<%

			'======================================
			'//��֗��w���̒ǉ�
			If m_bDaigae = True Then
				
				If m_Rs_D.EOF = false Then

					Do Until m_Rs_D.EOF
					'//���ټ�Ă̸׽���Z�b�g
					Call gs_cellPtn(w_Class) 

						%>
						<tr>
							<td class="<%=w_Class%>" align="center" height="28" nowrap><%=m_Rs_D("GAKUSEKI")%>
								<input type="hidden" name="GAKUSEI" value=<%=m_Rs_D("GAKUSEI")%>>
							</td>
							<td class="<%=w_Class%>" align="center" height="28" nowrap><%=m_Rs_D("SIMEI")%></td>
						<%
						for i = 0 to m_iRsCnt
							w_bNoChange = False

							'//���͋��E�񋖉̔���
							'//�ړ��󋵂̍l��(T13_IDOU_NUM��1�ȏ�̏ꍇ�͈ړ��󋵂𔻕ʂ���)�ړ����̏ꍇ�́A�o�����͕s��
							w_sIduoRiyu = ""
							If gf_SetNull2Zero(m_Rs_D("T13_IDOU_NUM")) > 0 Then 
								w_sIduoRiyu = f_Get_IdouInfo(m_Rs_M("GAKUSEI"),m_AryHead(4,i))
							End If

							'//���t�A�w��NO(5��),����,�S�C����(�S�C�����׸ނ�1�̏ꍇ�͒S�C�̂ݓ��͉�)
							Call f_Get_Syuketu(m_AryHead(4,i),m_Rs_D("GAKUSEKI"),m_AryHead(3,i),w_bJimInsData,w_Syuketu,w_Syuketu_R)

							'//���������׸ނ�1�̏ꍇ�͕ύX�s�Ƃ���
							If w_bJimInsData = True Then
								w_bNoChange = w_bJimInsData
							End If

							'//���Ԋ��̎����Ƒ�֎��Ԋ��̎�������v���Ă��邩
							'//��v���ĂȂ��ꍇ�͗��w���͂��̎�����I�����Ă��Ȃ��Ƃ݂Ȃ��A�o�����͂�s�Ƃ���
							If w_bNoChange = False Then
								If cstr(replace(m_AryHead(3,i),"$",".")) <> cstr(m_Rs_D("T23_JIGEN")) Then
									w_bNoChange = True
								End If
							End If

							'//�ړ���(���͕s��)
							If w_sIduoRiyu <> "" Then%>
								<td align="center" class=NOCHANGE height="28" nowrap><%=w_sIduoRiyu%><br></td>
								<input type="hidden" name="hidKBN<%=m_Rs_D("GAKUSEI") & "_" & replace(m_AryHead(4,i),"/","") & "_" & m_AryHead(3,i)%>" size="2" value="---"></td>
							<%
							'//�������� OR �I�����Ă��Ȃ�(���͕s��)
							ElseIf w_bNoChange = True Then%>
								<td align="center" class=NOCHANGE height="28" nowrap><%=w_Syuketu_R%><br></td>
								<input type="hidden" name="hidKBN<%=m_Rs_D("GAKUSEI") & "_" & replace(m_AryHead(4,i),"/","") & "_" & m_AryHead(3,i)%>" size="2" value="---"></td>
							<%Else%>
								<% '�ύX�\���Ԃ̏ꍇ %>
								<% w_bEndFLG = False %>
									<td class="<%=w_Class%>" align="center"  width="30" height="28"  nowrap>
									<input type="button" class="<%=w_Class%>" name="KBN<%=m_Rs_D("GAKUSEI") & "_" & replace(m_AryHead(4,i),"/","") & "_" & m_AryHead(3,i)%>" size="2" value="<%=w_Syuketu_R%>" style="border-style:none" style="text-align:center" tabindex="-1" onclick="return chg(this)">
									<input type="hidden" name="hidKBN<%=m_Rs_D("GAKUSEI") & "_" & replace(m_AryHead(4,i),"/","") & "_" & m_AryHead(3,i)%>" size="2" value="<%=w_Syuketu%>"></td>
								<% '�ύX�\���ԂłȂ��ꍇ %>


							<%End If

						Next

						m_Rs_D.MoveNext
					Loop
					m_Rs_D.MoveFirst
				End If
				w_Class=""
			End If
			'======================================
			%>

			</table>

		</td>
		<td width="10"><br></td>
		<td align="center" valign="top" width="120"  nowrap>

			<!--���E�w���̌��ȋy�ђx�����݌v-->
			<table	class="hyo" border="1" width="120">

			<%If m_Rs_M.EOF = False Then
				w_Class = ""
				Do Until m_Rs_M.EOF
					'//���ټ�Ă̸׽���Z�b�g
					Call gs_cellPtn(w_Class) 
					%>
					<tr>
					<%
					w_sGakusekiNo = m_Rs_M("GAKUSEKI")

					'//������
					w_TukiTikoku = 0   '//���x��
					w_TukiKekka  = 0   '//������
					w_RuiTikoku  = 0   '//�݌v�x��
					w_RuiKekka	 = 0   '//�݌v����
					
					'//���v���擾
					
					For i=0 To cInt(m_iTukiKeiCnt)
						If w_sGakusekiNo=m_AryTukiKei(0,i) Then
							w_TukiKekka  = m_AryTukiKei(1,i)	'//���ې�
							w_TukiTikoku = m_AryTukiKei(2,i)	'//�x����
							Exit For
						End If
					Next

					'//�݌v���擾
					For i=0 To m_iRuiKeiCnt
						If w_sGakusekiNo=m_AryRuiKei(0,i) Then

							w_RuiKekka	= m_AryRuiKei(1,i)	'//���ې�
							w_RuiTikoku = m_AryRuiKei(2,i)	'//�x����
							
							'//�ݐς̕\���@���ݐς̏ꍇ
							'//Public Const C_K_KEKKA_RUISEKI_SIKEN = 0    '������
							'//Public Const C_K_KEKKA_RUISEKI_KEI = 1	   '�ݐ�
							
							If cint(m_iSyubetu) = cint(C_K_KEKKA_RUISEKI_KEI) Then
								
								'//�������������̏ꍇ
								'If m_iSikenKbn = 0 Then
								'	w_iKbn = 4
								'Else
								'	w_iKbn = cint(m_iSikenKbn) - 1
								'End If
								
								'//�O��̎������̏o�����擾����
								w_sGakusekiNo = m_Rs_M("GAKUSEI")		'2001/12/17 Add
								'm_sSyubetu
								'Call gf_GetKekaChi(m_iSyoriNen,m_iShikenInsertType,m_sKamokuCd,w_sGakusekiNo,p_iKekka,p_iChikoku,w_iKekkaGai)
								Call gf_GetKekaChi(m_iSyoriNen,m_sSyubetu,m_iShikenInsertType,m_sKamokuCd,w_sGakusekiNo,p_iKekka,p_iChikoku,w_iKekkaGai)
								
								w_RuiKekka = cint(w_RuiKekka) + cint(gf_SetNull2Zero(p_iKekka))
								w_RuiTikoku = cint(w_RuiTikoku) + cint(gf_SetNull2Zero(p_iChikoku))

							End If

							Exit For
						End If
					Next

					%>
						<td class="<%=w_Class%>" align="center" height="28" width="30" nowrap><%=gf_IIF(w_TukiTikoku <> 0, w_TukiTikoku, "�@") %><br></td>
						<td class="<%=w_Class%>" align="center" height="28" width="30" nowrap><%=gf_IIF(w_TukiKekka <> 0, w_TukiKekka, "�@") %><br></td>
						<td class="<%=w_Class%>" align="center" height="28" width="30" nowrap><%=gf_IIF(w_RuiTikoku <> 0, w_RuiTikoku, "�@") %><br></td>
						<td class="<%=w_Class%>" align="center" height="28" width="30" nowrap><%=gf_IIF(w_RuiKekka <> 0,  w_RuiKekka, "�@")  %><br></td>
					</tr>
					<%m_Rs_M.MoveNext%>
				<%Loop%>
			<%End If%>
			<%
			'=================================================
			'//��֗��w���݌v�̒ǉ�
			If m_bDaigae = True Then
				If m_Rs_D.EOF = False Then
					'//���ټ�Ă̸׽���Z�b�g
					Call gs_cellPtn(w_Class) 
					
					Do Until m_Rs_D.EOF
					%>
						<tr>
						<%
						w_sGakusekiNo = m_Rs_D("GAKUSEKI")
						
						'//������
						
						w_TukiTikoku = 0   '//���x��
						w_TukiKekka  = 0   '//������
						w_RuiTikoku  = 0   '//�݌v�x��
						w_RuiKekka	 = 0   '//�݌v����
						
						'//���v���擾
						For i=0 To cInt(m_iTukiKeiCnt)
							If w_sGakusekiNo=m_AryTukiKei(0,i) Then
								w_TukiKekka  = m_AryTukiKei(1,i)	'//���ې�
								w_TukiTikoku = m_AryTukiKei(2,i)	'//�x����
								Exit For
							End If
						Next

						'//�݌v���擾
						For i=0 To m_iRuiKeiCnt
							If w_sGakusekiNo=m_AryRuiKei(0,i) Then
								w_RuiKekka	= m_AryRuiKei(1,i)	'//���ې�
								w_RuiTikoku = m_AryRuiKei(2,i)	'//�x����


								'// �ݐς̕\���@���ݐς̏ꍇ
								'//Public Const C_K_KEKKA_RUISEKI_SIKEN = 0    '������
								'//Public Const C_K_KEKKA_RUISEKI_KEI = 1	   '�ݐ�
								
								If m_iSyubetu = C_K_KEKKA_RUISEKI_KEI Then
									
									'//�������������̏ꍇ
									'If m_iSikenKbn = 0 Then
									'	w_iKbn = 4
									'Else
									'	w_iKbn = cint(m_iSikenKbn) - 1
									'End If
									
									'//�O��̎������̏o�����擾����
									w_sGakusekiNo = m_Rs_D("GAKUSEI")
									'Call gf_GetKekaChi(m_iSyoriNen,w_iKbn,m_sKamokuCd,w_sGakusekiNo,p_iKekka,p_iChikoku,w_iKekkaGai)
									Call gf_GetKekaChi(m_iSyoriNen,m_sSyubetu,m_iShikenInsertType,m_sKamokuCd,w_sGakusekiNo,p_iKekka,p_iChikoku,w_iKekkaGai)
									
									w_RuiKekka = cint(w_RuiKekka) + cint(gf_SetNull2Zero(p_iKekka))
									w_RuiTikoku = cint(w_RuiTikoku) + cint(gf_SetNull2Zero(p_iChikoku))
									
								End If

								Exit For
							End If
						Next
						%>
							<td class="<%=w_Class%>" align="center" height="28" width="30" nowrap><%=gf_IIF(w_TukiTikoku <> 0, w_TukiTikoku, "�@") %><br></td>
							<td class="<%=w_Class%>" align="center" height="28" width="30" nowrap><%=gf_IIF(w_TukiKekka <> 0, w_TukiKekka, "�@") %><br></td>
							<td class="<%=w_Class%>" align="center" height="28" width="30" nowrap><%=gf_IIF(w_RuiTikoku <> 0, w_RuiTikoku, "�@") %><br></td>
							<td class="<%=w_Class%>" align="center" height="28" width="30" nowrap><%=gf_IIF(w_RuiKekka <> 0,  w_RuiKekka, "�@")  %><br></td>
						</tr>
						<%m_Rs_D.MoveNext
					Loop
				End If

			End If
			'=================================================
			%>

			</table>

		</td>
		<tr><td height=10><br></td></tr>
		<tr>
		
		<% if w_bEndFLG = False Then %> 
			<td valign="bottom"  colspan=3 align="center">
				<input class=button type="button" onclick="javascript:f_Touroku();" value="�@�o�@�^�@">
				&nbsp;&nbsp;&nbsp;
				<input class=button type="button" onclick="javascript:f_Cancel();" value="�L�����Z��">
			</td>
		<% else %>
			<td valign="bottom"  colspan=3 align="center">
				<input class=button type="button" onclick="javascript:f_Cancel();" value=" �߁@�� ">
			</td>
		<% end If %>
		
		</tr>
		</table>


		<%Exit Do%>
	<%Loop%>

	<input type="hidden" name="JikanSU"   value="<%=m_iRsCnt%>">
	<input type="hidden" name="Tuki_Zenki_Start" value="<%=m_sZenki_Start%>">
	<input type="hidden" name="Tuki_Kouki_Start" value="<%=m_sKouki_Start%>">
	<input type="hidden" name="Tuki_Kouki_End"	 value="<%=m_sKouki_End%>">
	<INPUT TYPE=HIDDEN NAME="NENDO" 	value = "<%=m_iSyoriNen%>">
	<INPUT TYPE=HIDDEN NAME="KYOKAN_CD" value = "<%=m_iKyokanCd%>">
	<INPUT TYPE=HIDDEN NAME="TUKI"		value = "<%=m_sTuki%>">
	<INPUT TYPE=HIDDEN NAME="GAKKI" 	value = "<%=m_sGakki%>">
	<INPUT TYPE=HIDDEN NAME="GAKUNEN"	value = "<%=m_sGakunen%>">
	<INPUT TYPE=HIDDEN NAME="CLASSNO"	value = "<%=m_sClassNo%>">
	<INPUT TYPE=HIDDEN NAME="KAMOKU_CD" value = "<%=m_sKamokuCd%>">
	<INPUT TYPE=HIDDEN NAME="SYUBETU"	value = "<%=m_sSyubetu%>">
	<INPUT TYPE=HIDDEN NAME="EndFLG"   value = "<%=w_bEndFLG%>">
	<INPUT TYPE=HIDDEN NAME="cboGakunenCd"   value = "<%=request("cboGakunenCd")%>">
	<INPUT TYPE=HIDDEN NAME="cboClassCd"   value = "<%=request("cboClassCd")%>">
	<INPUT TYPE=HIDDEN NAME="txtFromDate"   value = "<%=request("txtFromDate")%>">
	<INPUT TYPE=HIDDEN NAME="txtToDate"   value = "<%=request("txtToDate")%>">
	
	
	<INPUT TYPE=HIDDEN NAME="KAMOKU_NAME" value="<%=Request("KAMOKU_NAME")%>">
	<INPUT TYPE=HIDDEN NAME="CLASS_NAME"  value="<%=Request("CLASS_NAME")%>">

	</form>
	</center>
	</body>
	</html>
<%
End Sub

'********************************************************************************
'*	[�@�\]	��HTML���o��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
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
	//	[�@�\]	�y�[�W���[�h������
	//	[����]
	//	[�ߒl]
	//	[����]
	//************************************************************
	function window_onload() {
	}
	//-->
	</SCRIPT>

	</head>
	<body LANGUAGE=javascript onload="return window_onload()">
	<form name="frm" mothod="post">

	<center>
	<br><br><br>
		<span class="msg"><%=Server.HTMLEncode(p_Msg)%></span>
	</center>

	<input type="hidden" name="txtMsg" value="<%=Server.HTMLEncode(p_Msg)%>">
	</form>
	</body>
	</html>
<%
End Sub
%>
