<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w����񌟍��ڍ�
' ��۸���ID : gak/gak0300/kojin_sita1.asp
' �@      �\: �������ꂽ�w���̏ڍׂ�\������(�l���)
'-------------------------------------------------------------------------
' ��      ��	Session("GAKUSEI_NO")  = �w���ԍ�
'            	Session("HyoujiNendo") = �\���N�x
'           
' ��      ��
' ��      �n
'           
'           
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/02 ��c
' ��      �X: 2001/07/02
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public m_bErrFlg        '�װ�׸�
	Public m_Rs				'ں��޾�ĵ�޼ު��
	Public m_SEIBETU		'����
	Public m_BLOOD			'���t�^
	Public m_RH				'RH
	Public m_HOG_ZOKU		'�ی�ґ���
	Public m_HOS_ZOKU		'�ۏؐl����
	Public m_RYOSEI_KBN		'�����敪
	Public m_RYUNEN_FLG		'�i���敪

	Public m_HyoujiFlg		'�\���׸�
	Public m_KakoRs			'ں��޾�ĵ�޼ު��(�ߋ��׽)
	Public mHyoujiNendo		'�\���N�x

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

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�w����񌟍�����"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""


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
		'//�ߋ��̃N���X���擾
		w_iRet = f_GetDetailKakoClass()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

		'//�\�����ڂ��擾
		w_iRet = f_GetDetailGakunen()
		If w_iRet <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

        '//�����\��
        if m_TxtMode = "" then
            Call showPage()
            Exit Do
        end if

        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\���iϽ�����ƭ��ɖ߂�j
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    '// �I������
    If Not IsNull(m_Rs) Then gf_closeObject(m_Rs)
    Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [�@�\]  �ߋ��̃N���X���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  0:����I��	1:�C�ӂ̃G���[  99:�V�X�e���G���[
'*  [����]  
'********************************************************************************
Function f_GetDetailKakoClass()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetDetailKakoClass = 1

	Do

		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & " 	T13.T13_NENDO, "
		w_sSql = w_sSql & " 	T13.T13_GAKUNEN,  "
		w_sSql = w_sSql & " 	T13.T13_CLASS "
		w_sSql = w_sSql & " FROM T13_GAKU_NEN T13 "
		w_sSql = w_sSql & " WHERE  "
		w_sSql = w_sSql & " 	    T13.T13_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' "
		'w_sSql = w_sSql & " 	AND T13.T13_RYUNEN_FLG <> 1"
		w_sSql = w_sSql & " ORDER BY T13.T13_NENDO DESC "

		iRet = gf_GetRecordset(m_KakoRs, w_sSql)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_GetDetailKakoClass = 99
			Exit Do
		End If

		if m_KakoRs.Eof then
			msMsg = "�w�N���擾���ɃG���[���������܂���"
			f_GetDetailKakoClass = 99
			Exit Do
		End if

		'//����I��
		f_GetDetailKakoClass = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [�@�\]  �N���X���擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �N���X����
'*  [����]  
'********************************************************************************
Function f_GetClass(p_sCLASS,p_iGakunen)

	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClass = ""

	Do 

		'// �N���X���
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "   M05_CLASSMEI	 "
		w_sSQL = w_sSQL & vbCrLf & "  ,M05_CLASSRYAKU	 "
		w_sSQL = w_sSQL & vbCrLf & "  ,M05_TANNIN	"	
		w_sSQL = w_sSQL & vbCrLf & " FROM M05_CLASS"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		'w_sSQL = w_sSQL & vbCrLf & "      M05_NENDO = " & request("selNendo")
		
		If request("selNendo") = "" Then
			w_sSQL = w_sSQL & vbCrLf & "      M05_NENDO = " & mHyoujiNendo
		Else
			w_sSQL = w_sSQL & vbCrLf & "      M05_NENDO = " & request("selNendo")
		End If

		w_sSQL = w_sSQL & vbCrLf & "  AND M05_GAKUNEN =" & p_iGakunen
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASSNO = '" & p_sCLASS & "'"
'response.write w_ssql
		iRet = gf_GetRecordset(rs, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			Exit Do
		End If

		If rs.EOF Then
			Exit Do
		End If

		f_GetClass = gf_HTMLTableSTR(rs("M05_CLASSMEI"))

		Exit Do
	Loop

	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �\�����ڂ��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  0:����I��	1:�C�ӂ̃G���[  99:�V�X�e���G���[
'*  [����]  
'********************************************************************************
Function f_GetDetailGakunen()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	'// �\������N�x�����߂�
	wSelNendo = request("selNendo")
	if gf_IsNull(wSelNendo) then
		mHyoujiNendo = Session("HyoujiNendo")
	Else
		mHyoujiNendo = wSelNendo
	End if

	f_GetDetailGakunen = 1

	Do

		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & "		A.T13_NENDO, "				'�����N�x
		w_sSql = w_sSql & "		A.T13_GAKUSEKI_NO, "		'�w�Дԍ�
		w_sSql = w_sSql & "		A.T13_GAKUNEN, "			'�w�N
		w_sSql = w_sSql & " 	E.M01_SYOBUNRUIMEI, "		'�ݐЋ敪�@��

		w_sSql = w_sSql & " 	E.M01_DAIBUNRUI_CD, "		'�ݐЋ敪�@��
		w_sSql = w_sSql & " 	E.M01_SYOBUNRUI_CD, "		'�ݐЋ敪�@��

		w_sSql = w_sSql & "		A.T13_GAKKA_CD, "			'�w��CD
		w_sSql = w_sSql & " 	B.M02_GAKKAMEI, "			'�w�Ȗ��́@��
		w_sSql = w_sSql & "		A.T13_COURCE_CD, "			'�R�[�XCD�@��
		w_sSql = w_sSql & "		A.T13_CLASS, "				'�N���XCD�@��
		w_sSql = w_sSql & "		A.T13_SYUSEKI_NO1, "		'�o�Ȕԍ��i�w�ȁj
		w_sSql = w_sSql & "		A.T13_SYUSEKI_NO2, "		'�o�Ȕԍ��i�N���X�j
		w_sSql = w_sSql & "		A.T13_RYOSEI_KBN, "			'�����敪�@��
		w_sSql = w_sSql & "		A.T13_RYUNEN_FLG, "			'���N�敪�@��
		w_sSql = w_sSql & "		A.T13_CLUB_1, "				'�N���u�P�@��
		w_sSql = w_sSql & "		A.T13_CLUB_1_NYUBI, "		'�N���u�����P
		w_sSql = w_sSql & "		A.T13_CLUB_2, "				'�N���u�Q�@��
		w_sSql = w_sSql & "		A.T13_CLUB_2_NYUBI, "		'�N���u�����Q
		w_sSql = w_sSql & "		A.T13_TOKUKATU, "			'���ʊ���
		w_sSql = w_sSql & "		A.T13_TOKUKATU_DET, "		'���ʊ����ڍ�
		w_sSql = w_sSql & "		A.T13_NENSYOKEN, "			'�w����Q�l�ƂȂ鏔����
		w_sSql = w_sSql & "		A.T13_NENSYOKEN2, "			'�w����Q�l�ƂȂ鏔����2
		w_sSql = w_sSql & "		A.T13_NENSYOKEN3, "			'�w����Q�l�ƂȂ鏔����3
		w_sSql = w_sSql & "		A.T13_SINTYO, "				'�g��
		w_sSql = w_sSql & "		A.T13_TAIJYU, "				'�̏d
		w_sSql = w_sSql & "		A.T13_SEKIJI_TYUKAN_Z, "	'�O�����ԐȎ�
		w_sSql = w_sSql & "		A.T13_SEKIJI_KIMATU_Z, " 	'�O�������Ȏ�
		w_sSql = w_sSql & "		A.T13_SEKIJI_TYUKAN_K, " 	'��������Ȏ�
		w_sSql = w_sSql & "		A.T13_SEKIJI, "				'�w�N���Ȏ�
		w_sSql = w_sSql & "		A.T13_NINZU_TYUKAN_Z, "		'�O�����ԃN���X�l��
		w_sSql = w_sSql & "		A.T13_NINZU_KIMATU_Z, "  	'�O�������N���X�l��
		w_sSql = w_sSql & "		A.T13_NINZU_TYUKAN_K, "  	'������ԃN���X�l��
		w_sSql = w_sSql & "		A.T13_CLASSNINZU, "			'�w�N���N���X�l��
		w_sSql = w_sSql & "		A.T13_HEIKIN_TYUKAN_Z, " 	'�O�����ԕ��ϓ_
		w_sSql = w_sSql & "		A.T13_HEIKIN_KIMATU_Z, " 	'�O���������ϓ_
		w_sSql = w_sSql & "		A.T13_HEIKIN_TYUKAN_K, " 	'������ԕ��ϓ_
		w_sSql = w_sSql & "		A.T13_HEIKIN_KIMATU_K, " 	'�w�N�����ϓ_
		w_sSql = w_sSql & "		A.T13_SUMJYUGYO, "			'�����Ɠ���
		w_sSql = w_sSql & "		A.T13_SUMSYUSSEKI, "		'�o�ȓ���
		w_sSql = w_sSql & "		A.T13_SUMRYUGAKU, "			'���w���̎��Ɠ���
		w_sSql = w_sSql & "		A.T13_KESSEKI_TYUKAN_Z, "	'�O�����Ԍ��ȓ���
		w_sSql = w_sSql & "		A.T13_KESSEKI_KIMATU_Z, "	'�O���������ȓ���
		w_sSql = w_sSql & "		A.T13_KESSEKI_TYUKAN_K, "	'�������Ԍ��ȓ���
		w_sSql = w_sSql & "		A.T13_SUMKESSEKI, "			'�w�N�����ȓ����i�����ȓ����j
		w_sSql = w_sSql & "		A.T13_KIBIKI_TYUKAN_Z, "	'�O�����Ԋ���������
		w_sSql = w_sSql & "		A.T13_KIBIKI_KIMATU_Z, "	'�O����������������
		w_sSql = w_sSql & "		A.T13_KIBIKI_TYUKAN_K, "	'������Ԋ���������
		w_sSql = w_sSql & "		A.T13_SUMKIBTEI, "			'�o�Ȓ�~�����������i�����������j
		w_sSql = w_sSql & "		A.T13_NENBIKO "				'�w���p�Q�l�ƂȂ鏔����
		w_sSql = w_sSql & " FROM  "
		w_sSql = w_sSql & " 	T13_GAKU_NEN A, "
		w_sSql = w_sSql & " 	M02_GAKKA    B, "
		w_sSql = w_sSql & " 	M01_KUBUN E  "
		w_sSql = w_sSql & " WHERE "
		w_sSql = w_sSql & " 	 A.T13_GAKKA_CD   = B.M02_GAKKA_CD(+) "
		w_sSql = w_sSql & "  AND A.T13_NENDO      = B.M02_NENDO(+) "
		w_sSql = w_sSql & "  AND A.T13_NENDO      = " & mHyoujiNendo
		w_sSql = w_sSql & "  AND A.T13_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' "
		w_sSql = w_sSql & " 	AND A.T13_NENDO		   = E.M01_NENDO "
		w_sSql = w_sSql & " 	AND E.M01_DAIBUNRUI_CD = " & C_ZAISEKI				'�ݐЋ敪
		w_sSql = w_sSql & " 	AND E.M01_SYOBUNRUI_CD = T13_ZAISEKI_KBN "				'�ݐЋ敪

		iRet = gf_GetRecordset(m_Rs, w_sSql)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_GetDetailGakunen = 99
			Exit Do
		End If

		'//�����敪���擾
		if Not gf_GetKubunName(C_NYURYO,m_Rs("T13_RYOSEI_KBN"),Session("HyoujiNendo"),m_RYOSEI_KBN) then Exit Do

		'//

		'//�i���敪���擾
		Select Case gf_SetNull2String(m_Rs("T13_RYUNEN_FLG"))
			Case "0"
				m_RYUNEN_FLG = " �| "
			Case "1"
				m_RYUNEN_FLG = "���N"
			Case Else
				m_RYUNEN_FLG = " �| "
		End Select

		'//����I��
		f_GetDetailGakunen = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [�@�\]  ���������擾����
'*  [����]  p_sClubCd:����CD
'*  [�ߒl]  f_GetClubName�F������
'*  [����]  
'********************************************************************************
Function f_GetClubName(p_sClubCd)

	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetClubName = ""
	w_sClubName = ""

	Do

		'//����CD����̎�
		If trim(gf_SetNull2String(p_sClubCd)) = "" Then
			Exit Do
		End If

		'//���������擾
		w_sSql = ""
		w_sSql = w_sSql & vbCrLf & " SELECT "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_BUKATUDOMEI "
		w_sSql = w_sSql & vbCrLf & " FROM "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO"
		w_sSql = w_sSql & vbCrLf & " WHERE "
		w_sSql = w_sSql & vbCrLf & "  M17_BUKATUDO.M17_NENDO=" & mHyoujiNendo
		w_sSql = w_sSql & vbCrLf & "  AND M17_BUKATUDO.M17_BUKATUDO_CD=" & p_sClubCd

		'//ں��޾�Ď擾
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit Do
		End If

		'//�f�[�^���擾�ł����Ƃ�
		If rs.EOF = False Then
			'//������
			w_sClubName = rs("M17_BUKATUDOMEI")
		End If

		Exit Do
	Loop

	'//�߂�l���
	f_GetClubName = w_sClubName

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �S�C���擾����
'*  [����]  p_sGAKKACd:�w��CD
'*  [�ߒl]  f_GetTanninName�F�S�C��
'*  [����]  
'********************************************************************************
Function f_GetTanninName(p_sGAKKACd,p_iGAKUNEN)

	Dim w_iRet
	Dim w_sSQL
	Dim rs

	On Error Resume Next
	Err.Clear

	f_GetTanninName = ""
	w_sTanninName = ""

	Do

		'//�w��CD����̎�
		If trim(gf_SetNull2String(p_sGAKKACd)) = "" Then
			Exit Do
		End If
		'//�w�N����̎�
		If trim(gf_SetNull2Zero(p_iGAKUNEN)) = 0 Then
			Exit Do
		End If

		'//�S�C�擾
		w_sSql = "Select "
	    w_sSql = w_sSql & " M04_KYOKANMEI_SEI,"
	    w_sSql = w_sSql & " M04_KYOKANMEI_MEI "
	    w_sSql = w_sSql & " From"
	    w_sSql = w_sSql & " M05_CLASS,"
	    w_sSql = w_sSql & " M04_KYOKAN "
	    w_sSql = w_sSql & " WHERE "
		w_sSql = w_sSql & " M05_NENDO =" & mHyoujiNendo
		w_sSql = w_sSql & " And "
	    w_sSql = w_sSql & " M05_NENDO = M04_NENDO "
		w_sSql = w_sSql & " And "
	    'w_sSql = w_sSql & " M04_GAKKA_CD =" & p_sGAKKACd   '�w�ȃR�[�h
	    w_sSql = w_sSql & " M05_CLASSNO =" & p_sGAKKACd   '�N���X�R�[�h
	    w_sSql = w_sSql & " And "
		w_sSql = w_sSql & " M05_GAKUNEN =" & p_iGAKUNEN '�w�N
	    w_sSql = w_sSql & " And "
		w_sSql = w_sSql & " M05_TANNIN = M04_KYOKAN_CD " '����
'response.write w_ssql
		'//ں��޾�Ď擾
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit Do
		End If

		'//�f�[�^���擾�ł����Ƃ�
		If rs.EOF = False Then
			'//������
			w_sTanninName = rs("M04_KYOKANMEI_SEI") & "  " & rs("M04_KYOKANMEI_MEI")
		End If

		Exit Do
	Loop

	'//�߂�l���
	f_GetTanninName = w_sTanninName

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �N���X�ψ����擾����
'*  [����]  p_sGAKKACd:�w��CD
'*  [�ߒl]  f_GetTanninName�F�S�C��
'*  [����]  
'********************************************************************************
Function f_GetIinName()

	Dim w_iRet
	Dim w_sSQL
	Dim rs

	Dim w_sGakki
	Dim w_sZenki_Start
	Dim w_sKouki_Start
	Dim w_sKouki_End

	On Error Resume Next
	Err.Clear

	f_GetIinName = ""
	w_sIinName = ""
	w_sGakki = ""
	w_sZenki_Start = ""
	w_sKouki_Start = ""
	w_sKouki_End = ""

	Do
		'�w�������i���݂̊w���j
		Call gf_GetGakkiInfo(w_sGakki,w_sZenki_Start,w_sKouki_Start,w_sKouki_End)

		'//�ψ��擾
		w_sSql = ""
		w_sSql = w_sSql & "SELECT "
		w_sSql = w_sSql & "M34_IIN_NAME "
		w_sSql = w_sSql & "FROM "
		w_sSql = w_sSql & "M34_IIN, "
		w_sSql = w_sSql & "T06_GAKU_IIN "
		w_sSql = w_sSql & "WHERE "
		w_sSql = w_sSql & "M34_NENDO =" & mHyoujiNendo
		w_sSql = w_sSql & " AND "
		w_sSql = w_sSql & "T06_NENDO = M34_NENDO "
		w_sSql = w_sSql & " AND "
		w_sSql = w_sSql & "M34_DAIBUN_CD = T06_DAIBUN_CD "
		w_sSql = w_sSql & "AND "
		w_sSql = w_sSql & "M34_SYOBUN_CD = T06_SYOBUN_CD "
		w_sSql = w_sSql & "AND "
		w_sSql = w_sSql & "T06_IIN_KBN=2 "
		w_sSql = w_sSql & "AND "
		w_sSql = w_sSql & "T06_GAKKI_KBN = " & w_sGakki
		w_sSql = w_sSql & "AND "
		w_sSql = w_sSql & "M34_IIN_KBN = T06_IIN_KBN "
		w_sSql = w_sSql & "AND "
		w_sSql = w_sSql & "T06_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' "

		'//ں��޾�Ď擾
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit Do
		End If

		'//�f�[�^���擾�ł����Ƃ�
		If rs.EOF = False Then
			'//�ψ���
			w_sIinName = rs("M34_IIN_NAME")
		End If

		Exit Do
	Loop

	'//�߂�l���
	f_GetIinName = w_sIinName

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  �N���X�ψ����擾����
'*  [����]  p_sGAKKACd:�w��CD�Ap_sCourseCd:�R�[�XCD
'*  [�ߒl]  p_sCourseCd�F�R�[�X��
'*  [����]  
'********************************************************************************
Function f_GetCourseName(p_iNendo,p_iGakunen,p_sGakkaCd,p_sCourseCd)

	Dim w_iRet
	Dim w_sSQL
	Dim rs
	Dim w_sName

	On Error Resume Next
	Err.Clear

	f_GetCourseName = ""
	w_sName = ""

	Do
		'//�ψ��擾
		w_sSql = ""
		w_sSql = ""
		w_sSql = w_sSql & "SELECT "
		w_sSql = w_sSql & "M20_NENDO,"
		w_sSql = w_sSql & "M20_GAKKA_CD,"
		w_sSql = w_sSql & "M20_GAKUNEN,"
		w_sSql = w_sSql & "M20_COURSE_CD,"
		w_sSql = w_sSql & "M20_COURSEMEI,"
		w_sSql = w_sSql & "M20_COURSEMEI_EIGO,"
		w_sSql = w_sSql & "M20_COURSERYAKSYO,"
		w_sSql = w_sSql & "M20_COURSE_KIGO,"
		w_sSql = w_sSql & "M20_COURSE_TEIIN "
		w_sSql = w_sSql & "FROM "
		w_sSql = w_sSql & "M20_COURSE "
		w_sSql = w_sSql & "WHERE "
		w_sSql = w_sSql & "M20_NENDO = " & p_iNendo & " AND "
		w_sSql = w_sSql & "M20_GAKKA_CD = '" & p_sGakkaCd & "' AND "
		w_sSql = w_sSql & "M20_GAKUNEN = " & p_iGakunen & " AND "
		w_sSql = w_sSql & "M20_COURSE_CD = '" & p_sCourseCd & "'"

		'//ں��޾�Ď擾
		w_iRet = gf_GetRecordset(rs, w_sSQL)
		If w_iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			Exit Do
		End If

		'//�f�[�^���擾�ł����Ƃ�
		If rs.EOF = False Then
			'//�ψ���
			w_sName = rs("M20_COURSEMEI")
		End If

		Exit Do
	Loop

	'//�߂�l���
	f_GetCourseName = w_sName

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

End Function

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()

	On Error Resume Next
	Err.Clear

	m_HyoujiFlg = 0 		'<!-- �\���t���O�i0:�Ȃ�  1:����j

	m_NENDO				= "" 	'�����N�x
	m_GAKUSEKI_NO		= "" 	'�w�Дԍ�
	m_GAKUNEN			= "" 	'�w�N
	m_ZAISEKI_KBN		= "" 	'�ݐЋ敪�@��
	m_GAKKA_CD			= "" 	'�w��CD�@��
	m_COURCE_CD			= "" 	'�R�[�XCD�@��
	m_CLASS				= "" 	'�N���XCD�@��
	m_TANNIN			= "" 	'�S�C���@��
	m_SYUSEKI_NO1		= "" 	'�o�Ȕԍ��i�w�ȁj
	m_SYUSEKI_NO2		= "" 	'�o�Ȕԍ��i�N���X�j
	'm_RYOSEI_KBN		= "" 	'�����敪�@��
	m_RYUNEN_FLG		= "" 	'���N�敪�@��
	m_IIN				= "" 	'�N���X�����@��
	m_CLUB_1			= "" 	'�N���u�P�@��
	m_CLUB_1_NYUBI		= "" 	'�N���u�����P
	m_CLUB_2			= "" 	'�N���u�Q�@��
	m_CLUB_2_NYUBI		= "" 	'�N���u�����Q
	m_TOKUKATU			= "" 	'���ʊ���
	m_TOKUKATU_DET		= "" 	'���ʊ����ڍ�
	m_NENSYOKEN			= ""	'�w����Q�l�ƂȂ鏔����
	m_NENSYOKEN2		= ""	'�w����Q�l�ƂȂ鏔����2
	m_NENSYOKEN3		= ""	'�w����Q�l�ƂȂ鏔����3
	m_SINTYO			= "" 	'�g��
	m_TAIJYU			= "" 	'�̏d
	m_SEKIJI_TYUKAN_Z	= "" 	'�O�����ԐȎ�
	m_SEKIJI_KIMATU_Z 	= "" 	'�O�������Ȏ�
	m_SEKIJI_TYUKAN_K 	= "" 	'��������Ȏ�
	m_SEKIJI		 	= "" 	'�w�N���Ȏ�
	m_NINZU_TYUKAN_Z	= ""	'�O�����ԃN���X�l��
	m_NINZU_KIMATU_Z	= "" 	'�O�������N���X�l��
	m_NINZU_TYUKAN_K  	= "" 	'������ԃN���X�l��
	m_CLASSNINZU		= "" 	'�w�N���N���X�l��
	m_HEIKIN_TYUKAN_Z 	= "" 	'�O�����ԕ��ϓ_
	m_HEIKIN_KIMATU_Z 	= "" 	'�O���������ϓ_
	m_HEIKIN_TYUKAN_K 	= "" 	'������ԕ��ϓ_
	m_HEIKIN_KIMATU_K 	= "" 	'�w�N�����ϓ_
	m_SUMJYUGYO			= "" 	'�����Ɠ���
	m_SUMSYUSSEKI		= "" 	'�o�ȓ���
	m_SUMRYUGAKU		= "" 	'���w���̎��Ɠ���
	m_KESSEKI_TYUKAN_Z	= "" 	'�O�����Ԍ��ȓ���
	m_KESSEKI_KIMATU_Z	= ""	'�O���������ȓ���
	m_KESSEKI_TYUKAN_K	= ""	'�������Ԍ��ȓ���
	m_SUMKESSEKI		= "" 	'�w�N�����ȓ����i�����ȓ����j
	m_KIBIKI_TYUKAN_Z	= ""	'�O�����Ԋ���������
	m_KIBIKI_KIMATU_Z	= ""	'�O����������������
	m_KIBIKI_TYUKAN_K	= ""	'������Ԋ���������
	m_SUMKIBTEI			= ""	'�o�Ȓ�~�����������i�����������j
'					    = "" 	'���Ɨ��Ə�
'					�@  = "" 	'���w��
	m_NENBIKO		    = "" 	'�w���p�Q�l�ƂȂ鏔����

	if Not m_Rs.Eof Then
		m_NENDO			= m_Rs("T13_NENDO")
		m_GAKUSEKI_NO	= m_Rs("T13_GAKUSEKI_NO")
		m_GAKUNEN		= m_Rs("T13_GAKUNEN")
		m_ZAISEKI_KBN	= m_Rs("M01_SYOBUNRUIMEI")
		m_GAKKAMEI	 	= m_Rs("M02_GAKKAMEI")
		m_COURCE_CD		= m_Rs("T13_COURCE_CD")
		m_COURCEMEI		= f_GetCourseName(gf_SetNull2Zero(m_Rs("T13_NENDO")),gf_SetNull2Zero(m_Rs("T13_GAKUNEN")),gf_SetNull2String(m_Rs("T13_GAKKA_CD")),gf_SetNull2String(m_Rs("T13_COURCE_CD")))
		m_CLASS			= m_Rs("T13_CLASS")
		m_TANNIN		= f_GetTanninName(gf_SetNull2String(m_Rs("T13_CLASS")),gf_SetNull2Zero(m_Rs("T13_GAKUNEN")))
		m_SYUSEKI_NO1	= m_Rs("T13_SYUSEKI_NO1")
		m_SYUSEKI_NO2	= m_Rs("T13_SYUSEKI_NO2")
		m_IIN			= f_GetIinName
		'm_RYOSEI_KBN	= m_Rs("T13_RYOSEI_KBN")

		If gf_SetNull2String(m_Rs("T13_RYUNEN_FLG")) = "1" Then
			m_RYUNEN_FLG	= "���N"
		Else
			m_RYUNEN_FLG	= " �| "
		End If

		m_CLUB_1		= f_GetClubName(gf_SetNull2String(m_Rs("T13_CLUB_1")))
		m_CLUB_1_NYUBI	= m_Rs("T13_CLUB_1_NYUBI")
		m_CLUB_2		= f_GetClubName(gf_SetNull2String(m_Rs("T13_CLUB_2")))
		m_CLUB_2_NYUBI	= m_Rs("T13_CLUB_2_NYUBI")
		m_TOKUKATU		= m_Rs("T13_TOKUKATU")
		m_TOKUKATU_DET	= m_Rs("T13_TOKUKATU_DET")
		m_NENSYOKEN		= m_Rs("T13_NENSYOKEN")
		m_NENSYOKEN2	= m_Rs("T13_NENSYOKEN2")
		m_NENSYOKEN3	= m_Rs("T13_NENSYOKEN3")
		m_SINTYO		= m_Rs("T13_SINTYO")
		m_TAIJYU		= m_Rs("T13_TAIJYU")
		m_SEKIJI_TYUKAN_Z	= m_Rs("T13_SEKIJI_TYUKAN_Z")
		m_SEKIJI_KIMATU_Z	= m_Rs("T13_SEKIJI_KIMATU_Z")
		m_SEKIJI_TYUKAN_K	= m_Rs("T13_SEKIJI_TYUKAN_K")
		m_SEKIJI		 	= m_Rs("T13_SEKIJI")
		m_NINZU_TYUKAN_Z	= m_Rs("T13_NINZU_TYUKAN_Z")
		m_NINZU_KIMATU_Z	= m_Rs("T13_NINZU_KIMATU_Z")
		m_NINZU_TYUKAN_K	= m_Rs("T13_NINZU_TYUKAN_K")
		m_CLASSNINZU		= m_Rs("T13_CLASSNINZU")
		m_HEIKIN_TYUKAN_Z	= m_Rs("T13_HEIKIN_TYUKAN_Z")
		m_HEIKIN_KIMATU_Z	= m_Rs("T13_HEIKIN_KIMATU_Z")
		m_HEIKIN_TYUKAN_K	= m_Rs("T13_HEIKIN_TYUKAN_K")
		m_HEIKIN_KIMATU_K	= m_Rs("T13_HEIKIN_KIMATU_K")
		m_SUMJYUGYO			= m_Rs("T13_SUMJYUGYO")
		m_SUMSYUSSEKI		= m_Rs("T13_SUMSYUSSEKI")
		m_SUMRYUGAKU		= m_Rs("T13_SUMRYUGAKU")
		m_KESSEKI_TYUKAN_Z	= m_Rs("T13_KESSEKI_TYUKAN_Z")
		m_KESSEKI_KIMATU_Z	= m_Rs("T13_KESSEKI_KIMATU_Z")
		m_KESSEKI_TYUKAN_K	= m_Rs("T13_KESSEKI_TYUKAN_K")
		m_SUMKESSEKI		= m_Rs("T13_SUMKESSEKI")
		m_KIBIKI_TYUKAN_Z	= m_Rs("T13_KIBIKI_TYUKAN_Z")
		m_KIBIKI_KIMATU_Z	= m_Rs("T13_KIBIKI_KIMATU_Z")
		m_KIBIKI_TYUKAN_K	= m_Rs("T13_KIBIKI_TYUKAN_K")
		m_SUMKIBTEI	= m_Rs("T13_SUMKIBTEI")
		m_NENBIKO	= m_Rs("T13_NENBIKO")
        
	End if

%>

	<html>
	<head>
	<title>�w�Ѓf�[�^�Q��</title>
	<meta http-equiv="Content-Type" content="text/html; charset=x-sjis">
    <link rel=stylesheet href=../../common/style.css type=text/css>
	<style type="text/css">
	<!--
		a:link { color:#cc8866; text-decoration:none; }
		a:visited { color:#cc8866; text-decoration:none; }
		a:active { color:#888866; text-decoration:none; }
		a:hover { color:#888866; text-decoration:underline; }
		b { color:#88bbbb; font-weight: bold; font-size:14px}
	//-->
	</style>
	<script language="javascript">
	<!--
		//**************************************
		//*   �N�x�ڸ��ޯ�����ύX���ꂽ�Ƃ�
		//**************************************
		function jf_ChangSelect(){

			document.frm.submit();

		}

	//-->
	</script>
	</head>

	<body>
	<form action="kojin_sita3.asp" method="post" name="frm" target="fMain">
	<div align="center">

	<br><br>
	<table border="0" cellpadding="0" cellspacing="0" width="600">
		<tr>
			<td nowrap><a href="kojin_sita0.asp">����{���</a></td>
			<td nowrap><a href="kojin_sita1.asp">���l���</a></td>
			<td nowrap><a href="kojin_sita2.asp">�����w���</a></td>
			<td nowrap><b>���w�N���</b></td>
			<td nowrap><a href="kojin_sita4.asp">�����̑��\�����</a></td>
			<td nowrap><a href="kojin_sita5.asp">���ٓ����</a></td>
		</tr>
	</table>
	<br>

	<table border="0" cellpadding="1" cellspacing="1">
		<tr>
			<td colspan="3">
				<span class="msg"><font size="2">�� �����N�x��ύX����ƁA�ߋ��̊w�N�������邱�Ƃ��ł��܂�<BR></font></span>
			</td>
		</tr>
		<tr>
			<td valign="top" align="left">

				<table class="disp" border="1" width="240">
						<tr>
							<td class="disph" width="80">�����N�x</td>
							<td class="disp"><select name="selNendo" onChange="jf_ChangSelect();">
												<% do until m_KakoRs.Eof 
													wSelected = ""
													if Cint(mHyoujiNendo) = Cint(m_KakoRs("T13_NENDO")) then
														wSelected = "selected"
													End if
													%>
													<option value="<%=m_KakoRs("T13_NENDO")%>" <%=wSelected%>><%=m_KakoRs("T13_NENDO")%>�N�x
												<% m_KakoRs.MoveNext : Loop %>
											</select></td>
						</tr>
<!--
						<tr>
							<td class="disph" width="80" height="16">�����N�x</td>
							<td class="disp"><%= m_NENDO %>&nbsp</td>
						</tr>
-->

					<% if gf_empItem(C_T13_GAKUSEKI_NO) then %>
						<tr>
							<td class="disph" height="16"><%=gf_GetGakuNomei(Session("HyoujiNendo"),C_K_KOJIN_1NEN)%></td>
							<td class="disp"><%= m_GAKUSEKI_NO %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_GAKUNEN) then %>
						<tr>
							<td class="disph" height="16">�w�@�@�N</td>
							<td class="disp"><%= m_GAKUNEN %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_ZAISEKI_KBN) then %>
						<tr>
							<td class="disph" height="16">�ݐЋ敪</td>
							<td class="disp"><%= m_ZAISEKI_KBN %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_GAKKA_CD) then %>
						<tr>
							<td class="disph" height="16">�����w��</td>
							<td class="disp"><%= m_GAKKAMEI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_COURCE_CD) then %>
						<tr>
							<td class="disph" height="16">�R�[�X</td>
							<td class="disp"><%= m_COURCEMEI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_CLASS) then %>
						<tr>
							<td class="disph" height="16">�N���X</td>
							<td class="disp"><%= f_GetClass(m_CLASS,m_GAKUNEN) %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T_TANNIN) then %>
						<tr>
							<td class="disph" height="16">�S�C��</td>
							<td class="disp"><%= m_TANNIN %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SYUSEKI_NO1) then %>
						<tr>
							<td class="disph" height="16">�o�Ȕԍ�(�w��)</td>
							<td class="disp"><%= m_SYUSEKI_NO1 %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SYUSEKI_NO2) then %>
						<tr>
							<td class="disph" height="16">�o�Ȕԍ�(�N���X)</td>
							<td class="disp"><%= m_SYUSEKI_NO2 %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_RYOSEI_KBN) then %>
						<tr>
							<td class="disph" height="16">�����敪</td>
							<td class="disp"><%= m_RYOSEI_KBN %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_RYUNEN_FLG) then %>
						<tr>
							<td class="disph" height="16">�i���敪</td>
							<td class="disp"><%= m_RYUNEN_FLG %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T_CLASSIIN) then %>
						<tr>
							<td class="disph" height="16">�N���X����</td>
							<td class="disp"><%= m_IIN %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_CLUB_1) then %>
						<tr>
							<td class="disph" height="16">�N���u�����P</td>
							<td class="disp"><%= m_CLUB_1 %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_CLUB_1_NYUBI) then %>
						<tr>
							<td class="disph" height="16">�������P</td>
							<td class="disp"><%= m_CLUB_1_NYUBI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_CLUB_2) then %>
						<tr>
							<td class="disph" height="16">�N���u�����Q</td>
							<td class="disp"><%= m_CLUB_2 %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_CLUB_2_NYUBI) then %>
						<tr>
							<td class="disph" height="16">�������Q</td>
							<td class="disp"><%= m_CLUB_2_NYUBI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_TOKUKATU) then %>
						<tr>
							<td class="disph" height="16">���ʊ���</td>
							<td class="disp"><%= m_TOKUKATU %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_TOKUKATU_DET) then %>
						<tr>
							<td class="disph" height="16">���ʊ����ڍ�</td>
							<td class="disp"><%= m_TOKUKATU_DET %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SINTYO) then %>
						<tr>
							<td class="disph" height="16">�g�@�@��</td>
							<td class="disp"><%= m_SINTYO %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_TAIJYU) then %>
						<tr>
							<td class="disph" height="16">�́@�@�d</td>
							<td class="disp"><%= m_TAIJYU %>&nbsp</td>
						</tr>
					<% End if %>
				</table>

			</td>
			<td valign="top" align="left">

					<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T13_SEKIJI_TYUKAN_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">�O�����ԐȎ�</td>
							<td class="disp"><%= m_SEKIJI_TYUKAN_Z %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SEKIJI_KIMATU_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">�O�������Ȏ�</td>
							<td class="disp"><%= m_SEKIJI_KIMATU_Z  %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SEKIJI_TYUKAN_K) then %>
						<tr>
							<td class="disph" width="140" height="16">������ԐȎ�</td>
							<td class="disp"><%= m_SEKIJI_TYUKAN_K %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SEKIJI) then %>
						<tr>
							<td class="disph" width="140" height="16">�w�N���Ȏ�</td>
							<td class="disp"><%= m_SEKIJI %>&nbsp</td>
						</tr>
					<% End if %>
			</table>
			<br>

					<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T13_NINZU_TYUKAN_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">�O�����ԃN���X�l��</td>
							<td class="disp"><%= m_NINZU_TYUKAN_Z %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_NINZU_KIMATU_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">�O�������N���X�l��</td>
							<td class="disp"><%= m_NINZU_KIMATU_Z %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_HEIKIN_TYUKAN_K) then %>
						<tr>
							<td class="disph" width="140" height="16">������ԃN���X�l��</td>
							<td class="disp"><%= m_NINZU_TYUKAN_K %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_HEIKIN_KIMATU_K) then %>
						<tr>
							<td class="disph" width="140" height="16">�w�N���N���X�l��</td>
							<td class="disp"><%= m_CLASSNINZU %>&nbsp</td>
						</tr>
					<% End if %>
			</table>
			<br>

					<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T13_HEIKIN_TYUKAN_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">�O�����ԕ��ϓ_</td>
							<td class="disp"><%= m_HEIKIN_TYUKAN_Z %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_HEIKIN_KIMATU_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">�O���������ϓ_</td>
							<td class="disp"><%= m_HEIKIN_KIMATU_Z %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_HEIKIN_TYUKAN_K) then %>
						<tr>
							<td class="disph" width="140" height="16">������ԕ��ϓ_</td>
							<td class="disp"><%= m_HEIKIN_TYUKAN_K  %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_HEIKIN_KIMATU_K) then %>
						<tr>
							<td class="disph" width="140" height="16">�w�N�����ϓ_</td>
							<td class="disp"><%= m_HEIKIN_KIMATU_K  %>&nbsp</td>
						</tr>
					<% End if %>
			</table>
			<br>

					<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T13_SUMJYUGYO) then %>
						<tr>
							<td class="disph" width="140" height="16">�� �� �� �� ��</td>
							<td class="disp"><%= m_SUMJYUGYO %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SUMSYUSSEKI) then %>
						<tr>
							<td class="disph" width="140" height="16">�o �� �� ��</td>
							<td class="disp"><%= m_SUMSYUSSEKI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SUMRYUGAKU) then %>
						<tr>
							<td class="disph" width="140" height="16">���w���̎��Ɠ���</td>
							<td class="disp"><%= m_SUMRYUGAKU %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_KESSEKI_TYUKAN_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">�O�����Ԍ��ȓ���</td>
							<td class="disp"><%= m_KESSEKI_TYUKAN_Z %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_KESSEKI_KIMATU_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">�O���������ȓ���</td>
							<td class="disp"><%= m_KESSEKI_KIMATU_Z %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_KESSEKI_TYUKAN_K) then %>
						<tr>
							<td class="disph" width="140" height="16">������Ԍ��ȓ���</td>
							<td class="disp"><%= m_KESSEKI_TYUKAN_K %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SUMKESSEKI) then %>
						<tr>
							<td class="disph" width="140" height="16">�w�N�����ȓ���</td>
							<td class="disp"><%= m_SUMKESSEKI %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_KIBIKI_TYUKAN_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">�O�����Ԋ�������</td>
							<td class="disp"><%= m_KIBIKI_TYUKAN_Z %>&nbsp</td>
						</tr>
					<% End if %>					
					<% if gf_empItem(C_T13_KIBIKI_KIMATU_Z) then %>
						<tr>
							<td class="disph" width="140" height="16">�O��������������</td>
							<td class="disp"><%= m_KIBIKI_KIMATU_Z %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_KIBIKI_TYUKAN_K) then %>
						<tr>
							<td class="disph" width="140" height="16">������Ԋ�������</td>
							<td class="disp"><%= m_KIBIKI_TYUKAN_K %>&nbsp</td>
						</tr>
					<% End if %>
					<% if gf_empItem(C_T13_SUMKIBTEI) then %>
						<tr>
							<td class="disph" width="140" height="16">�o�Ȓ�~�E��������</td>
							<td class="disp"><%= m_SUMKIBTEI %>&nbsp</td>
						</tr>
					<% End if %>
				</table>
			</td>
			<td valign="top" align="left">

				<table class="disp" border="1" width="220">
					<% if gf_empItem(C_T13_NENSYOKEN) then %>
						<tr><td class="disph" width="220" height="16">�w����Q�l�ƂȂ鏔����</td></tr>
						<tr><td class="disph" width="220" height="16">(1)�w�K�ɂ����������<br>(2)�s���̓����A���Z��</td></tr>
						<tr><td class="disp" valign="top" height="60"><%= m_NENSYOKEN %></td></tr>
					<% End if %>
					<% if gf_empItem(C_T13_NENSYOKEN2) then %>
						<tr><td class="disph" width="220" height="16">(3)�������A�{�����e�B�A������<br>(4)�擾���i�A���蓙</td></tr>
						<tr><td class="disp" valign="top" height="60"><%= m_NENSYOKEN2 %></td></tr>
					<% End if %>
					<% if gf_empItem(C_T13_NENSYOKEN3) then %>
						<tr><td class="disph" width="220" height="16">(5)���̑�</td></tr>
						<tr><td class="disp" valign="top" height="60"><%= m_NENSYOKEN3 %></td></tr>
					<% End if %>
				</table>

			</td>
		</tr>
	</table>

	<% if m_HyoujiFlg = 0 then %>
		<BR>
		�\���ł���f�[�^������܂���<BR>
		<BR>
	<% End if %>

	<BR>
	<input type="button" class="button" value="�@����@" onClick="parent.window.close();">

	</div>
	</form>
	</body>
	</html>
<% End Sub %>