<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ѓo�^
' ��۸���ID : sei/sei0150/sei0150_23_hidprint.asp
' �@      �\: ����������s�Ȃ�
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h		��		SESSION���i�ۗ��j
'           :�N�x			��		SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h		��		SESSION���i�ۗ��j
'           :�N�x			��		SESSION���i�ۗ��j
' ��      ��:
'	(�p�^�[��)
'	�E�ʏ���ƁA���ʊ���
'	�E�Ȗڋ敪(0:��ʉȖ�,1:���Ȗ�)
'	�E�K�C�I���敪(1:�K�C,2:�I��)
'	�E���x���ʋ敪(0:��ʉȖ�,1:���x���ʉȖ�)�𒲂ׂ�
'-------------------------------------------------------------------------
' ��      ��: 2003/05/08 hirota
' �X�@�@�@�V: 2011/06/06 iwata ���C�A�E�g�ύX�i�S�����������A���Ԑ��L����)
' �X�@�@�@�V: 2018/02/09 ���{ ����ԍ������N���X�o�Ȕԍ��ň󎚂���
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
    Dim m_bErrFlg				'//�װ�׸�

    Const C_ERR_GETDATA = "�f�[�^�̎擾�Ɏ��s���܂���"

    '�����I��p��Where����
    Dim m_iNendo				'//�N�x
    Dim m_sKyokanCd				'//�����R�[�h
    Dim m_sSikenKBN				'//�����敪
    Dim m_iGakunen				'//�w�Nm_sGakuNo
    Dim m_sClassNo				'//�w��
    Dim m_sKamokuCd				'//�ȖڃR�[�h
    Dim m_sSikenNm				'//������
    Dim m_sGakkaCd
    Dim m_iKamoku_Kbn
    Dim m_iHissen_Kbn
	Dim m_ilevelFlg
	Dim m_Rs
    Dim m_rCnt					'//���R�[�h�J�E���g
	Dim m_SRs
	Dim m_bSeiInpFlg			'//���͊��ԃt���O
	Dim m_bKekkaNyuryokuFlg		'//���ۓ��͉\�׸�(True:���͉� / False:���͕s��)
	Dim m_iShikenInsertType
	Dim m_sSyubetu
	Dim m_iKamokuKbn			'//�Ȗڋ敪( 0:�ʏ���ƁA 1:���ʉȖ�)
	Dim m_sKamokuBunrui			'//�Ȗڕ���(01:�ʏ���ƁA02:�F��ȖځA03:���ʉȖ�)
	Dim m_iSeisekiInpType
	Dim m_Date
	Dim m_bZenkiOnly
	Dim m_bNiteiFlg
	Dim m_sGakkoNO				'�w�Z�ԍ�
	Dim m_sUpdDate

    Dim m_iIdouEnd        '//�ٓ��Ώۊ��ԏI����

'///////////////////////////���C������/////////////////////////////
	'Ҳ�ٰ�ݎ��s
	Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////

'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Sub Main()
	Dim w_iRet
	Dim w_sSQL
	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

	'Message�p�̕ϐ��̏�����
	w_sWinTitle = "�L�����p�X�A�V�X�g"
	w_sMsgTitle = "���ѓo�^"
	w_sMsg = ""
	w_sRetURL = C_RetURL & C_ERR_RETURL
	w_sTarget = ""

	On Error Resume Next
	Err.Clear

	m_bErrFlg = false

	Do

		'//�ް��ް��ڑ�
		If gf_OpenDatabase() <> 0 Then
			m_bErrFlg = True
			Exit Do
		End If

		'//�s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

		'//���Ұ�SET
		Call s_SetParam()

		'�w�Z�ԍ��̎擾
		if Not gf_GetGakkoNO(m_sGakkoNO) then Exit Do


		'//���ѓ��͕��@�̎擾(0:�_��[C_SEISEKI_INP_TYPE_NUM]�A1:����[C_SEISEKI_INP_TYPE_STRING]�A2:���ہA�x��[C_SEISEKI_INP_TYPE_KEKKA])
		if not gf_GetKamokuSeisekiInp(m_iNendo,m_sKamokuCd,m_sKamokuBunrui,m_iSeisekiInpType) then
			m_bErrFlg = True
			Exit Do
		end if


		'//���сA���ۓ��͊��ԃ`�F�b�N
		If not f_Nyuryokudate() Then
			m_bErrFlg = True
			Exit Do
		End If

		'//�O���̂݊J�݂��ʔN�����ׂ�
		if not f_SikenInfo(m_bZenkiOnly) then
			m_bErrFlg = True
			Exit Do
		end if

		'//�F��O����擾
		if not gf_GetGakunenNintei(m_iNendo,cint(m_iGakunen),m_bNiteiFlg) then
			m_bErrFlg = True
			Exit Do
		end if

		If m_iKamokuKbn = C_JIK_JUGYO then  '�ʏ���Ƃ̏ꍇ
			'//�Ȗڏ����擾
			'//�Ȗڋ敪(0:��ʉȖ�,1:���Ȗ�)�A�y�сA�K�C�I���敪(1:�K�C,2:�I��)�𒲂ׂ�
			'//���x���ʋ敪(0:��ʉȖ�,1:���x���ʉȖ�)�𒲂ׂ�
			If not f_GetKamokuInfo(m_iKamoku_Kbn,m_iHissen_Kbn,m_ilevelFlg) Then m_bErrFlg = True : Exit Do
		end if

		'//���сA�w���f�[�^�擾
		If not f_GetStudent() Then m_bErrFlg = True : Exit Do


		If m_Rs.EOF Then
			Call gs_showWhitePage("�l���C�f�[�^�����݂��܂���B","���ѓo�^")
			Exit Do
		End If

		'// �y�[�W��\��
		Call showPage()
		Exit Do
	Loop

	'// �װ�̏ꍇ�ʹװ�߰�ނ�\��
	If m_bErrFlg = True Then
		w_sMsg = gf_GetErrMsg()
		if w_sMsg = "" then w_sMsg = C_ERR_GETDATA
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If

	'// �I������
	Call gf_closeObject(m_Rs)
	Call gf_closeObject(m_SRs)
	Call gs_CloseDatabase()


End Sub

'********************************************************************************
'*	[�@�\]	�S���ڂɈ����n����Ă����l��ݒ�
'********************************************************************************
Sub s_SetParam()

	m_iNendo	 = request("txtNendo")
	m_sKyokanCd	 = request("txtKyokanCd")
	m_sSikenKBN	 = cint(request("sltShikenKbn"))
	m_iGakunen	 = Cint(request("txtGakuNo"))
	m_sClassNo	 = cint(request("txtClassNo"))
	m_sKamokuCd	 = request("txtKamokuCd")
	m_sGakkaCd	 = request("txtGakkaCd")
	m_sSyubetu	 = trim(Request("hidSyubetu"))
	m_iShikenInsertType = 0

	m_iKamokuKbn = cint(Request("hidKamokuKbn"))

	if m_iKamokuKbn = C_JIK_JUGYO then
		'�ʏ�Ȗ�
		m_sKamokuBunrui = C_KAMOKUBUNRUI_TUJYO
	else
		'���ʉȖ�
		m_sKamokuBunrui = C_KAMOKUBUNRUI_TOKUBETU
	end if

	m_Date = gf_YYYY_MM_DD(year(date()) & "/" & month(date()) & "/" & day(date()),"/")

End Sub

'********************************************************************************
'*  [�@�\]  �O���J�݂��ǂ������ׂ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Function f_SikenInfo(p_bZenkiOnly)
    Dim w_sSQL
    Dim w_Rs

    On Error Resume Next
    Err.Clear

    f_SikenInfo = false
	p_bZenkiOnly = false

	'//�����敪���O�������̎��́A���̉Ȗڂ��O���݂̂��ʔN���𒲂ׂ�
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	T15_KAMOKU_CD "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T15_RISYU "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	T15_NYUNENDO = " & Cint(m_iNendo)-cint(m_iGakunen)+1
	w_sSQL = w_sSQL & " AND T15_GAKKA_CD = '" & m_sGakkaCd & "'"
	w_sSQL = w_sSQL & " AND T15_KAMOKU_CD= '" & Trim(m_sKamokuCd) & "'"
	w_sSQL = w_sSQL & " AND T15_KAISETU" & m_iGakunen & "=" & C_KAI_ZENKI

	if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function

	'Response.Write "0"

	'//�߂�l���
	If w_Rs.EOF = False Then
		p_bZenkiOnly = True
	End If

	f_SikenInfo = true

	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �R���{�őI�����ꂽ�Ȗڂ̉Ȗڋ敪�y�сA�K�C�I���敪�𒲂ׂ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Function f_GetKamokuInfo(p_iKamoku_Kbn,p_iHissen_Kbn,p_ilevelFlg)
	Dim w_sSQL
	Dim w_Rs

	On Error Resume Next
	Err.Clear

	f_GetKamokuInfo = false

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	T15_RISYU.T15_KAMOKU_KBN"
	w_sSQL = w_sSQL & " 	,T15_RISYU.T15_HISSEN_KBN"
	w_sSQL = w_sSQL & " 	,T15_RISYU.T15_LEVEL_FLG"
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T15_RISYU"
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	T15_RISYU.T15_NYUNENDO=" & cint(m_iNendo) - cint(m_iGakunen) + 1
	w_sSQL = w_sSQL & " AND T15_RISYU.T15_GAKKA_CD='" & m_sGakkaCd & "'"
	w_sSQL = w_sSQL & " AND T15_RISYU.T15_KAMOKU_CD='" & m_sKamokuCd & "' "

	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function

	'//�߂�l���
	If w_Rs.EOF = False Then
		p_iKamoku_Kbn = w_Rs("T15_KAMOKU_KBN")
		p_iHissen_Kbn = w_Rs("T15_HISSEN_KBN")
		p_ilevelFlg = w_Rs("T15_LEVEL_FLG")
	End If

	f_GetKamokuInfo = true

	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*	[�@�\]	�f�[�^�̎擾
'********************************************************************************
Function f_GetStudent()

	Dim w_sSQL
	Dim w_FieldName
	Dim w_Table
	Dim w_TableName
	Dim w_KamokuName

	On Error Resume Next
	Err.Clear

	f_GetStudent = false

	'�Ȗڋ敪
	if m_iKamokuKbn = C_JIK_JUGYO then  '�ʏ���Ƃ̏ꍇ
		w_Table      = "T16"
		w_TableName  = "T16_RISYU_KOJIN"
		w_KamokuName = "T16_KAMOKU_CD"
	else								'�����Ȃǂ̏ꍇ
		w_Table      = "T34"
		w_TableName  = "T34_RISYU_TOKU"
		w_KamokuName = "T34_TOKUKATU_CD"
	end if

	'//�����A���l���͂ɂ��A����Ă���t�B�[���h��ς���
	if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then
		if m_bNiteiFlg and m_iKamokuKbn = C_JIK_JUGYO then
			w_FieldName = "HTEN"
		else
			w_FieldName = "SEI"
		end if
	else
		w_FieldName = "HYOKA"
	end if

	'//�������ʂ̒l���ꗗ��\��
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_Z AS SEI1, "
	w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_Z AS SEI2, "
	w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_TYUKAN_K AS SEI3, "
	w_sSQL = w_sSQL & w_Table & "_" & w_FieldName & "_KIMATU_K AS SEI4, "
	w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_Z       AS KEKA_ZT, "			'���ہi�O�����ԁj
	w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_Z       AS KEKA_ZK, "			'���ہi�O�������j
	w_sSQL = w_sSQL & w_Table & "_KEKA_TYUKAN_K       AS KEKA_KT, "			'���ہi������ԁj
	w_sSQL = w_sSQL & w_Table & "_KEKA_KIMATU_K       AS KEKA_KK, "			'���ہi��������j
	w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_Z  AS TEISI_ZT,"			'��~�i�O�����ԁj
	w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_Z  AS TEISI_ZK,"			'��~�i�O�������j
	w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_TYUKAN_K  AS TEISI_KT,"			'��~�i�������ԁj
	w_sSQL = w_sSQL & w_Table & "_KEKA_NASI_KIMATU_K  AS TEISI_KK,"			'��~�i��������j
	w_sSQL = w_sSQL & w_Table & "_KIBI_TYUKAN_Z       AS KIBI_ZT, "			'�����i�O�����ԁj
	w_sSQL = w_sSQL & w_Table & "_KIBI_KIMATU_Z       AS KIBI_ZK, "			'�����i�O�������j
	w_sSQL = w_sSQL & w_Table & "_KIBI_TYUKAN_K       AS KIBI_KT, "			'�����i��������j
	w_sSQL = w_sSQL & w_Table & "_KIBI_KIMATU_K       AS KIBI_KK, "			'�����i��������j
	w_sSQL = w_sSQL & w_Table & "_KOUKETSU_TYUKAN_Z   AS HAKEN_ZT,"			'�h���i�O�����ԁj
	w_sSQL = w_sSQL & w_Table & "_KOUKETSU_KIMATU_Z   AS HAKEN_ZK,"			'�h���i�O�������j
	w_sSQL = w_sSQL & w_Table & "_KOUKETSU_TYUKAN_K   AS HAKEN_KT,"			'�h���i������ԁj
	w_sSQL = w_sSQL & w_Table & "_KOUKETSU_KIMATU_K   AS HAKEN_KK,"			'�h���i��������j
	w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_Z    AS SOUJI_ZT,"
	w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_Z    AS SOUJI_ZK,"
	w_sSQL = w_sSQL & w_Table & "_SOJIKAN_TYUKAN_K    AS SOUJI_KT,"
	w_sSQL = w_sSQL & w_Table & "_SOJIKAN_KIMATU_K    AS SOUJI_KK,"
	w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_Z   AS JUNJI_ZT,"
	w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_Z   AS JUNJI_ZK,"
	w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_TYUKAN_K   AS JUNJI_KT,"
	w_sSQL = w_sSQL & w_Table & "_JUNJIKAN_KIMATU_K   AS JUNJI_KK,"
	w_sSQL = w_sSQL & w_Table & "_J_JUNJIKAN_TYUKAN_Z AS J_JUNJI_ZT,"
	w_sSQL = w_sSQL & w_Table & "_J_JUNJIKAN_KIMATU_Z AS J_JUNJI_ZK,"
	w_sSQL = w_sSQL & w_Table & "_J_JUNJIKAN_TYUKAN_K AS J_JUNJI_KT,"
	w_sSQL = w_sSQL & w_Table & "_J_JUNJIKAN_KIMATU_K AS J_JUNJI_KK,"
	w_sSQL = w_sSQL & w_Table & "_HYOKA_TYUKAN_Z      AS HYOKA_ZT,  "
	w_sSQL = w_sSQL & w_Table & "_HYOKA_KIMATU_Z      AS HYOKA_ZK,  "
	w_sSQL = w_sSQL & w_Table & "_HYOKA_TYUKAN_K      AS HYOKA_KT,  "
	w_sSQL = w_sSQL & w_Table & "_HYOKA_KIMATU_K      AS HYOKA_KK,  "
	w_sSQL = w_sSQL & w_Table & "_KOUSINBI_TYUKAN_Z   AS KOUSINBI_ZT,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINBI_KIMATU_Z   AS KOUSINBI_ZK,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINBI_TYUKAN_K   AS KOUSINBI_KT,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINBI_KIMATU_K   AS KOUSINBI_KK,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINTIME_TYUKAN_Z AS KOUSINTIME_ZT,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINTIME_KIMATU_Z AS KOUSINTIME_ZK,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINTIME_TYUKAN_K AS KOUSINTIME_KT,"
	w_sSQL = w_sSQL & w_Table & "_KOUSINTIME_KIMATU_K AS KOUSINTIME_KK,"
	w_sSQL = w_sSQL & w_Table & "_HYOKA_FUKA_KBN      AS HYOKA_FUKA, "
	w_sSQL = w_sSQL & w_Table & "_HAITOTANI           AS HAITOTANI, "

	if m_iKamokuKbn = C_JIK_JUGYO then
		w_sSQL = w_sSQL & " 	T16_SELECT_FLG, "
		w_sSQL = w_sSQL & " 	T16_LEVEL_KYOUKAN, "
		w_sSQL = w_sSQL & " 	T16_OKIKAE_FLG, "

'2009/06/15 ins str iwata
	'��������Ǝ��Ԃ̕\���̂��ߍ��ڒǉ�
		w_sSQL = w_sSQL & "T16_MENJYO_FLG        AS Menjo,"
		Select Case m_sSikenKBN
			Case C_SIKEN_ZEN_TYU	'�O������
				w_sSQL = w_sSQL & "T16_DATAKBN_TYUKAN_Z  AS DataKbn,"
			Case C_SIKEN_ZEN_KIM	'�O������
				w_sSQL = w_sSQL & "T16_DATAKBN_KIMATU_Z  AS DataKbn,"
			Case C_SIKEN_KOU_TYU	'�������
				w_sSQL = w_sSQL & "T16_DATAKBN_TYUKAN_K  AS DataKbn,"
			Case C_SIKEN_KOU_KIM	'�������
				w_sSQL = w_sSQL & "T16_DATAKBN_KIMATU_K  AS DataKbn,"
		End Select
'2009/06/15 ins end iwata
	end if

	w_sSQL = w_sSQL & " 	T13_GAKUSEI_NO  AS GAKUSEI_NO, "
	w_sSQL = w_sSQL & " 	T13_GAKUSEKI_NO AS GAKUSEKI_NO,"
	w_sSQL = w_sSQL & " 	T13_SYUSEKI_NO2 AS SYUSEKI_NO,"		'2018.02.09 Add Kiyomoto	�N���X�o�Ȕԍ�
	w_sSQL = w_sSQL & " 	T11_SIMEI       AS SIMEI       "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & 		w_TableName & ","
	w_sSQL = w_sSQL & " 	T11_GAKUSEKI,   "
	w_sSQL = w_sSQL & " 	T13_GAKU_NEN    "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & 				w_Table & "_NENDO      = " & Cint(m_iNendo)
	w_sSQL = w_sSQL & " 	AND	" & w_Table & "_GAKUSEI_NO = T11_GAKUSEI_NO "
	w_sSQL = w_sSQL & " 	AND	" & w_Table & "_GAKUSEI_NO = T13_GAKUSEI_NO "
	w_sSQL = w_sSQL & " 	AND	T13_GAKUNEN = " & cint(m_iGakunen)
	w_sSQL = w_sSQL & " 	AND	T13_CLASS   = " & cint(m_sClassNo)
	w_sSQL = w_sSQL & " 	AND	" & w_KamokuName & "  = '" & m_sKamokuCd & "' "
	w_sSQL = w_sSQL & " 	AND	" & w_Table & "_NENDO = T13_NENDO "

	if m_iKamokuKbn = C_JIK_JUGYO then
		'//�u�����̐��k�͂͂���(C_TIKAN_KAMOKU_MOTO = 1    '�u����)
		w_sSQL = w_sSQL & " AND	T16_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO
	end if

	w_sSQL = w_sSQL & " ORDER BY " & w_Table & "_GAKUSEKI_NO "

	'���R�[�h�擾
	If gf_GetRecordset(m_Rs,w_sSQL) <> 0 Then Exit function

	'�\������X�V���t & ����
	Select Case Cint(m_sSikenKBN)
		Case C_SIKEN_ZEN_TYU : m_sUpdDate = f_fmtWareki(gf_SetNull2String(m_Rs("KOUSINBI_ZT"))) & "�@" & gf_SetNull2String(m_Rs("KOUSINTIME_ZT"))
		Case C_SIKEN_ZEN_KIM : m_sUpdDate = f_fmtWareki(gf_SetNull2String(m_Rs("KOUSINBI_ZK"))) & "�@" & gf_SetNull2String(m_Rs("KOUSINTIME_ZK"))
		Case C_SIKEN_KOU_TYU : m_sUpdDate = f_fmtWareki(gf_SetNull2String(m_Rs("KOUSINBI_KT"))) & "�@" & gf_SetNull2String(m_Rs("KOUSINTIME_KT"))
		Case C_SIKEN_KOU_KIM : m_sUpdDate = f_fmtWareki(gf_SetNull2String(m_Rs("KOUSINBI_KK"))) & "�@" & gf_SetNull2String(m_Rs("KOUSINTIME_KK"))
	End Select

	'//ں��ރJ�E���g�擾
	m_rCnt = gf_GetRsCount(m_Rs)

	f_GetStudent = true

End Function

'********************************************************************************
'*	[�@�\]	�f�[�^�̎擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]
'********************************************************************************
Function f_Nyuryokudate()

	Dim w_sSysDate
	Dim w_Rs

	On Error Resume Next
	Err.Clear

	f_Nyuryokudate = false

	m_bKekkaNyuryokuFlg = false		'���ۓ����׸�
	m_bSeiInpFlg = false

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & "     T24_IDOU_SYURYO "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T24_SIKEN_NITTEI "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & "     T24_NENDO=" & Cint(m_iNendo)
	w_sSQL = w_sSQL & " AND T24_SIKEN_KBN=" & Cint(m_sSikenKBN)
	w_sSQL = w_sSQL & " AND T24_SIKEN_CD='0'"
	w_sSQL = w_sSQL & " AND T24_GAKUNEN=" & m_iGakunen

	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then exit function

	If w_Rs.EOF Then
		exit function
	Else
		m_iIdouEnd = gf_SetNull2String(w_Rs("T24_IDOU_SYURYO"))  '�ٓ��ΏۏI��
	End If

	'���͊��ԓ��Ȃ琳��
	If gf_YYYY_MM_DD(m_iNKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iNSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
		m_bSeiInpFlg = true
	End If

	'���ۓ��͉\�׸�
	If gf_YYYY_MM_DD(m_iKekkaKaishi,"/") <= gf_YYYY_MM_DD(w_sSysDate,"/") And gf_YYYY_MM_DD(m_iKekkaSyuryo,"/") >= gf_YYYY_MM_DD(w_sSysDate,"/") Then
		m_bKekkaNyuryokuFlg = True
	End If

	f_Nyuryokudate = true

End Function

'********************************************************************************
'*	[�@�\]	�f�[�^�̎擾
'********************************************************************************
Function f_Syukketu2New(p_gaku,p_kbn)
	Dim w_GAKUSEI_NO
	Dim w_SYUKKETU_KBN

	f_Syukketu2New = ""
	w_GAKUSEI_NO = ""
	w_SYUKKETU_KBN = ""
	w_SKAISU = ""

	If m_SRs.EOF Then
		Exit Function
	Else
		Do Until m_SRs.EOF
			w_GAKUSEI_NO = m_SRs("T21_GAKUSEKI_NO")
			w_SYUKKETU_KBN = m_SRs("T21_SYUKKETU_KBN")
			w_SKAISU = gf_SetNull2String(m_SRs("KAISU"))

			If Cstr(w_GAKUSEI_NO) = Cstr(p_gaku) AND cstr(w_SYUKKETU_KBN) = cstr(p_kbn) Then
				f_Syukketu2New = w_SKAISU
				Exit Do
			End If

			m_SRs.MoveNext
		Loop

		m_SRs.MoveFirst
	End If

End Function

'********************************************************************************
'*  [�@�\] �ٓ��`�F�b�N
'********************************************************************************
Sub s_IdouCheck(p_GakusekiNo,p_IdouKbn,p_IdouName,p_bNoChangeZK,p_bNoChangeKT,p_bNoChangeKK,p_IdouDate)
	Dim w_IdoutypeName	'�ٓ��󋵖�
	Dim w_IdouDate
	w_IdoutypeName = ""
	w_IdouDate = ""

	p_IdouName = ""
	p_IdouDate = ""

	m_Date = m_iIdouEnd

	Call f_Get_IdouChk(p_GakusekiNo,m_Date,m_iNendo,w_IdoutypeName,p_IdouKbn,w_IdouDate)

	if Cstr(p_IdouKbn) <> "" and Cstr(p_IdouKbn) <> CStr(C_IDO_FUKUGAKU) AND _
		Cstr(p_IdouKbn) <> Cstr(C_IDO_TEI_KAIJO) AND Cstr(p_IdouKbn) <> Cstr(C_IDO_TENKO) AND _
		Cstr(p_IdouKbn) <> Cstr(C_IDO_TENKA) AND Cstr(p_IdouKbn) <> Cstr(C_IDO_KOKUHI) Then

		p_IdouName = "[" & w_IdoutypeName & "]"
		p_IdouDate = w_IdouDate

		p_bNoChangeZK = True
		p_bNoChangeKT = True
		p_bNoChangeKK = True
	end if

end Sub

'********************************************************************************
'*	[�@�\]	�ٓ�����̏ꍇ�ړ��󋵂̎擾
'*	[����]	p_Gakusei_No:�w��NO
'*			p_Date		:���Ǝ��{��
'*	[�ߒl]	0:���擾���� 99:���s
'*	[����]	2001.12.19 �ŁF���c
'********************************************************************************
Function f_Get_IdouChk(p_Gakuseki_No,p_Date,p_iNendo,ByRef p_sKubunName,ByRef p_sKubunCD,ByRef p_sIdouDate)

	Dim w_sSQL
	Dim w_Rs
	Dim w_IdoFlg
	Dim w_sKubunName

	On Error Resume Next
	Err.Clear

	f_Get_IdouChk = False
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
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_NENDO=" & cint(p_iNendo) & " AND "
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_GAKUSEKI_NO='" & p_Gakuseki_No & "' AND"
		w_sSQL = w_sSQL & vbCrLf & "  T13_GAKU_NEN.T13_IDOU_NUM>0"

'response.write w_sSQL

		iRet = gf_GetRecordset(w_Rs, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			Exit Do
		End If

		If w_Rs.EOF = 0 Then
			i = 1
			'//8�c�ő�ړ���
			Do Until Cint(i) > cint(8)    '//C_IDO_MAX_CNT = 8�c�ő�ړ���
				If gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) = "" Then
					Exit Do
				End If
'Response.Write "[" & gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) & " > " & p_Date & "]"
				If gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i)) > p_Date  Then
					'//1���ڂ̈ٓ����Ώۓ��t��薢���̏ꍇ�̏���
					If i = 1 then
						i = 0
					End if

					Exit Do
				End If
				i = i + 1
			Loop

'response.write "�w���m�n" & p_Gakuseki_No & " : i = " & i
			w_sKubunName = ""

			If i = 1 then
				'//�ŏ��̈ړ��������Ɠ���薢���̏ꍇ�A���Ɠ��Ɉړ���Ԃł͂Ȃ�
				'w_IdoFlg = False
				'w_sKubunName = ""

				w_sKubunName = gf_SetNull2String(w_Rs("T13_IDOU_KBN_" & i))
				p_sIdouDate = gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i))
				w_bRet = gf_GetKubunName_R(C_IDO,Trim(w_Rs("T13_IDOU_KBN_" & i)),p_iNendo,p_sKubunName)
			Elseif i = 0 then '//1���ڂ̈ٓ����Ώۓ��t��薢���̏ꍇ

				w_bRet = False
				w_sKubunName = ""
				p_sIdouDate = ""
			Else

   				w_sKubunName = gf_SetNull2String(w_Rs("T13_IDOU_KBN_" & i-1))
				p_sIdouDate = gf_SetNull2String(w_Rs("T13_IDOU_BI_" & i-1))

				 w_bRet = gf_GetKubunName_R(C_IDO,Trim(w_Rs("T13_IDOU_KBN_" & i-1)),p_iNendo,p_sKubunName)

			End If
'response.write "����:" & w_sKubunName & "�ٓ����R�F" & p_sKubunName  & p_sIdouDate
		End If

		Exit Do
	Loop

	p_sKubunCD = w_sKubunName

	Call gf_closeObject(w_Rs)

	Err.Clear

	f_Get_IdouChk = True

End Function



'********************************************************************************
'*  [�@�\] ���т̃Z�b�g
'********************************************************************************
Sub s_SetGrades(p_sSeiseki_ZK,  p_sSeiseki_KT,  p_sSeiseki_KK, _
				p_sHyoka_ZK,    p_sHyoka_KT,    p_sHyoka_KK, _
				p_bNoChange_ZK, p_bNoChange_KT, p_bNoChange_KK)

	Dim w_UpdDateZK
    Dim w_UpdDateKK

	'/�����敪�ɂ���Ď���Ă���A�t�B�[���h��ς���B
	Select Case Cint(m_sSikenKBN)
		Case C_SIKEN_ZEN_TYU
			p_sSeiseki_ZK = gf_SetNull2String(m_Rs("SEI1"))
			p_sHyoka_ZK   = gf_SetNull2String(m_Rs("HYOKA_ZT"))
		Case Else
			p_sSeiseki_ZK = gf_SetNull2String(m_Rs("SEI2"))
			p_sHyoka_ZK   = gf_SetNull2String(m_Rs("HYOKA_ZK"))
	End Select
'DEL2005/10/04 ���� �O���J�݉Ȗڂ͊w�N���\������悤��
'	Select Case Cint(m_sSikenKBN)
'		Case C_SIKEN_KOU_KIM
'			p_sSeiseki_KK = gf_SetNull2String(m_Rs("SEI4"))
'			p_sHyoka_KK   = gf_SetNull2String(m_Rs("HYOKA_KK"))
'		Case Else
'			p_sSeiseki_KK = ""
'			p_sHyoka_KK   = ""
'	End Select
'DEL2005/10/04 ����

	p_sSeiseki_KT = gf_SetNull2String(m_Rs("SEI3"))
	p_sSeiseki_KK = gf_SetNull2String(m_Rs("SEI4"))			'UPP ����2005/10/04 ����
	p_sHyoka_KT   = gf_SetNull2String(m_Rs("HYOKA_KT"))
	p_sHyoka_KK   = gf_SetNull2String(m_Rs("HYOKA_KK"))		'UPP ����2005/10/04 ����

	'�w�N�������̏ꍇ�̂�
'	If m_sSikenKBN = C_SIKEN_KOU_KIM and m_bZenkiOnly = True Then
'		w_UpdDateZK = gf_SetNull2String(m_Rs("KOUSINBI_ZK"))
'		w_UpdDateKK = gf_SetNull2String(m_Rs("KOUSINBI_KK"))
'		if w_UpdDateKK = "" or w_UpdDateZK > w_UpdDateKK then
'			p_sSeiseki_KK = gf_SetNull2String(m_Rs("SEI2"))
'			p_sHyoka_KK   = gf_SetNull2String(m_Rs("HYOKA_ZK"))
'		End If
'	End If

	'//�ʏ���Ƃ̂Ƃ�
	if m_iKamokuKbn = C_JIK_JUGYO then

		p_bNoChange_ZK = False
		p_bNoChange_KT = False
		p_bNoChange_KK = False

		'//�Ȗڂ��I���Ȗڂ̏ꍇ�́A���k���I�����Ă��邩�ǂ����𔻕ʂ���B�I�������Ȃ����k�͓��͕s�Ƃ���B
		if cint(gf_SetNull2Zero(m_iHissen_Kbn)) = cint(gf_SetNull2Zero(C_HISSEN_SEN)) Then
			if cint(gf_SetNull2Zero(m_Rs("T16_SELECT_FLG"))) = cint(C_SENTAKU_NO) Then
				p_bNoChange_ZK = true
				p_bNoChange_KT = true
				p_bNoChange_KK = true
			end if
		else
			if Cstr(m_iLevelFlg) = "1" then
				if isNull(m_Rs("T16_LEVEL_KYOUKAN")) = true then
					p_bNoChange_ZK = true
					p_bNoChange_KT = true
					p_bNoChange_KK = true
				else
					if m_Rs("T16_LEVEL_KYOUKAN") <> m_sKyokanCd then
						p_bNoChange_ZK = true
						p_bNoChange_KT = true
						p_bNoChange_KK = true
					End if
				End if
			End if
		end if
	end if

end Sub

'********************************************************************************
'*  [�@�\]  ���ہA�x�����̃Z�b�g
'********************************************************************************
Sub s_SetKekka(p_sKekka_ZK, p_sKekka_KT, p_sKekka_KK, _
			   p_sKibi_ZK , p_sKibi_KT , p_sKibi_KK , _
			   p_sTeisi_ZK, p_sTeisi_KT, p_sTeisi_KK, _
			   p_sHaken_ZK, p_sHaken_KT, p_sHaken_KK)

	Dim w_UpdDateZK
	Dim w_UpdDateKK

	'/�����敪�ɂ���Ď���Ă���A�t�B�[���h��ς���B
	Select Case Cint(m_sSikenKBN)
		Case C_SIKEN_ZEN_TYU
			p_sKekka_ZK  = gf_SetNull2String(m_Rs("KEKA_ZT"))
			p_sKibi_ZK   = gf_SetNull2String(m_Rs("KIBI_ZT"))
			p_sTeisi_ZK  = gf_SetNull2String(m_Rs("TEISI_ZT"))
			p_sHaken_ZK  = gf_SetNull2String(m_Rs("HAKEN_ZT"))
		Case Else
			p_sKekka_ZK  = gf_SetNull2String(m_Rs("KEKA_ZK"))
			p_sKibi_ZK   = gf_SetNull2String(m_Rs("KIBI_ZK"))
			p_sTeisi_ZK  = gf_SetNull2String(m_Rs("TEISI_ZK"))
			p_sHaken_ZK  = gf_SetNull2String(m_Rs("HAKEN_ZK"))
	End Select

	p_sKekka_KT  = gf_SetNull2String(m_Rs("KEKA_KT"))
	p_sKibi_KT   = gf_SetNull2String(m_Rs("KIBI_KT"))
	p_sTeisi_KT  = gf_SetNull2String(m_Rs("TEISI_KT"))
	p_sHaken_KT  = gf_SetNull2String(m_Rs("HAKEN_KT"))

	'2005/02/28 UP
'DEL 2005/10/04 ���� �O���J�݂̉Ȗڂ͊w�N���т�\�����邽��
'	Select Case Cint(m_sSikenKBN)
'		Case C_SIKEN_KOU_KIM
'			p_sKekka_KK  = gf_SetNull2String(m_Rs("KEKA_KK"))
'			p_sKibi_KK   = gf_SetNull2String(m_Rs("KIBI_KK"))
'			p_sTeisi_KK  = gf_SetNull2String(m_Rs("TEISI_KK"))
'			p_sHaken_KK  = gf_SetNull2String(m_Rs("HAKEN_KK"))
'		Case Else
'			p_sKekka_KK  = ""
'			p_sKibi_KK   = ""
'			p_sTeisi_KK  = ""
'			p_sHaken_KK  = ""
'	End Select

	p_sKekka_KK  = gf_SetNull2String(m_Rs("KEKA_KK"))		'UPP ����2005/10/04 ����
	p_sKibi_KK   = gf_SetNull2String(m_Rs("KIBI_KK"))		'UPP ����2005/10/04 ����
	p_sTeisi_KK  = gf_SetNull2String(m_Rs("TEISI_KK"))		'UPP ����2005/10/04 ����
	p_sHaken_KK  = gf_SetNull2String(m_Rs("HAKEN_KK"))		'UPP ����2005/10/04 ����

End Sub

'********************************************************************************
'*	[�@�\]	�������擾
'********************************************************************************
Function f_ShikenMei()
	Dim w_Rs

	On Error Resume Next
	Err.Clear

	f_ShikenMei = ""

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUIMEI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M01_KUBUN"
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUI_CD = " & cint(m_sSikenKBN)
	w_sSQL = w_sSQL & " AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN)
	w_sSQL = w_sSQL & " AND M01_NENDO = " & cint(m_iNendo)

	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function

	If not w_Rs.EOF Then
		f_ShikenMei = gf_SetNull2String(w_Rs("M01_SYOBUNRUIMEI"))
	End If

    call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �w�Z�����擾
'********************************************************************************
Function f_GetSchoolName()

	Dim w_Rs
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetSchoolName = ""

    '// �w�Z���擾
    w_sSQL = ""
    w_sSQL = w_sSQL & "Select "
    w_sSQL = w_sSQL & "     M19_NAME "
    w_sSQL = w_sSQL & "FROM M19_GAKKO "

	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function

    '// �w�Z��
    f_GetSchoolName = w_Rs("M19_NAME")

    call gf_closeObject(w_Rs)

End Function

'****************************************************
'[�@�\]	�a��t�H�[�}�b�g	:MM��DD���i�j���j
'[����]	pDate : �Ώۓ��t(YYYY/MM/DD)
'[�ߒl]
'****************************************************
Function f_fmtWareki(pDate)

	f_fmtWareki = ""

	'// Null�Ȃ甲����
	if gf_IsNull(trim(pDate)) then	Exit Function

	'// MM��DD���쐬
	w_YY = Left(FormatYYYYMMDD(pDate),4) & "�N"
	w_MM = Mid(FormatYYYYMMDD(pDate),6,2) & "��"
	w_DD = Right(FormatYYYYMMDD(pDate),2) & "��"

	'// �j�����擾
	w_Youbi = WeekdayName(Weekday(FormatYYYYMMDD(pDate))) & "<BR>"
	w_Youbi = "�i" & Left(w_Youbi,1) & "�j"

	f_fmtWareki = w_YY & w_MM & w_DD

End Function

'***********************************************************
' �@�@�@�\�F����N�x����a��N�x�����߂�
' �߁@�@�l�F�ϊ�����
'           (����):�a��A(���s):""
' ���@�@���Fp_sNendo - ����̔N�x
' �ڍ׋@�\�F����N�x����a��N�x�����߂�
' ���@�@�l�F�a��N�x��Ԃ��B�����͂��Ȃ��B
'***********************************************************
Function f_Nendo2Wareki(p_iNendo)
    Dim w_sSql
    Dim w_Rs

	On Error Resume Next
	Err.Clear

    '== ������ ==
    f_Nendo2Wareki = ""

    '== �a��̎擾 ==
    w_sSql = ""
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & " 	M00_KANRI "
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & " 	M00_KANRI "
    w_sSql = w_sSql & " WHERE "
    w_sSql = w_sSql & " 		M00_NENDO = " & p_iNendo & " "
    w_sSql = w_sSql & " 	AND M00_NO    = " & C_K_WAREKI_NENDO

    '== �f�[�^�擾 ==
    If gf_GetRecordset(w_Rs,w_sSql) <> 0 Then Exit function

    f_Nendo2Wareki = gf_SetNull2String(w_Rs("M00_KANRI"))

    '== ���� ==
    call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �K�I�敪���̂��擾
'********************************************************************************
Function f_GetHissenNM(p_iHissen)
    Dim w_sSQL
    Dim w_Rs

	On Error Resume Next
	Err.Clear

	f_GetHissenNM = ""

    '== �a��̎擾 ==
    w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUIMEI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M01_KUBUN "
    w_sSQL = w_sSQL & " WHERE "
    w_sSQL = w_sSQL & " 		M01_NENDO        = " & m_iNendo
    w_sSQL = w_sSQL & " 	AND M01_SYOBUNRUI_CD = " & p_iHissen
	w_sSQL = w_sSQL & " 	AND M01_DAIBUNRUI_CD = " & C_HISSEN

    '== �f�[�^�擾 ==
    If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function

	f_GetHissenNM = gf_SetNull2String(w_Rs("M01_SYOBUNRUIMEI"))

    '== ���� ==
    call gf_closeObject(w_Rs)

End Function

'2009/06/15 upd str iwata
'���E�����Ǝ��Ԃ̕\���̂��ߕύX
'********************************************************************************
'*  [�@�\]  ���Ǝ��Ԑ����Z�b�g
'********************************************************************************
'Sub s_GetJikan(p_sJ_JunJikan)
'
'	Select Case Cint(m_sSikenKBN)
'		Case C_SIKEN_ZEN_TYU
'			p_sJ_JunJikan = m_Rs("J_JUNJI_ZT")
'		Case Else
'			p_sJ_JunJikan = m_Rs("J_JUNJI_ZK")
'	End Select
'End Sub

'********************************************************************************
'*  [�@�\]  ���Ǝ��Ԑ����Z�b�g
'********************************************************************************
Sub s_GetJikan(p_sJ_JunJikan_Z,p_sJ_JunJikan_KT,p_sJ_JunJikan_KK)

	If Cint(m_sSikenKBN) = C_SIKEN_ZEN_TYU then
			p_sJ_JunJikan_Z = m_Rs("J_JUNJI_ZT")
	Else
			p_sJ_JunJikan_Z = m_Rs("J_JUNJI_ZK")
	End If
	p_sJ_JunJikan_KT = m_Rs("J_JUNJI_KT")
	p_sJ_JunJikan_KK = m_Rs("J_JUNJI_KK")

'2009/06/15 ins str iwata
	'1�l�ڂ̊w�����Ə��Ώۂ̊w���������ꍇ
	If m_iKamokuKbn = C_JIK_JUGYO then
'2009/10/05 upd iwata
'		If ( gf_SetNull2String(m_Rs("Menjo")) = "1" ) OR ( gf_SetNull2Zero(m_Rs("DataKbn")) <> 0 ) Then

'2009/10/05 DEBUG Yuki
'response.write "Menjo = " & gf_SetNull2String(m_Rs("Menjo")) & "<br>"
'response.write "Menjo2 = " & cstr(gf_SetNull2String(m_Rs("Menjo"))) & "<br>"
'response.write "DataKBN = " & gf_SetNull2Zero(m_Rs("DataKbn")) & "<br>"
'response.write "DataKBN2 = " & cstr(gf_SetNull2Zero(m_Rs("DataKbn"))) & "<br>"

		If ( cstr(gf_SetNull2String(m_Rs("Menjo"))) = "1" ) OR ( cstr(gf_SetNull2Zero(m_Rs("DataKbn"))) <> "0" ) Then

'response.write "Root = " & "A" & "<br>"

			'���C���Ă���w���̌���
			If Not m_Rs.EOF Then
				m_Rs.MoveNext
				Do Until m_Rs.EOF
					'�Ə��Ȗڂ�����
'2009/10/05 upd iwata
'					If ( gf_SetNull2String(m_Rs("Menjo")) <> "1" ) AND ( gf_SetNull2Zero(m_Rs("DataKbn")) = 0 ) Then
					If ( cstr(gf_SetNull2String(m_Rs("Menjo"))) <> "1" ) AND ( cstr(gf_SetNull2Zero(m_Rs("DataKbn"))) = "0" ) Then

'response.write "Root = " & "B" & "<br>"

						If Cint(m_sSikenKBN) = C_SIKEN_ZEN_TYU then
								p_sJ_JunJikan_Z = m_Rs("J_JUNJI_ZT")
						Else
								p_sJ_JunJikan_Z = m_Rs("J_JUNJI_ZK")
						End If
						p_sJ_JunJikan_KT = m_Rs("J_JUNJI_KT")
						p_sJ_JunJikan_KK = m_Rs("J_JUNJI_KK")

						'INSERT 2009/06/17
						'�\������X�V���t & ����
						Select Case Cint(m_sSikenKBN)
							Case C_SIKEN_ZEN_TYU : m_sUpdDate = f_fmtWareki(gf_SetNull2String(m_Rs("KOUSINBI_ZT"))) & "�@" & gf_SetNull2String(m_Rs("KOUSINTIME_ZT"))
							Case C_SIKEN_ZEN_KIM : m_sUpdDate = f_fmtWareki(gf_SetNull2String(m_Rs("KOUSINBI_ZK"))) & "�@" & gf_SetNull2String(m_Rs("KOUSINTIME_ZK"))
							Case C_SIKEN_KOU_TYU : m_sUpdDate = f_fmtWareki(gf_SetNull2String(m_Rs("KOUSINBI_KT"))) & "�@" & gf_SetNull2String(m_Rs("KOUSINTIME_KT"))
							Case C_SIKEN_KOU_KIM : m_sUpdDate = f_fmtWareki(gf_SetNull2String(m_Rs("KOUSINBI_KK"))) & "�@" & gf_SetNull2String(m_Rs("KOUSINTIME_KK"))
						End Select

						Exit Do
					End If
					m_Rs.MoveNext
				loop
			End If
			m_Rs.MoveFirst
		End If

'response.end

	End If
'2009/06/15 ins str iwata

End Sub

'********************************************************************************
'*  [�@�\]  �Ə��t���O�̃Z�b�g
'********************************************************************************
Sub s_SetMenjo(p_Menjo)

	p_Menjo = 0
	'//�����̏ꍇ��0��Ԃ�
	If m_iKamokuKbn = C_JIK_JUGYO Then
		p_Menjo = CInt(gf_SetNull2Zero(m_Rs("Menjo")))
	End If

End Sub

'********************************************************************************
'*  [�@�\]  HTML���o��
'********************************************************************************
Sub showPage()

	Dim w_sSeiseki
	Dim w_sHyoka
	Dim w_sKekka_ZK
	Dim w_sKekka_KT
	Dim w_sKekka_KK
	Dim w_sKibi_ZK
	Dim w_sKibi_KT
	Dim w_sKibi_KK
	Dim w_sTeisi_ZK
	Dim w_sTeisi_KT
	Dim w_sTeisi_KK
	Dim w_sHaken_ZK
	Dim w_sHaken_KT
	Dim w_sHaken_KK
	Dim w_sSeiseki_ZK
	Dim w_sSeiseki_KT
	Dim w_sSeiseki_KK
	Dim w_sHyoka_ZK
	Dim w_sHyoka_KT
	Dim w_sHyoka_KK
	Dim w_bNoChange_ZK
	Dim w_bNoChange_KT
	Dim w_bNoChange_KK
	Dim i
	Dim w_IdouKbn									'�ٓ��^�C�v
	Dim w_IdouName
	Dim w_IdouDate
	Dim w_sInputClass
	Dim w_Padding
	Dim w_cell
	Dim w_sJ_JunJikan_Z
	Dim w_sJ_JunJikan_KT	'2009/06/15 ins iwata
	Dim w_sJ_JunJikan_KK	'2009/06/15 ins iwata

	Dim w_Menjo	'2013/11/15 ins urakawa

	w_Padding   = "style='padding:2px 0px;font-size:10px;text-align:center'"
	w_Padding2  = "style='padding:2px 0px;font-size:10px;writing-mode:tb-rl'"
	w_Padding3  = "style='padding:2px 0px;font-size:10px'"

	i = 1

	'//�O������ or �O�������i�����敪�ɂ���ĕ���j�f�[�^�Z�b�g
'2009/06/15 upd iwata
'	Call s_GetJikan(w_sJ_JunJikan_Z)
	Call s_GetJikan(w_sJ_JunJikan_Z,w_sJ_JunJikan_KT,w_sJ_JunJikan_KK)

	'//NN�Ή�
	If session("browser") = "IE" Then
		w_sInputClass  = "class='num'"
	Else
		w_sInputClass = ""
	End If

%>
<html>
<head>
<link rel="stylesheet" href="../../common/style.css" type=text/css>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<!--#include file="../../Common/jsCommon.htm"-->
<!--OBJECT ID="thebrowser" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2" -->
<!--/OBJECT -->
<SCRIPT language="javascript">
<!--
	function window_onload(){
	alert("<%=C_TOUROKU_OK_MSG%>");
	window.focus();
	window.print();
	document.frm.target = "main";
	document.frm.action = "sei0150_23_bottom.asp"
	document.frm.submit();
	}
//-->
</SCRIPT>
<style TYPE="text/css">
table.hyo1 {
	border-layout : fixed;
	border-collapse:collapse;
	border-style:solid;
	border-width:1px;
	padding:0px;
	margin:0px;
}
td.head1 {
	font-size:8pt;
	padding:2px 5px;
}
td.head2 {
	font-size:8pt;
	padding:2px 5px;
	writing-mode:tb-rl;
}
td.head3 {
	font-size:10pt;
	padding:2px 5px;
}
p.margin1 {
	margin: 0px 0px 0px 0px
}
<!--
	@media screen,print{
		BODY {
			margin: 0;  ?*�u���b�N�̈�ƃu���b�N�g�̗]�����[���w��*?
			padding: 0; ?*�u���b�N�g�ƃu���b�N�����̗]�����[���w��*?
		}
	}
//-->
</style>
</head>
<body LANGUAGE="javascript" onload="window_onload();">
<center>
<form name="frm" method="post">
	<p class="margin1"></p>
	<table aling="center">
		<tr>
			<td class="head1" colspan="3" height="10"></td>
			<th width="350" align="center" rowspan="2"><font size="5">���@�с@�]�@���@�\</font></th>
			<td class="head1"></td>
		</tr>
		<tr>
			<td class="head3" aling="right">����</td>
			<td class="head3" aling="center"><%=f_Nendo2Wareki(m_iNendo)%></td>
			<td class="head3">�N�x</td>
			<td class="head3"><%=f_GetSchoolName%></td>
		</tr>
	</table>
	<table aling="center" cellpadding="0" cellspacing="0">
		<tr height="5">
			<td></td>
		</tr>
	</table>
	<table aling="center" cellpadding="0" cellspacing="0">
		<tr>
			<td class="head3" width="140" align="center"><%=gf_GetClassName(m_iNendo,m_iGakunen,m_sClassNo)%></td>
			<td class="head3" width="140" align="center">�� <%=m_iGakunen%> �w�N</td>
			<td class="head3" width="140" align="center"><%=f_ShikenMei%></td>
			<td class="head3" width="230" align="right"><%=m_sUpdDate%>�@�o�^</td>
		</tr>
	</table>




	<table>
		<tr>
			<td>
				<table class="hyo1" align="center" border="1">
					<tr>
						<td class="head1" colspan="3"  align="center" nowrap>���Ȗ�</td>
						<td class="head1" colspan="2"  align="center" nowrap>�P�ʐ�</td>
						<!-- 2011.06.06 upd iwata�@�S���������� => �S�����������@-->
						<td class="head1" colspan="19" align="center" nowrap>�S����������&nbsp;&nbsp;&nbsp;&nbsp;<%=Session("USER_NM")%></td>
					</tr>
					<tr>
						<td class="head1" colspan="3" rowspan="2"  align="center" nowrap><%=gf_GetKamokuMei(m_iNendo,m_sKamokuCd,m_iKamokuKbn)%></td>
						<td class="head1" colspan="2" align="center" nowrap><%=f_GetHissenNM(m_iHissen_Kbn)%></td>
						<td class="head2" rowspan="2" align="center" nowrap>�O��</td>
						<td class="head1" colspan="5" rowspan="2"  align="center" nowrap>
							<table>
								<tr>
									<td class="head1" width="85" align="center" nowrap><!-- <%=Session("USER_NM")%> --></td>
									<td class="head1" align="right" nowrap><!-- �� --></td>
								</tr>
							</table>
						</td>
						<td class="head2" rowspan="2" align="center" nowrap>���</td>
						<td class="head1" colspan="5" rowspan="2"  align="center" nowrap>
							<table>
								<tr>
									<td class="head1" width="85" align="center" nowrap><!-- <%=Session("USER_NM")%> --></td>
									<td class="head1" align="right" nowrap><!-- �� --></td>
								</tr>
							</table>
						</td>
						<td class="head2" rowspan="2" align="center" nowrap>�w�N</td>
						<td class="head1" colspan="5" rowspan="2"  align="center" nowrap>
							<table>
								<tr>
									<td class="head1" width="85" align="center" nowrap><!-- <%=Session("USER_NM")%> --></td>
									<td class="head1" align="right" nowrap><!-- �� --></td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td class="head1" colspan="2" align="center" nowrap><%=gf_SetNull2String(m_Rs("HAITOTANI"))%></td>
					</tr>
					<tr>
						<td class="head2" rowspan="4" align="center" width="15"  nowrap>����ԍ�</td>
						<td class="head1" colspan="2" align="center" nowrap>��</td>
						<td class="head1" colspan="4" align="center" nowrap>�O�@��</td>
						<td class="head1" colspan="4" align="center" nowrap>��@��</td>
						<td class="head1" colspan="4" align="center" nowrap>�w�@�N</td>
						<td class="head1" colspan="3" rowspan="2" align="center" nowrap>���ѕ]��</td>
						<td class="head1" colspan="5" rowspan="4" align="center" nowrap>���@�l</td>
					</tr>
					<tr>
						<td class="head1" colspan="2" align="center" nowrap>���Ǝ���</td>
						<td class="head1" colspan="4" align="center" nowrap><%=gf_SetNull2String(w_sJ_JunJikan_Z)%> ����</td>
<% '2009/06/15 upd iwata ���Ǝ����̕\���ύX�@�Ə��Ȗڂ̊w�����ǂ����`�F�b�N����
   '	m_Rs("J_JUNJI_KT")=>w_sJ_JunJikan_KT,m_Rs("J_JUNJI_KK")=>w_sJ_JunJikan_KK �ύX %>
						<td class="head1" colspan="4" align="center" nowrap><%=gf_SetNull2String(w_sJ_JunJikan_KT)%> ����</td>
						<!-- 2011.06.06 ins iwata '()' �ǉ� -->
						<td class="head1" colspan="4" align="center" nowrap><%=gf_SetNull2String(w_sJ_JunJikan_KK)%> ����(&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)</td>
					</tr>
					<tr>
						<td class="head1" colspan="2" align="center" nowrap>���ێ���</td>
						<td class="head2" align="center" rowspan="2" nowrap>����</td>
						<td class="head2" align="center" rowspan="2" nowrap>����</td>
						<td class="head2" align="center" rowspan="2" nowrap>��~</td>
						<td class="head2" align="center" rowspan="2" nowrap>�h��</td>
						<td class="head2" align="center" rowspan="2" nowrap>����</td>
						<td class="head2" align="center" rowspan="2" nowrap>����</td>
						<td class="head2" align="center" rowspan="2" nowrap>��~</td>
						<td class="head2" align="center" rowspan="2" nowrap>�h��</td>
						<td class="head2" align="center" rowspan="2" nowrap>����</td>
						<td class="head2" align="center" rowspan="2" nowrap>����</td>
						<td class="head2" align="center" rowspan="2" nowrap>��~</td>
						<td class="head2" align="center" rowspan="2" nowrap>�h��</td>
						<td class="head2" align="center" rowspan="2" nowrap>�O��</td>
						<td class="head2" align="center" rowspan="2" nowrap>���</td>
						<td class="head2" align="center" rowspan="2" nowrap>�w�N</td>
					</tr>
					<tr>
						<td class="head1" width="55" align="center" nowrap>�w���ԍ�</td>
						<td class="head1" width="90" align="center" nowrap>�w�@���@���@��</td>
					</tr>

				<%
					m_Rs.MoveFirst
					Do Until m_Rs.EOF

						j = j + 1

						w_sKekka_ZK = ""
						w_sKekka_KT = ""
						w_sKekka_KK = ""
						w_sKibi_ZK  = ""
						w_sKibi_KT  = ""
						w_sKibi_KK  = ""
						w_sTeisi_ZK = ""
						w_sTeisi_KT = ""
						w_sTeisi_KK = ""
						w_sHaken_ZK = ""
						w_sHaken_KT = ""
						w_sHaken_KK = ""
						w_sSeiseki  = ""
						w_sHyoka    = ""
						w_bNoChange = false

						Call gs_cellPtn(w_cell)

						'//���ہA�x�����̃Z�b�g
						Call s_SetKekka(w_sKekka_ZK, w_sKekka_KT, w_sKekka_KK, _
										w_sKibi_ZK , w_sKibi_KT , w_sKibi_KK, _
										w_sTeisi_ZK, w_sTeisi_KT, w_sTeisi_KK, _
										w_sHaken_ZK, w_sHaken_KT, w_sHaken_KK)

						'//���уf�[�^�Z�b�g
						Call s_SetGrades(w_sSeiseki_ZK, w_sSeiseki_KT, w_sSeiseki_KK, _
										 w_sHyoka_ZK, w_sHyoka_KT, w_sHyoka_KK, _
										 w_bNoChange_ZK, w_bNoChange_KT, w_bNoChange_KK)

						'//�ٓ��`�F�b�N
						Call s_IdouCheck(m_Rs("GAKUSEKI_NO"),w_IdouKbn,w_IdouName,w_bNoChange_ZK, w_bNoChange_KT, w_bNoChange_KK,w_IdouDate)

						'2013/11/15 ins urakawa
						'//�Ə��t���O�̃Z�b�g
						Call s_SetMenjo(w_Menjo)
						'2013/11/15 end urakawa

				%>
					<tr>
						<!--<td class="<%=w_cell%>" align="center" nowrap <%=w_Padding3%>><%=i%></td>-->
						<td class="<%=w_cell%>" align="center" nowrap <%=w_Padding3%>><%=m_Rs("SYUSEKI_NO")%></td>	<!--2018.02.09 Upd Kiyomoto �A�Ԃ��N���X�o�Ȕԍ��ɕύX-->
						<td class="<%=w_cell%>" align="center" width="55"  nowrap <%=w_Padding3%>><%=m_Rs("GAKUSEI_NO")%></td>
						<td class="<%=w_cell%>" align="left"   width="90" nowrap <%=w_Padding3%>><%=trim(m_Rs("SIMEI"))%><%=w_IdouName%></td>

						<!-- ���� -->
						<!-- �O������ -->
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sKekka_ZK%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sKibi_ZK%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sTeisi_ZK%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sHaken_ZK%></td>
						<!-- ������� -->
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sKekka_KT%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sKibi_KT%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sTeisi_KT%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sHaken_KT%></td>
						<!-- �w�N�� -->
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sKekka_KK%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sKibi_KK%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sTeisi_KK%></td>
						<td class="<%=w_cell%>" align="center" width="25"  nowrap <%=w_Padding%>><%=w_sHaken_KK%></td>

						<!--�I���Ȗڂ̎��ɖ��I���̏ꍇ�A���͕s�B�܂��A�x�w�Ȃ�-->
						<% If w_bNoChange_ZK = True Then %>
							<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>>-</td>
							<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>>-</td>
							<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>>-</td>

						<!-- ���� (���l���́A�������́A���тȂ����͂ɂ�菈���𕪂���) -->
						<% Else %>
							<!-- ���l���� -->
							<% if m_iSeisekiInpType = C_SEISEKI_INP_TYPE_NUM then %>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>><%=w_sSeiseki_ZK%></td>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>><%=w_sSeiseki_KT%></td>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>><%=w_sSeiseki_KK%></td>

							<!-- �������� -->
							<% elseif m_iSeisekiInpType = C_SEISEKI_INP_TYPE_STRING then %>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>><%=w_sSeiseki_ZK%></td>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>><%=w_sSeiseki_KT%></td>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>><%=w_sSeiseki_KK%></td>

							<!-- �ȊO -->
							<% else %>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>>-</td>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>>-</td>
								<td class="<%=w_cell%>" align="center" width="25" nowrap <%=w_Padding%>>-</td>
							<% end if
						End If %>

						<!-- upd str 2013.11.15 urakawa -->
						<!-- <td class="<%=w_cell%>" width="125" colspan="5" nowrap <%=w_Padding%>><%=w_IdouDate%></td> -->
						<% If w_Menjo = 1 Then %>
							<td class="<%=w_cell%>" width="125" colspan="5" nowrap <%=w_Padding%>><%=w_IdouDate & " " & "�y�C���ρz"%></td>
						<% else %>
							<td class="<%=w_cell%>" width="125" colspan="5" nowrap <%=w_Padding%>><%=w_IdouDate%></td>
						<% end if %>
						<!-- upd end 2013.11.15 urakawa -->

					</tr>
					<%

							if (i Mod 5) = 0 then
								Response.write "<tr>"
									Response.write "<td colspan='23'>"
									Response.write "</td>"
								Response.write "</tr>"
							end if

							m_Rs.MoveNext
							i = i + 1
						Loop
					%>

				</table>
			</td>
		</tr>
		<!-- ins str 2011.06.06 iwata -->
		<tr>
		<!-- upd str 2013.06.17 yuki -->
		<!-- <td><font size="3">&nbsp;&nbsp;���w�N���Ǝ�������( )�ɂ́A�����������Ԑ������������Ǝ��Ԑ����L�����Ă��������B</font></td> -->
			<td><font size="2">&nbsp;&nbsp;���w�N���Ǝ��Ԑ�����( )���̎��Ԑ��̋L���ɂ��ẮA�����蒠�u���ѕ񍐕��@�}�j���A���v���Q�Ƃ��Ă��������B </font></td>
		<!-- upd end 2013.06.17 yuki -->
		</tr>
		<!-- ins end 2011.06.06 iwata -->
	</table>

	<input type="hidden" name="txtNendo"     value="<%=trim(Request("txtNendo"))%>">
	<input type="hidden" name="txtKyokanCd"  value="<%=trim(Request("txtKyokanCd"))%>">
	<input type="hidden" name="sltShikenKbn" value="<%=trim(Request("sltShikenKbn"))%>">
	<input type="hidden" name="txtGakuNo"    value="<%=trim(Request("txtGakuNo"))%>">
	<input type="hidden" name="txtClassNo"   value="<%=trim(Request("txtClassNo"))%>">
	<input type="hidden" name="txtKamokuCd"  value="<%=trim(Request("txtKamokuCd"))%>">
	<input type="hidden" name="txtGakkaCd"   value="<%=trim(Request("txtGakkaCd"))%>">
	<input type="hidden" name="hidKamokuKbn" value="<%=request("hidKamokuKbn")%>">
	<input type="hidden" name="hidSyubetu"   value="<%=request("hidSyubetu")%>">

</form>
</center>
</body>
</html>
<%
End sub
%>