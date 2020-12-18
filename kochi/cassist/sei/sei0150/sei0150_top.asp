<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ѓo�^
' ��۸���ID : sei/sei0150/sei0150_top.asp
' �@      �\: ��y�[�W ���ѓo�^�̌������s��
'-------------------------------------------------------------------------
' ��      ��:
'           :
' ��      ��:
' ��      �n:
'           :
' ��      ��:
'           �������\��
'               �R���{�{�b�N�X�͋󔒂ŕ\��
'           ���\���{�^���N���b�N��
'               ���̃t���[���Ɏw�肵�������ɂ��Ȃ��������̓��e��\��������
'-------------------------------------------------------------------------
' ��      ��: 2002/06/20 shin
' ��      �X: 2010/05/20 ��c�@���m����@�����N���X�A��ʁA�w��Ȗڂ��w�ȕ\���ɂ���
' ��      �X: 2016/06/14 ���{�@���m����@���Ȗڂ��N���X�\���ɂ���
' ��      �X: 2017/12/12 �����@���m����@�����N���X�A��ʁA�w��Ȗڂ��w�ȕ\���ɂ����140050�����͏��O����
' ��      �X: 2018/05/08 ���{�@���m����@1�w�ȕ����R�[�X�Ή�
' ��      �X: 2018/06/18 ���{�@���m����@�l���C�ǉ��Ȗڂ��J�ݎ����ɂ���ĉȖڂ𐧌䂷��
' ��      �X: 2018/06/24 ���с@���m����@�ȖڃR���{�́A�N���X�A�w�ȁA�Ȗڂ��Ƃɕ\������悤�ɕύX�BVIEW���g�p����悤�ɕύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Dim  m_bErrFlg           '�װ�׸�

	Dim m_iNendo             '�N�x
	Dim m_sKyokanCd          '�����R�[�h
	Dim m_iSikenKbn			'�����敪

	Dim gDisabled

	Dim gRs

	Public m_sGakkoNO       '�w�Z�ԍ�

	Dim m_iGakunen
	Dim m_iClass
	Dim m_iKongo

'///////////////////////////���C������/////////////////////////////

	Call Main()

'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Sub Main()
	Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="���ѓo�^"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = false

    Do
		'//�ް��ް��ڑ�
		If gf_OpenDatabase() <> 0 Then
			m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
			Exit Do
		End If

		'//�l���擾
		call s_SetParam()

		'//�����敪
		if request("sltShikenKbn")  = "" then
			'//�ŏ�
			w_iRet = gf_Get_SikenKbn(m_iSikenKbn,C_SEISEKI_KIKAN,0)

			if w_iRet <> 0 then m_bErrFlg = true : exit do
		else
		    '//�����[�h��
		    m_iSikenKbn = cint(Request("sltShikenKbn"))
		end if

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

		'�w�Z�ԍ��̎擾
		if Not gf_GetGakkoNO(m_sGakkoNO) then Exit Do

		'//���O�C�������̒S���Ȗڂ̎擾
		if not f_GetSubject() then Exit Do

		'�Ȗڃf�[�^�Ȃ�
		if gRs.EOF Then
			gDisabled = "disabled"
		'	Call showWhitePage("�S���Ȗڃf�[�^������܂���")
		'	response.end
		End If

		'// �y�[�W��\��
		Call showPage()

		m_bErrFlg = true
		Exit Do
	Loop

	'// �װ�̏ꍇ�ʹװ�߰�ނ�\��
	If not m_bErrFlg Then
		w_sMsg = gf_GetErrMsg()
		Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
	End If

	'// �I������
	Call gf_closeObject(gRs)
	Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Sub s_SetParam()

	gDisabled = ""

    m_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")

End Sub


'********************************************************************************
'*  [�@�\]  ���O�C�������̎󎝋��Ȃ��擾(�N�x�A����CD�A�w�����)
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]
'********************************************************************************
Function f_GetSubject()
	Dim w_sSQL
    Dim w_sJiki
    Dim w_iCnt

    On Error Resume Next
    Err.Clear

    f_GetSubject = false

	'�I�񂾎����ɂ���āA�J�݊��Ԃ�ς���
	Select Case cint(m_iSikenKbn)
		Case C_SIKEN_ZEN_TYU : w_sJiki = C_KAI_ZENKI	'�O������
		case C_SIKEN_ZEN_KIM : w_sJiki = C_KAI_ZENKI	'�O������
		Case C_SIKEN_KOU_TYU : w_sJiki = C_KAI_KOUKI	'�������
        case C_SIKEN_KOU_KIM : w_sJiki = C_KAI_KOUKI	'�������
	End Select

	'--2019/06/24 Upd Fujibayashi(���̉�ʂƂ̐�������ۂ��߁AVIEW���g�p����)
	''�ʏ�A���w����։Ȗڎ擾
	'w_sSQL = ""
	'w_sSQL = w_sSQL & " select distinct "
	'w_sSQL = w_sSQL & "		T27_GAKUNEN as GAKUNEN "
	'w_sSQL = w_sSQL & "		,T27_CLASS as CLASS "
	'w_sSQL = w_sSQL & "		,T27_KAMOKU_CD as KAMOKU_CD "
	'w_sSQL = w_sSQL & "		,M03_KAMOKUMEI as KAMOKU_NAME "
	'w_sSQL = w_sSQL & "		,M03_KAMOKU_KBN as KAMOKU_KBN_IS "
	'w_sSQL = w_sSQL & "		,T27_KAMOKU_BUNRUI as KAMOKU_KBN "
	'w_sSQL = w_sSQL & "		,M05_CLASSMEI as CLASS_NAME "
	'w_sSQL = w_sSQL & "		,M05_GAKKA_CD as GAKKA_CD "
	'w_sSQL = w_sSQL & " from"
	'w_sSQL = w_sSQL & "		T27_TANTO_KYOKAN "
	'w_sSQL = w_sSQL & "		,M03_KAMOKU "
	'w_sSQL = w_sSQL & "		,T15_RISYU "
	'w_sSQL = w_sSQL & "		,M05_CLASS "
	'w_sSQL = w_sSQL & " where "
	'w_sSQL = w_sSQL & "		T27_NENDO =" & cint(m_iNendo)
	'w_sSQL = w_sSQL & "	and	T27_KYOKAN_CD ='" & m_sKyokanCd & "'"
	'w_sSQL = w_sSQL & "	and	T27_KAMOKU_CD = M03_KAMOKU_CD "
	'w_sSQL = w_sSQL & "	and	T27_KAMOKU_BUNRUI   = " & C_JIK_JUGYO
	'w_sSQL = w_sSQL & "	and	T27_SEISEKI_INP_FLG = " & C_SEISEKI_INP_FLG_YES
	'w_sSQL = w_sSQL & "	and	T27_KAMOKU_CD = T15_KAMOKU_CD(+) "
	'w_sSQL = w_sSQL & "	and	T15_NYUNENDO = T27_NENDO - T27_GAKUNEN + 1 "
	'w_sSQL = w_sSQL & "	and	T15_GAKKA_CD = M05_GAKKA_CD"	'2003.01.29
	'w_sSQL = w_sSQL & "	and T15_COURSE_CD IN ('0', CASE WHEN M05_COURSE_CD IS NOT NULL THEN M05_COURSE_CD ELSE T15_COURSE_CD END) "	'2018.05.08 Add Kiyomoto
	'w_sSQL = w_sSQL & " and T27_NENDO = M05_NENDO "
	'w_sSQL = w_sSQL & " and T27_GAKUNEN = M05_GAKUNEN "
	'w_sSQL = w_sSQL & " and T27_CLASS = M05_CLASSNO "
	'w_sSQL = w_sSQL & "	and	M03_NENDO =" & cint(m_iNendo)
    '
	''���㍂��̏ꍇ�A������Ԃ̎��O���J�݂̉Ȗڂ��\������
    'if m_sGakkoNO = cstr(C_NCT_YATSUSIRO) then
    '
	'	'//������ԁA���������������Ȃ��Ƃ�
	'	if m_iSikenKbn = C_SIKEN_KOU_KIM Or m_iSikenKbn = C_SIKEN_KOU_TYU then
	'		'w_sSQL = w_sSQL & " and	("
	'		'w_sSQL = w_sSQL & "			T15_KAISETU1 <" & C_KAI_NASI & " or "
	'		'w_sSQL = w_sSQL & "			T15_KAISETU2 <" & C_KAI_NASI & " or "
	'		'w_sSQL = w_sSQL & "			T15_KAISETU3 <" & C_KAI_NASI & " or "
	'		'w_sSQL = w_sSQL & "			T15_KAISETU4 <" & C_KAI_NASI & " or "
	'		'w_sSQL = w_sSQL & "			T15_KAISETU5 <" & C_KAI_NASI
	'		'w_sSQL = w_sSQL & "		 )"
    '
	'		'�w�N�ɍ��킹���J�ݎ����������ɂ��� 2003.01.29
	'		w_sSQL = w_sSQL & " and DECODE(T27_GAKUNEN "
	'		w_sSQL = w_sSQL & " ,1, T15_KAISETU1 "
	'		w_sSQL = w_sSQL & " ,2, T15_KAISETU2 "
	'		w_sSQL = w_sSQL & " ,3, T15_KAISETU3 "
	'		w_sSQL = w_sSQL & " ,4, T15_KAISETU4 "
	'		w_sSQL = w_sSQL & " ,5, T15_KAISETU5 "
	'		w_sSQL = w_sSQL & " ) < " & C_KAI_NASI
    '
	'	else
	'		w_sSQL = w_sSQL & " and	("
	'		for  w_iCnt = 1 to 5
	'		     w_sSQL = w_sSQL & "((T15_KAISETU" & w_iCnt & " = " & w_sJiki & "  "
	'             w_sSQL = w_sSQL & "or  T15_KAISETU" & w_iCnt & " = " &  C_KAI_TUNEN & ")  "
	'             w_sSQL = w_sSQL & " and  T27_GAKUNEN = " & w_iCnt  & ")  "
    '
	'		     if w_iCnt <> 5 then
	'		     	  w_sSQL = w_sSQL & " or  "
	'                     end if
	'		next
	'		     w_sSQL = w_sSQL & " ) "
    '
	'	end if
    '
	''���̑��̊w�Z�͊w�N�������̎������O���J�݂̉Ȗڂ�\������
	'else
	'	'�V���l����̏ꍇ�A��������͑O���Ȗڂ͕\�����Ȃ�
	'    if m_sGakkoNO = cstr(C_NCT_NIIHAMA) or m_sGakkoNO = cstr(C_NCT_KURUME) then
	'			w_sSQL = w_sSQL & " and	("
	'			for  w_iCnt = 1 to 5
	'			     w_sSQL = w_sSQL & "((T15_KAISETU" & w_iCnt & " = " & w_sJiki & "  "
	'                 w_sSQL = w_sSQL & "or  T15_KAISETU" & w_iCnt & " = " &  C_KAI_TUNEN & ")  "
	'                 w_sSQL = w_sSQL & " and  T27_GAKUNEN = " & w_iCnt  & ")  "
    '
	'			     if w_iCnt <> 5 then
	'			     	  w_sSQL = w_sSQL & " or  "
	'	              end if
	'			next
    '
	'			w_sSQL = w_sSQL & " ) "
    '    else
	'		'//���������������Ȃ��Ƃ�
	'		if m_iSikenKbn <> C_SIKEN_KOU_KIM then
	'			w_sSQL = w_sSQL & " and	("
	'			for  w_iCnt = 1 to 5
	'			     w_sSQL = w_sSQL & "((T15_KAISETU" & w_iCnt & " = " & w_sJiki & "  "
	'                 w_sSQL = w_sSQL & "or  T15_KAISETU" & w_iCnt & " = " &  C_KAI_TUNEN & ")  "
	'                 w_sSQL = w_sSQL & " and  T27_GAKUNEN = " & w_iCnt  & ")  "
    '
	'			     if w_iCnt <> 5 then
	'			     	  w_sSQL = w_sSQL & " or  "
	'	              end if
	'			next
    '
	'			w_sSQL = w_sSQL & " ) "
    '
	'		else
    '
	'			'�w�N�ɍ��킹���J�ݎ����������ɂ��� 2003.01.29
	'			w_sSQL = w_sSQL & " and DECODE(T27_GAKUNEN "
	'			w_sSQL = w_sSQL & " ,1, T15_KAISETU1 "
	'			w_sSQL = w_sSQL & " ,2, T15_KAISETU2 "
	'			w_sSQL = w_sSQL & " ,3, T15_KAISETU3 "
	'			w_sSQL = w_sSQL & " ,4, T15_KAISETU4 "
	'			w_sSQL = w_sSQL & " ,5, T15_KAISETU5 "
	'			w_sSQL = w_sSQL & " ) < " & C_KAI_NASI
    '
	'		end if
	'    end if
	'end if
    '
    '
	'w_sSQL = w_sSQL & vbCrLf & "  UNION ALL "
    '
	'w_sSQL = w_sSQL & " SELECT DISTINCT "
	'w_sSQL = w_sSQL & " 	T27_GAKUNEN AS GAKUNEN"
	'w_sSQL = w_sSQL & " 	,T27_CLASS AS CLASS"
	'w_sSQL = w_sSQL & " 	,T27_KAMOKU_CD AS KAMOKU_CD "
	'w_sSQL = w_sSQL & " 	,T16_KAMOKUMEI AS KAMOKU_NAME "
	'w_sSQL = w_sSQL & "		,T16_KAMOKU_KBN as KAMOKU_KBN_IS "
	'w_sSQL = w_sSQL & "		,T27_KAMOKU_BUNRUI as KAMOKU_KBN "
	'w_sSQL = w_sSQL & " 	,M05_CLASSMEI AS CLASS_NAME "
	'w_sSQL = w_sSQL & " 	,M05_GAKKA_CD AS GAKKA_CD "
	'w_sSQL = w_sSQL & " FROM"
	'w_sSQL = w_sSQL & " 	T27_TANTO_KYOKAN "
	'w_sSQL = w_sSQL & " 	,T16_RISYU_KOJIN "
	'w_sSQL = w_sSQL & " 	,M05_CLASS "
	'w_sSQL = w_sSQL & " WHERE "
	'w_sSQL = w_sSQL & " 		T27_NENDO = M05_NENDO "
	'w_sSQL = w_sSQL & "    AND T27_GAKUNEN = M05_GAKUNEN "
	'w_sSQL = w_sSQL & "    AND T27_GAKUNEN = T16_HAITOGAKUNEN "
	'w_sSQL = w_sSQL & "    AND T27_CLASS = M05_CLASSNO	"
	'w_sSQL = w_sSQL & "    AND T27_KAMOKU_CD = T16_KAMOKU_CD(+)"
	'w_sSQL = w_sSQL & "    AND M05_GAKKA_CD(+) = T16_GAKKA_CD "
	'w_sSQL = w_sSQL & "    AND T16_NENDO(+) = T27_NENDO "
	'w_sSQL = w_sSQL & "    AND T27_NENDO = " & m_iNendo
	'w_sSQL = w_sSQL & "    AND T27_KYOKAN_CD ='" & m_sKyokanCd & "' "
	'w_sSQL = w_sSQL & "    AND T27_SEISEKI_INP_FLG =" & C_SEISEKI_INP_FLG_YES & " "
	'w_sSQL = w_sSQL & "    AND T16_OKIKAE_FLG >= " & C_TIKAN_KAMOKU_SAKI
    '
	''INS 2008/09/11
	''�V���l����̏ꍇ�A�O���́u�O���E�ʔN�v����́u����E�ʔN�v�Ƃ���
	''2018.06.18 Add ���m����̏ꍇ�����l
    'if m_sGakkoNO = cstr(C_NCT_NIIHAMA) or m_sGakkoNO = cstr(C_NCT_KOCHI) then
	'	w_sSQL = w_sSQL & "    AND ((T16_KAISETU = " & w_sJiki & "  "
	'	w_sSQL = w_sSQL & "    or  T16_KAISETU = " &  C_KAI_TUNEN & "))  "
	'end if
	''INS END 2008/09/11
    '
    '
    '
	'w_sSQL = w_sSQL & " union "
    '
	''T27��T27_SEISEKI_INP_FLG=1�̐l�����̐搶�ɐ��ѓo�^���������ہA
	''T26�Ƀf�[�^�����邽�߁A�ȉ���SQL�����s����K�v������
	'w_sSQL = w_sSQL & " SELECT distinct "
	'w_sSQL = w_sSQL & "		T26_GAKUNEN AS GAKUNEN "
	'w_sSQL = w_sSQL & "		,T26_CLASS AS CLASS "
	'w_sSQL = w_sSQL & "		,T26_KAMOKU AS KAMOKU_CD "
	'w_sSQL = w_sSQL & "		,M03_KAMOKUMEI as KAMOKU_NAME "
	'w_sSQL = w_sSQL & "		,M03_KAMOKU_KBN as KAMOKU_KBN_IS "
	'w_sSQL = w_sSQL & "		,0 as KAMOKU_KBN "
	'w_sSQL = w_sSQL & "		,M05_CLASSMEI as CLASS_NAME "
	'w_sSQL = w_sSQL & "		,M05_GAKKA_CD as GAKKA_CD "
	'w_sSQL = w_sSQL & " FROM "
	'w_sSQL = w_sSQL & "		T26_SIKEN_JIKANWARI "
	'w_sSQL = w_sSQL & "		,M03_KAMOKU "
	'w_sSQL = w_sSQL & "		,T15_RISYU "
	'w_sSQL = w_sSQL & "		,M05_CLASS "
	'w_sSQL = w_sSQL & " WHERE "
	'w_sSQL = w_sSQL & "		 T26_NENDO = " & cint(m_iNendo)
	'w_sSQL = w_sSQL & "	and ("
	'w_sSQL = w_sSQL & "		T26_JISSI_KYOKAN    ='" & m_sKyokanCd & "'"
	'w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN1 ='" & m_sKyokanCd & "'"
	'w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN2 ='" & m_sKyokanCd & "'"
	'w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN3 ='" & m_sKyokanCd & "'"
	'w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN4 ='" & m_sKyokanCd & "'"
	'w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN5 ='" & m_sKyokanCd & "'"
	'w_sSQL = w_sSQL & "		)"
	'w_sSQL = w_sSQL & "	and	T26_KAMOKU = M03_KAMOKU_CD "
	'w_sSQL = w_sSQL & "	and T26_KAMOKU = T15_KAMOKU_CD(+) "
	'w_sSQL = w_sSQL & "	and T15_NYUNENDO(+) = T26_NENDO - T26_GAKUNEN + 1 "
	'w_sSQL = w_sSQL & "	and	T15_GAKKA_CD = M05_GAKKA_CD"	'2003.01.29
	'w_sSQL = w_sSQL & "	and T26_SIKEN_CD ='" & C_SIKEN_CODE_NULL & "' "
	'w_sSQL = w_sSQL & "	and	T26_SEISEKI_INP_FLG = " & C_SEISEKI_INP_FLG_YES
	'w_sSQL = w_sSQL & "	and	M03_NENDO =" & cint(m_iNendo)
	'w_sSQL = w_sSQL & " and T26_NENDO = M05_NENDO "
	'w_sSQL = w_sSQL & " and T26_GAKUNEN = M05_GAKUNEN "
	'w_sSQL = w_sSQL & " and T26_CLASS = M05_CLASSNO "
    '
	''���㍂��̏ꍇ�A������Ԃ̎��O���J�݂̉Ȗڂ��\������
    'if m_sGakkoNO = cstr(C_NCT_YATSUSIRO) then
	'	'//������ԁA���������������Ȃ��Ƃ�
	'	if m_iSikenKbn = C_SIKEN_KOU_KIM Or m_iSikenKbn = C_SIKEN_KOU_TYU then
	'	else
	'		w_sSQL = w_sSQL & "	and T26_SIKEN_KBN =" & m_iSikenKbn
	'	end if
	'else
	''�V���l����̏ꍇ�A��������͑O���Ȗڂ͕\�����Ȃ�
	'    if m_sGakkoNO = cstr(C_NCT_NIIHAMA) or m_sGakkoNO = cstr(C_NCT_KURUME) then
	'			w_sSQL = w_sSQL & "	and T26_SIKEN_KBN =" & m_iSikenKbn
	'    else
	'		'//���������������Ȃ��Ƃ�
	'		if m_iSikenKbn <> C_SIKEN_KOU_KIM then
	'			w_sSQL = w_sSQL & "	and T26_SIKEN_KBN =" & m_iSikenKbn
	'		end if
	'	end if
    'end if
    '
	''�J�ݎ��������̒ǉ� 2003.01.29
	''���㍂��̏ꍇ�A������Ԃ̎��O���J�݂̉Ȗڂ��\������
    'if m_sGakkoNO = cstr(C_NCT_YATSUSIRO) then
    '
	'	'//������ԁA���������������Ȃ��Ƃ�
	'	if m_iSikenKbn = C_SIKEN_KOU_KIM Or m_iSikenKbn = C_SIKEN_KOU_TYU then
    '
	'		'�w�N�ɍ��킹���J�ݎ����������ɂ���
	'		w_sSQL = w_sSQL & " and DECODE(T26_GAKUNEN "
	'		w_sSQL = w_sSQL & " ,1, T15_KAISETU1 "
	'		w_sSQL = w_sSQL & " ,2, T15_KAISETU2 "
	'		w_sSQL = w_sSQL & " ,3, T15_KAISETU3 "
	'		w_sSQL = w_sSQL & " ,4, T15_KAISETU4 "
	'		w_sSQL = w_sSQL & " ,5, T15_KAISETU5 "
	'		w_sSQL = w_sSQL & " ) < " & C_KAI_NASI
    '
	'	else
    '
	'		'�w�N�ɍ��킹���J�ݎ����������ɂ���
	'		w_sSQL = w_sSQL & " and DECODE(T26_GAKUNEN "
	'		w_sSQL = w_sSQL & " ,1, T15_KAISETU1 "
	'		w_sSQL = w_sSQL & " ,2, T15_KAISETU2 "
	'		w_sSQL = w_sSQL & " ,3, T15_KAISETU3 "
	'		w_sSQL = w_sSQL & " ,4, T15_KAISETU4 "
	'		w_sSQL = w_sSQL & " ,5, T15_KAISETU5 "
	'		w_sSQL = w_sSQL & ") IN (" & w_sJiki & "," & C_KAI_TUNEN & ")"
    '
    '
	'	end if
    '
	''���̑��̊w�Z�͊w�N�������̎������O���J�݂̉Ȗڂ�\������
	'else
	''�V���l����̏ꍇ�A��������͑O���Ȗڂ͕\�����Ȃ�
	'    if m_sGakkoNO = cstr(C_NCT_NIIHAMA) or m_sGakkoNO = cstr(C_NCT_KURUME) then
    '
	'			'�w�N�ɍ��킹���J�ݎ����������ɂ���
	'			w_sSQL = w_sSQL & " and DECODE(T26_GAKUNEN "
	'			w_sSQL = w_sSQL & " ,1, T15_KAISETU1 "
	'			w_sSQL = w_sSQL & " ,2, T15_KAISETU2 "
	'			w_sSQL = w_sSQL & " ,3, T15_KAISETU3 "
	'			w_sSQL = w_sSQL & " ,4, T15_KAISETU4 "
	'			w_sSQL = w_sSQL & " ,5, T15_KAISETU5 "
	'			w_sSQL = w_sSQL & ") IN (" & w_sJiki & "," & C_KAI_TUNEN & ")"
    '    else
	'		'//���������������Ȃ��Ƃ�
	'		if m_iSikenKbn <> C_SIKEN_KOU_KIM then
    '
	'			'�w�N�ɍ��킹���J�ݎ����������ɂ���
	'			w_sSQL = w_sSQL & " and DECODE(T26_GAKUNEN "
	'			w_sSQL = w_sSQL & " ,1, T15_KAISETU1 "
	'			w_sSQL = w_sSQL & " ,2, T15_KAISETU2 "
	'			w_sSQL = w_sSQL & " ,3, T15_KAISETU3 "
	'			w_sSQL = w_sSQL & " ,4, T15_KAISETU4 "
	'			w_sSQL = w_sSQL & " ,5, T15_KAISETU5 "
	'			w_sSQL = w_sSQL & ") IN (" & w_sJiki & "," & C_KAI_TUNEN & ")"
    '
	'		else
    '
	'			'�w�N�ɍ��킹���J�ݎ����������ɂ���
	'			w_sSQL = w_sSQL & " and DECODE(T26_GAKUNEN "
	'			w_sSQL = w_sSQL & " ,1, T15_KAISETU1 "
	'			w_sSQL = w_sSQL & " ,2, T15_KAISETU2 "
	'			w_sSQL = w_sSQL & " ,3, T15_KAISETU3 "
	'			w_sSQL = w_sSQL & " ,4, T15_KAISETU4 "
	'			w_sSQL = w_sSQL & " ,5, T15_KAISETU5 "
	'			w_sSQL = w_sSQL & ") < " & C_KAI_NASI
	'		end if
    '    end if
	'end if
    '
    '
    '
    '
	'w_sSQL = w_sSQL & " union "
    '
	''���ʊ����擾
	'w_sSQL = w_sSQL & " select distinct "
	'w_sSQL = w_sSQL & "		T27_GAKUNEN as GAKUNEN "
	'w_sSQL = w_sSQL & "		,T27_CLASS as CLASS "
	'w_sSQL = w_sSQL & "		,T27_KAMOKU_CD as KAMOKU_CD "
	'w_sSQL = w_sSQL & "		,M41_MEISYO as KAMOKU_NAME "
	'w_sSQL = w_sSQL & "		,0 as KAMOKU_KBN_IS "
	'w_sSQL = w_sSQL & "		,T27_KAMOKU_BUNRUI as KAMOKU_KBN "
	'w_sSQL = w_sSQL & "		,M05_CLASSMEI as CLASS_NAME "
	'w_sSQL = w_sSQL & "		,M05_GAKKA_CD as GAKKA_CD "
	'w_sSQL = w_sSQL & " from "
	'w_sSQL = w_sSQL & "		T27_TANTO_KYOKAN "
	'w_sSQL = w_sSQL & "		,M41_TOKUKATU "
	'w_sSQL = w_sSQL & "		,M05_CLASS "
	'w_sSQL = w_sSQL & " where "
	'w_sSQL = w_sSQL & "		T27_NENDO =" & cint(m_iNendo)
	'w_sSQL = w_sSQL & "	and	T27_KYOKAN_CD ='" & m_sKyokanCd & "'"
	'w_sSQL = w_sSQL & "	and	T27_KAMOKU_CD = M41_TOKUKATU_CD "
	'w_sSQL = w_sSQL & "	and	T27_KAMOKU_BUNRUI = " & C_JIK_TOKUBETU
	'w_sSQL = w_sSQL & "	and	T27_SEISEKI_INP_FLG = " & C_SEISEKI_INP_FLG_YES
	'w_sSQL = w_sSQL & "	and	M41_NENDO =" & cint(m_iNendo)
	'w_sSQL = w_sSQL & " and T27_NENDO = M05_NENDO "
	'w_sSQL = w_sSQL & " and T27_GAKUNEN = M05_GAKUNEN "
	'w_sSQL = w_sSQL & " and T27_CLASS = M05_CLASSNO "
    '
	'w_sSQL = w_sSQL & " order by GAKUNEN,CLASS,KAMOKU_KBN "


	w_sSQL = "SELECT GAKUNEN"
	w_sSQL = w_sSQL & " ,CLASS"
	w_sSQL = w_sSQL & " ,KAMOKU_CD"
	w_sSQL = w_sSQL & " ,KAMOKUMEI AS KAMOKU_NAME"
	w_sSQL = w_sSQL & " ,KAMOKU_KBN_IS"
	w_sSQL = w_sSQL & " ,KAMOKU_KBN"
	w_sSQL = w_sSQL & " ,DECODE(MAIN_GAKKA , 1 , M05_CLASSMEI , M02_GAKKARYAKSYO) AS CLASS_NAME"
	w_sSQL = w_sSQL & " ,GAKKA_CD"
	w_sSQL = w_sSQL & " FROM VWEB_RISYU"
	w_sSQL = w_sSQL & " WHERE NENDO = " & cint(m_iNendo)
	w_sSQL = w_sSQL & " AND (KAISETU IN (" & w_sJiki & " ," & C_KAI_TUNEN & ")"
	w_sSQL = w_sSQL & "     OR SIKEN_KBN = " & cint(m_iSikenKbn)
	w_sSQL = w_sSQL & "     )"
	w_sSQL = w_sSQL & " AND KYOKAN_CD = '" & m_sKyokanCd & "' "
	w_sSQL = w_sSQL & " ORDER BY GAKUNEN"
	w_sSQL = w_sSQL & "         ,CLASS"
	w_sSQL = w_sSQL & "         ,KAMOKU_KBN"
	w_sSQL = w_sSQL & "         ,KAMOKU_CD"
	w_sSQL = w_sSQL & "         ,MAIN_GAKKA DESC"
	'--2019/06/24 Upd End

'response.write "w_sSQL = " & w_sSQL
'response.end

	If gf_GetRecordset(gRs,w_sSQL) <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		Exit function
	End If

	f_GetSubject = true

End Function

'********************************************************************************
'*  [�@�\]  �����N���X���ǂ����𒲂ׂ�
'*  [����]
'********************************************************************************
Function f_GetKongoClass(p_iGakunen,p_iClass,p_iKongo)
	Dim w_sSQL
	Dim w_Rs

	On Error Resume Next
	Err.Clear

	f_GetKongoClass = false

    '== SQL�쐬 ==
    w_sSQL = ""
    w_sSQL = w_sSQL & "SELECT "
    w_sSQL = w_sSQL & "M05_SYUBETU "
    w_sSQL = w_sSQL & "FROM M05_CLASS "
    w_sSQL = w_sSQL & "Where "
    w_sSQL = w_sSQL & "M05_NENDO = " & m_iNendo & " "
    w_sSQL = w_sSQL & "AND "
    w_sSQL = w_sSQL & "M05_GAKUNEN = " & p_iGakunen & " "
    w_sSQL = w_sSQL & "AND "
    w_sSQL = w_sSQL & "M05_CLASSNO = " & p_iClass & " "

	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit function

	'Response.Write w_sSQL

	'//�߂�l���
	If w_Rs.EOF = False Then
		p_iKongo = Cint(w_Rs("M05_SYUBETU"))
	End If

	f_GetKongoClass = true

	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  [�@�\]  �w�ȗ������擾(�\���p)
'*  [����]  �Ȃ�
'*  [�ߒl]  gf_GetGakkaNm:�w�Ȗ�
'*  [����]
'********************************************************************************
Function f_GetGakkaNm(p_iNendo,p_sCD)
	Dim rs
	Dim w_sName

    On Error Resume Next
    Err.Clear

    f_GetGakkaNm = ""
	w_sName = ""

    Do
        w_sSQL = ""
        w_sSQL = w_sSQL & vbCrLf & " SELECT  "
        w_sSQL = w_sSQL & vbCrLf & "    M02_GAKKARYAKSYO "
        w_sSQL = w_sSQL & vbCrLf & " FROM "
        w_sSQL = w_sSQL & vbCrLf & "    M02_GAKKA "
        w_sSQL = w_sSQL & vbCrLf & " WHERE"
        w_sSQL = w_sSQL & vbCrLf & "        M02_GAKKA_CD = '" & p_sCD & "' "
        w_sSQL = w_sSQL & vbCrLf & "    AND M02_NENDO = " & p_iNendo & " "

        iRet = gf_GetRecordset(rs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
			'm_sErrMsg = ""
            Exit Do
        End If

        If rs.EOF = False Then
            w_sName = rs("M02_GAKKARYAKSYO")
        End If

        Exit Do
    Loop

	'//�߂�l���
    f_GetGakkaNm = w_sName

	'//ں��޾��CLOSE
	Call gf_closeObject(rs)

    Err.Clear

End Function


'********************************************************************************
'*  [�@�\]  ���C�f�[�^����X�V�����擾����B
'*  [����]
'*			p_iNendo - �����N�x
'*			p_iGakunen - �w�N
'*			p_sGakkaCd - �w�ȃR�[�h
'*			p_sKamokuCd - �ȖڃR�[�h
'*  [�ߒl]  �X�V���t
'*  [����]
'********************************************************************************
Function f_GetUpdDate(p_iNendo,p_iGakunen,p_sGakkaCd,p_sKamokuCd,p_KamokuKbn)

	Dim w_sSQL
	Dim w_Rs
	Dim w_FieldName
	Dim w_Table,w_TableName,w_KamokuName

	On Error Resume Next
	Err.Clear

	f_GetUpdDate = ""

	if p_KamokuKbn = C_JIK_JUGYO then
		w_Table = "T16"
		w_TableName = "T16_RISYU_KOJIN"
		w_KamokuName = "T16_KAMOKU_CD"
	else
		w_Table = "T34"
		w_TableName = "T34_RISYU_TOKU"
		w_KamokuName = "T34_TOKUKATU_CD"
	end if

	select case m_iSikenKbn
		case C_SIKEN_ZEN_TYU : w_FieldName = w_Table & "_KOUSINBI_TYUKAN_Z"
		case C_SIKEN_ZEN_KIM : w_FieldName = w_Table & "_KOUSINBI_KIMATU_Z"
		case C_SIKEN_KOU_TYU : w_FieldName = w_Table & "_KOUSINBI_TYUKAN_K"
		case C_SIKEN_KOU_KIM : w_FieldName = w_Table & "_KOUSINBI_KIMATU_K"
	end select

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	Max(" & w_FieldName & ") as UPD_DATE "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & 		w_TableName
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	" & w_Table & "_NENDO        =  " & p_iNendo
	w_sSQL = w_sSQL & " And " & w_Table & "_HAITOGAKUNEN =  " & p_iGakunen
	w_sSQL = w_sSQL & " And " & w_Table & "_GAKKA_CD     = '" & p_sGakkaCd & "'"
	w_sSQL = w_sSQL & " And " & w_KamokuName & "    = '" & p_sKamokuCd & "'"
	w_sSQL = w_sSQL & " And " & w_FieldName & " is not NULL "

	if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function

	if w_Rs.EOF then exit function

	f_GetUpdDate = gf_SetNull2String(w_Rs("UPD_DATE"))

	Call gf_closeObject(w_Rs)

End Function

'********************************************************************************
'*  HTML���o��
'********************************************************************************
Sub showPage()
	Dim w_TukuName
	Dim w_SubjectDisp
	Dim w_SubjectValue
	Dim w_sWhere

	Dim w_iGakunen_s
	Dim w_sGakkaCd_s
	Dim w_sKamokuCd_s

	On Error Resume Next
    Err.Clear

%>
	<html>
	<head>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--
	//************************************************************
	//  [�@�\]  �������ύX���ꂽ�Ƃ��A�ĕ\������
	//************************************************************
	function f_ReLoadMyPage(){
		document.frm.action="sei0150_top.asp";
		document.frm.target="topFrame";
		document.frm.submit();
	}

	//************************************************************
	//  [�@�\]  �\���{�^���N���b�N���̏���
	//************************************************************
	function f_Search(){
		// �I�����ꂽ�R���{�̒l���
		f_SetData();

	    document.frm.action="sei0150_bottom.asp";
	    document.frm.target="main";
	    document.frm.submit();
	}

	//************************************************************
	//  [�@�\]  �\���{�^���N���b�N���ɑI�����ꂽ�f�[�^���
	//************************************************************
	function f_SetData(){
		//�f�[�^�擾
		var vl = document.frm.sltSubject.value.split('#@#');

		//�I�����ꂽ�f�[�^���(�w�N�A�N���X�A�Ȗ�CD���擾)
		document.frm.txtGakuNo.value=vl[0];
		document.frm.txtClassNo.value=vl[1];
		document.frm.txtKamokuCd.value=vl[2];
		document.frm.txtGakkaCd.value=vl[3];
		document.frm.txtUpdDate.value=vl[4];
		document.frm.SYUBETU.value=vl[5];
		document.frm.hidKamokuKbn.value=vl[6];
		document.frm.txtClassMei.value=vl[7];		//2019/06/24 Add Fujibayashi
		document.frm.txtKamokuMei.value=vl[8];		//2019/06/24 Add Fujibayashi
	}

	//************************************************************
	//  [�@�\]  �X�V���̃Z�b�g
	//************************************************************
	function f_SetUpdDate(){
		<% if gDisabled = "" then %>
			var vl = document.frm.sltSubject.value.split('#@#');
			document.frm.txtUpdDate.value=vl[4];
		<% end if %>
	}

	//-->
	</SCRIPT>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">
	</head>

    <body LANGUAGE="javascript" onload="f_SetUpdDate();">

	<center>
	<form name="frm" METHOD="post">

	<% call gs_title(" ���ѓo�^ "," �o�@�^ ") %>
	<br>

	<table border="0">
		<tr><td valign="bottom">

			<table border="0" width="100%">
				<tr><td class="search">

					<table border="0">
						<tr valign="middle">
							<td align="left" nowrap>�����敪</td>
							<td align="left" colspan="3">
							<%
								w_sWhere = " M01_NENDO = " & m_iNendo
								w_sWhere = w_sWhere & " AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN)
								w_sWhere = w_sWhere & " AND M01_SYOBUNRUI_CD < " & cint(C_SIKEN_JITURYOKU)
								Call gf_ComboSet("sltShikenKbn",C_CBO_M01_KUBUN,w_sWhere," onchange = 'f_ReLoadMyPage();' style='width:140px;'",false,m_iSikenKbn)
							%>
							</td>
							<td>&nbsp;</td>

							<td align="left" nowrap>�Ȗ�</td>
							<td align="left">

								<select name="sltSubject" onChange="f_SetUpdDate();" <%=gDisabled%>>
								<%
									if gRs.EOF Then
										'//�Ȗڃf�[�^�Ȃ�
										response.write "<option value=''>�S���Ȗڃf�[�^������܂���"
									else
										'//�Ȗڃf�[�^����
										do until gRs.EOF

											'�ȖڃR���{�\����������
											w_SubjectDisp =""
											w_SubjectDisp = w_SubjectDisp & gRs("GAKUNEN") & "�N�@"

											'�����N���X�̐��Ȗڂ̏ꍇ�w�Ȗ��̂�\������B

											If Not f_GetKongoClass(gRs("GAKUNEN"),gRs("CLASS"),m_iKongou) Then
													m_iKongou = C_CLASS_GAKKA
											End If

											If cint(m_iKongou) = cint(C_CLASS_KONGO) Then
												If cint(gRs("KAMOKU_KBN_IS")) = cint(C_KAMOKU_SENMON) Then

													If m_sGakkoNO <> cstr(C_NCT_OKINAWA) Then
														'���ߍ���͂P�N�����w���̐�傾���N���X���ɍs��
														If m_sGakkoNO = cstr(C_NCT_MAIZURU) Then
															'INS 2007/06/11
															If (gRs("KAMOKU_CD") = 80001 ) OR (gRs("KAMOKU_CD") = 80002 ) Then
																w_SubjectDisp = w_SubjectDisp & gRs("CLASS_NAME") & "�@"
															else
																w_SubjectDisp = w_SubjectDisp & f_GetGakkaNm(m_iNendo,gRs("GAKKA_CD")) & "�@"
															End If
															'DEL 2007/06/11 w_SubjectDisp = w_SubjectDisp & gRs("CLASS_NAME") & "�@"
															'INS END 2007/06/11
														'INS 2016/06/14 kiyomoto -->
														elseIf m_sGakkoNO = cstr(C_NCT_KOCHI) Then
																w_SubjectDisp = w_SubjectDisp & gRs("CLASS_NAME") & "�@"
														'INS 2016/06/14 kiyomoto <--
														else
															w_SubjectDisp = w_SubjectDisp & f_GetGakkaNm(m_iNendo,gRs("GAKKA_CD")) & "�@"
														End If
													else	'���ꂾ������ 2004.09.13 suitoh
														If gRs("KAMOKU_CD") >= 900000 then
															w_SubjectDisp = w_SubjectDisp & gRs("CLASS_NAME") & "�@"
														Else
															w_SubjectDisp = w_SubjectDisp & f_GetGakkaNm(m_iNendo,gRs("GAKKA_CD")) & "  "
														End If
													End If
												Else
													'INS STR 2010/05/20 iwata ���m �����N���X�A��ʁA�w��Ȗڂ́@�w�ȕ\���ɂ���
													'UPP Nishimura  (gRs("KAMOKU_CD") = 140050 )���폜
													If m_sGakkoNO = cstr(C_NCT_KOCHI) Then
														If (gRs("KAMOKU_CD") = 140046 ) OR (gRs("KAMOKU_CD") = 140047 ) OR (gRs("KAMOKU_CD") = 140048 ) OR (gRs("KAMOKU_CD") = 180011 ) OR (gRs("KAMOKU_CD") = 180012 ) OR (gRs("KAMOKU_CD") = 180013 ) OR (gRs("KAMOKU_CD") = 180014 ) Then
															w_SubjectDisp = w_SubjectDisp & f_GetGakkaNm(m_iNendo,gRs("GAKKA_CD"))  & "�@"
														Else
															w_SubjectDisp = w_SubjectDisp & gRs("CLASS_NAME") & "�@"
														End If
													Else
													'INS END 2010/05/20 iwata
														w_SubjectDisp = w_SubjectDisp & gRs("CLASS_NAME") & "�@"
													End If
												End If
											Else

												w_SubjectDisp = w_SubjectDisp & gRs("CLASS_NAME") & "�@"
											End If

											w_SubjectDisp = w_SubjectDisp & gRs("KAMOKU_NAME") & "�@"

											w_TukuName = ""

											if cint(gf_SetNull2Zero(gRs("KAMOKU_KBN"))) = 1 then
												w_TukuName = "TOKU"
											else
												w_TukuName = "TUJO"
											end if

											'�ȖڃR���{VALUE��������
											w_SubjectValue = ""
											w_SubjectValue = w_SubjectValue & gRs("GAKUNEN")   & "#@#"
											w_SubjectValue = w_SubjectValue & gRs("CLASS")     & "#@#"
											w_SubjectValue = w_SubjectValue & gRs("KAMOKU_CD") & "#@#"
											w_SubjectValue = w_SubjectValue & gRs("GAKKA_CD")  & "#@#"
											w_SubjectValue = w_SubjectValue & f_GetUpdDate(m_iNendo,gRs("GAKUNEN"),gRs("GAKKA_CD"),gRs("KAMOKU_CD"),cint(gf_SetNull2Zero(gRs("KAMOKU_KBN")))) & "#@#"
											w_SubjectValue = w_SubjectValue & w_TukuName  & "#@#"
											w_SubjectValue = w_SubjectValue & cint(gf_SetNull2Zero(gRs("KAMOKU_KBN")))
											w_SubjectValue = w_SubjectValue & "#@#" & gRs("CLASS_NAME")			'2019/06/24 Add Fujibayashi
											w_SubjectValue = w_SubjectValue & "#@#" & gRs("KAMOKU_NAME")		'2019/06/24 Add Fujibayashi

								%>
										<option value="<%=w_SubjectValue%>"><%=w_SubjectDisp%>
								<%
											gRs.movenext
										loop
									end if
								%>
								</select>
							</td>
						</tr>

						<tr>
							<td align="left" nowrap>�ŏI�X�V��</td>
							<td align="left" colspan="3" nowrap>
								<input type="text" name="txtUpdDate" value="" onFocus="blur();" readonly style="BACKGROUND-COLOR: #E4E4ED">
							</td>

							<td colspan="7" align="right">
								<input type="button" class="button" value="�@�\�@���@" onclick="javasript:f_Search();" <%=gDisabled%>>
							</td>
						</tr>
					</table>

				</td>
				</tr>
			</table>
			</td>
		</tr>
	</table>

	<input type="hidden" name="txtNendo"     value="<%=m_iNendo%>">
	<input type="hidden" name="txtKyokanCd"  value="<%=m_sKyokanCd%>">
	<input type="hidden" name="txtGakuNo"    value="<%=w_iGakunen_s%>">
	<input type="hidden" name="txtClassNo"   value="">
	<input type="hidden" name="txtKamokuCd"  value="<%=w_sKamokuCd_s%>">
	<input type="hidden" name="txtGakkaCd"   value="<%=w_sGakkaCd_s%>">
	<input type="hidden" name="SYUBETU"      value="">
	<input type="hidden" name="hidKamokuKbn" value="">
	<input type="hidden" name="txtClassMei"   value="">		<%	'2019/06/24 Add Fujibayashi	%>
	<input type="hidden" name="txtKamokuMei"   value="">	<%	'2019/06/24 Add Fujibayashi	%>

	</form>
	</center>
	</body>
	</html>
<%
End Sub

'********************************************************************************
'*	��HTML���o��
'********************************************************************************
Sub showWhitePage(p_Msg)
%>
	<html>
	<head>
	<title>���ѓo�^</title>
	<link rel=stylesheet href=../../common/style.css type=text/css>
	</head>

	<body LANGUAGE="javascript">
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