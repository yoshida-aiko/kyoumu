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
' �C      ��: 2005/11/22 ����
'             2007.01.29 Takaoka
'             2010/02/17 iwata
'             ���N���X�̉Ȗڂ�I�����Ă���ꍇ�A�R���{�{�b�N�X�ɉȖڂ��\������Ȃ��̂�
'             ���C�Ȗڂ̊J�u���Ă���N���X�̊w�ȂŎ擾����
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Dim  m_bErrFlg           '�װ�׸�
	
	Dim m_iNendo             '�N�x�i�Ȗڂ̔N�x�j�j
	Dim g_iNendo             '�N�x�i�����N�x�j
	Dim m_sKyokanCd          '�����R�[�h
	Dim m_iSikenKbn			'�����敪
	Dim m_bNoData			'���͉Ȗڂ��Ȃ��Ƃ�True
	Dim gRs
	
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
		
		'//�����敪  INS 2005/02/10
		if request("sltShikenKbn")  = "" then
			'//�ŏ�
			w_iRet = gf_Get_SikenKbn(m_iSikenKbn,C_SEISEKI_KIKAN,0)
			
			if w_iRet <> 0 then m_bErrFlg = true : exit do
		else
		    '//�����[�h��
		    m_iSikenKbn = cint(Request("sltShikenKbn"))

		end if

		'�����܂� INS 2005/02/10

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
		
		'//���O�C�������̒S���Ȗڂ̎擾
		if not f_GetSubject() then Exit Do
		
'' �C�� 2007.01.29 Takaoka �Ȗڂ��Ȃ��Ă��\������悤�ύX �R�����g�A�E�g
		If gRs.EOF Then
			m_bNoData = True
		Else
			m_bNoData = False
		End If

'		'�Ȗڃf�[�^�Ȃ�
'		if gRs.EOF Then
'			Call showWhitePage("�S���Ȗڃf�[�^������܂���")
'			response.end
'		End If
'' �C�� 2007.01.29 Takaoka �Ȗڂ��Ȃ��Ă��\������悤�ύX �R�����g�A�E�g�@�����܂�

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
	
    m_iNendo    = session("NENDO")
    g_iNendo    = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")
	
	'//�����敪
	'If Request("sltShikenKbn")  = "" Then
	'	m_iSikenKbn = C_SIKEN_ZEN_TYU
	'Else
	'    m_iSikenKbn = cint(Request("sltShikenKbn"))
	'End If
	
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
    Dim w_Rs
    Dim w_sMinNendo 
    
    On Error Resume Next
    Err.Clear
	
    f_GetSubject = false

	'ADD START 2006/10/13 �V�J 
	'T11�ɑ��݂��郌�R�[�h�̓��w�N�x�̍ŏ��l���擾
	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	MIN(T11_NYUNENDO) as NENDO "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & "     T11_GAKUSEKI "
	
	if gf_GetRecordset(w_Rs,w_sSQL) <> 0 then exit function
	
	if w_Rs.EOF then exit function
	
	w_sMinNendo = gf_SetNull2String(w_Rs("NENDO"))
	
	Call gf_closeObject(w_Rs)

'response.write "w_sMinNendo = " & w_sMinNendo
	'ADD END 2006/10/13 �V�J 

	'�I�񂾎����ɂ���āA�J�݊��Ԃ�ς���
	Select Case cint(m_iSikenKbn)
		Case cint(C_SIKEN_ZEN_TYU),cint(C_SIKEN_ZEN_KIM) : w_sJiki = C_KAI_ZENKI	'�O�����ԁA�O������
		Case cint(C_SIKEN_KOU_TYU),cint(C_SIKEN_KOU_KIM) : w_sJiki = C_KAI_KOUKI	'������ԁA�������
	End Select

	'�ʏ�A���w����։Ȗڎ擾
	w_sSQL = ""
	w_sSQL = w_sSQL & " select distinct "
	w_sSQL = w_sSQL & "		T27_GAKUNEN as GAKUNEN "
	w_sSQL = w_sSQL & "		,T27_CLASS as CLASS "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_CD as KAMOKU_CD "
	w_sSQL = w_sSQL & "		,M03_KAMOKUMEI as KAMOKU_NAME "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_BUNRUI as KAMOKU_KBN "
	w_sSQL = w_sSQL & "		,M05_CLASSMEI as CLASS_NAME "
	w_sSQL = w_sSQL & "		,M05_GAKKA_CD as GAKKA_CD "
	'ADD START 2006/10/13 �V�J 
	w_sSQL = w_sSQL & "		,T27_NENDO as NENDO "
	'ADD END 2006/10/13 �V�J 
	w_sSQL = w_sSQL & " from"
	w_sSQL = w_sSQL & "		T27_TANTO_KYOKAN "
	w_sSQL = w_sSQL & "		,M03_KAMOKU "
	w_sSQL = w_sSQL & "		,T15_RISYU "
	w_sSQL = w_sSQL & "		,M05_CLASS "
	w_sSQL = w_sSQL & " where "
	'UPD START 2006/10/13 �V�J 
	'w_sSQL = w_sSQL & "		T27_NENDO =" & cint(m_iNendo)
	w_sSQL = w_sSQL & "		T27_NENDO >=" & w_sMinNendo
	'UPD END 2006/10/13 �V�J 
	w_sSQL = w_sSQL & "	and	T27_KYOKAN_CD ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_CD = M03_KAMOKU_CD "
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_BUNRUI   = " & C_JIK_JUGYO
	w_sSQL = w_sSQL & "	and	T27_SEISEKI_INP_FLG = " & C_SEISEKI_INP_FLG_YES
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_CD = T15_KAMOKU_CD(+) "
	w_sSQL = w_sSQL & "	and	T15_NYUNENDO = T27_NENDO - T27_GAKUNEN + 1 "
	w_sSQL = w_sSQL & " and T27_NENDO = M05_NENDO "
	w_sSQL = w_sSQL & " and T27_GAKUNEN = M05_GAKUNEN "
	w_sSQL = w_sSQL & " and T27_CLASS = M05_CLASSNO "
	'UPD START 2006/10/13 �V�J 
	'w_sSQL = w_sSQL & "	and	M03_NENDO =" & cint(m_iNendo)
	w_sSQL = w_sSQL & "	and	M03_NENDO >=" & w_sMinNendo
	'UPD END 2006/10/13 �V�J 
	
	'//���������������Ȃ��Ƃ�
'	if m_iSikenKbn <> C_SIKEN_KOU_KIM then
'		w_sSQL = w_sSQL & " and	(("
'		w_sSQL = w_sSQL & "			T15_KAISETU1 =" & w_sJiki & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU2 =" & w_sJiki & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU3 =" & w_sJiki & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU4 =" & w_sJiki & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU5 =" & w_sJiki
'		w_sSQL = w_sSQL & "		 ) "
'		w_sSQL = w_sSQL & "		 or "
'		w_sSQL = w_sSQL & "		 ("
'		w_sSQL = w_sSQL & "			T15_KAISETU1 =" & C_KAI_TUNEN & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU2 =" & C_KAI_TUNEN & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU3 =" & C_KAI_TUNEN & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU4 =" & C_KAI_TUNEN & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU5 =" & C_KAI_TUNEN
'		w_sSQL = w_sSQL & "		 ) )"
'	else
'		w_sSQL = w_sSQL & " and	("
'		w_sSQL = w_sSQL & "			T15_KAISETU1 <" & C_KAI_NASI & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU2 <" & C_KAI_NASI & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU3 <" & C_KAI_NASI & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU4 <" & C_KAI_NASI & " or "
'		w_sSQL = w_sSQL & "			T15_KAISETU5 <" & C_KAI_NASI
'		w_sSQL = w_sSQL & "		 )"
'	end if

	w_sSQL = w_sSQL & " and	(("
	w_sSQL = w_sSQL & "			T15_KAISETU1 =" & w_sJiki & " or "
	w_sSQL = w_sSQL & "			T15_KAISETU2 =" & w_sJiki & " or "
	w_sSQL = w_sSQL & "			T15_KAISETU3 =" & w_sJiki & " or "
	w_sSQL = w_sSQL & "			T15_KAISETU4 =" & w_sJiki & " or "
	w_sSQL = w_sSQL & "			T15_KAISETU5 =" & w_sJiki
	w_sSQL = w_sSQL & "		 ) "

	'�O���w��̏ꍇ�́A�ʔN�͏����ɓ���Ȃ�
'2006.11.20 �O���w��̏ꍇ�́A�ʔN������悤�ɂ���
'2007.01.29 ����ς�O���w��̂Ƃ��͒ʔN�����Ȃ�'
	If w_sJiki <> C_KAI_ZENKI Then
		w_sSQL = w_sSQL & "		 or "
		w_sSQL = w_sSQL & "		 ("
		w_sSQL = w_sSQL & "			T15_KAISETU1 =" & C_KAI_TUNEN & " or "
		w_sSQL = w_sSQL & "			T15_KAISETU2 =" & C_KAI_TUNEN & " or "
		w_sSQL = w_sSQL & "			T15_KAISETU3 =" & C_KAI_TUNEN & " or "
		w_sSQL = w_sSQL & "			T15_KAISETU4 =" & C_KAI_TUNEN & " or "
		w_sSQL = w_sSQL & "			T15_KAISETU5 =" & C_KAI_TUNEN
		'w_sSQL = w_sSQL & "		 ) )"
		w_sSQL = w_sSQL & "		 )"
	End If
'2007.01.29 ����ς�O���w��̂Ƃ��͒ʔN�����Ȃ�'
	w_sSQL = w_sSQL & "		)"
	
	w_sSQL = w_sSQL & vbCrLf & "  UNION ALL "
			
	w_sSQL = w_sSQL & " SELECT DISTINCT "
	w_sSQL = w_sSQL & " 	T27_GAKUNEN AS GAKUNEN"
	w_sSQL = w_sSQL & " 	,T27_CLASS AS CLASS"
	w_sSQL = w_sSQL & " 	,T27_KAMOKU_CD AS KAMOKU_CD "
	w_sSQL = w_sSQL & " 	,T16_KAMOKUMEI AS KAMOKU_NAME "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_BUNRUI as KAMOKU_KBN "
	w_sSQL = w_sSQL & " 	,M05_CLASSMEI AS CLASS_NAME "
	w_sSQL = w_sSQL & " 	,M05_GAKKA_CD AS GAKKA_CD "
	'ADD START 2006/10/13 �V�J 
	w_sSQL = w_sSQL & "		,T27_NENDO as NENDO "
	'ADD END 2006/10/13 �V�J 
	w_sSQL = w_sSQL & " FROM"
	w_sSQL = w_sSQL & " 	T27_TANTO_KYOKAN "
	w_sSQL = w_sSQL & " 	,T16_RISYU_KOJIN "
	w_sSQL = w_sSQL & " 	,M05_CLASS "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 		T27_NENDO = M05_NENDO "
	w_sSQL = w_sSQL & "    AND T27_GAKUNEN = M05_GAKUNEN "
	w_sSQL = w_sSQL & "    AND T27_CLASS = M05_CLASSNO	"
	w_sSQL = w_sSQL & "    AND T27_KAMOKU_CD = T16_KAMOKU_CD(+)"

'2010/02/17 upd iwata
'���N���X�̉Ȗڂ�I�����Ă���ꍇ�A�R���{�{�b�N�X�ɉȖڂ��\������Ȃ��̂�
'���C�Ȗڂ̊J�u���Ă���N���X�̊w�ȂŎ擾����
'	w_sSQL = w_sSQL & "    AND M05_GAKKA_CD(+) = T16_GAKKA_CD "
	w_sSQL = w_sSQL & "    AND M05_GAKKA_CD(+) = T16_CURRI_GAKKA_CD "

	w_sSQL = w_sSQL & "    AND T16_NENDO(+) = T27_NENDO "
	'UPD START 2006/10/13 �V�J 
	'w_sSQL = w_sSQL & "    AND T27_NENDO = " & m_iNendo
	w_sSQL = w_sSQL & "    AND T27_NENDO >= " & w_sMinNendo
	'UPD END 2006/10/13 �V�J 
	w_sSQL = w_sSQL & "    AND T27_KYOKAN_CD ='" & m_sKyokanCd & "' "
	w_sSQL = w_sSQL & "    AND T27_SEISEKI_INP_FLG =" & C_SEISEKI_INP_FLG_YES & " "
	w_sSQL = w_sSQL & "    AND T16_OKIKAE_FLG >= " & C_TIKAN_KAMOKU_SAKI 

'2010.02.12�l���C�ɒǉ������Ȗڂ̊J�ݎ���������悤�ɏC�� �ɓ�
	w_sSQL = w_sSQL & " and	(("
	w_sSQL = w_sSQL & "			T16_KAISETU =" & w_sJiki
	w_sSQL = w_sSQL & "		 ) "
'�O���w��̂Ƃ��͒ʔN�����Ȃ�'
	If w_sJiki <> C_KAI_ZENKI Then
		w_sSQL = w_sSQL & "		 or "
		w_sSQL = w_sSQL & "		 ("
		w_sSQL = w_sSQL & "			T16_KAISETU =" & C_KAI_TUNEN
		w_sSQL = w_sSQL & "		 )"
	End If
'�O���w��̂Ƃ��͒ʔN�����Ȃ�'
	w_sSQL = w_sSQL & "		)"
'2010.02.12�l���C�ɒǉ������Ȗڂ̊J�ݎ���������悤�ɏC�� �ɓ�

	w_sSQL = w_sSQL & " union "
	
	'T27��T27_SEISEKI_INP_FLG=1�̐l�����̐搶�ɐ��ѓo�^���������ہA
	'T26�Ƀf�[�^�����邽�߁A�ȉ���SQL�����s����K�v������
	w_sSQL = w_sSQL & " SELECT distinct "
	w_sSQL = w_sSQL & "		T26_GAKUNEN AS GAKUNEN "
	w_sSQL = w_sSQL & "		,T26_CLASS AS CLASS "
	w_sSQL = w_sSQL & "		,T26_KAMOKU AS KAMOKU_CD "
	w_sSQL = w_sSQL & "		,M03_KAMOKUMEI as KAMOKU_NAME "
	w_sSQL = w_sSQL & "		,0 as KAMOKU_KBN "
	w_sSQL = w_sSQL & "		,M05_CLASSMEI as CLASS_NAME "
	w_sSQL = w_sSQL & "		,M05_GAKKA_CD as GAKKA_CD "
	'ADD START 2006/10/13 �V�J 
	w_sSQL = w_sSQL & "		,T26_NENDO as NENDO "
	'ADD END 2006/10/13 �V�J 
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & "		T26_SIKEN_JIKANWARI "
	w_sSQL = w_sSQL & "		,M03_KAMOKU "
	w_sSQL = w_sSQL & "		,T15_RISYU "
	w_sSQL = w_sSQL & "		,M05_CLASS "
	w_sSQL = w_sSQL & " WHERE "
	'UPD START 2006/10/13 �V�J 
	'w_sSQL = w_sSQL & "		 T26_NENDO = " & cint(m_iNendo)
	w_sSQL = w_sSQL & "		 T26_NENDO >= " & w_sMinNendo
	'UPD END 2006/10/13 �V�J 
	w_sSQL = w_sSQL & "	and ("
	w_sSQL = w_sSQL & "		T26_JISSI_KYOKAN    ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN1 ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN2 ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN3 ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN4 ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & "		OR T26_SEISEKI_KYOKAN5 ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & "		)"
	w_sSQL = w_sSQL & "	and	T26_KAMOKU = M03_KAMOKU_CD "	
	w_sSQL = w_sSQL & "	and T26_KAMOKU = T15_KAMOKU_CD(+) "
	w_sSQL = w_sSQL & "	and T15_NYUNENDO(+) = T26_NENDO - T26_GAKUNEN + 1 "
	w_sSQL = w_sSQL & "	and T26_SIKEN_CD ='" & C_SIKEN_CODE_NULL & "' "
	w_sSQL = w_sSQL & "	and	T26_SEISEKI_INP_FLG = " & C_SEISEKI_INP_FLG_YES
	'UPD START 2006/10/13 �V�J 
	'w_sSQL = w_sSQL & "	and	M03_NENDO =" & cint(m_iNendo)
	w_sSQL = w_sSQL & "	and	M03_NENDO >=" & w_sMinNendo
	'UPD END 2006/10/13 �V�J 
	w_sSQL = w_sSQL & " and T26_NENDO = M05_NENDO "
	w_sSQL = w_sSQL & " and T26_GAKUNEN = M05_GAKUNEN "
	w_sSQL = w_sSQL & " and T26_CLASS = M05_CLASSNO "
	
	'//���������������Ȃ��Ƃ�
	if cint(m_iSikenKbn) <> cint(C_SIKEN_KOU_KIM) then
		w_sSQL = w_sSQL & "	and T26_SIKEN_KBN =" & m_iSikenKbn
	end if
	
	w_sSQL = w_sSQL & " union "
	
	'���ʊ����擾
	w_sSQL = w_sSQL & " select distinct "
	w_sSQL = w_sSQL & "		T27_GAKUNEN as GAKUNEN "
	w_sSQL = w_sSQL & "		,T27_CLASS as CLASS "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_CD as KAMOKU_CD "
	w_sSQL = w_sSQL & "		,M41_MEISYO as KAMOKU_NAME "
	w_sSQL = w_sSQL & "		,T27_KAMOKU_BUNRUI as KAMOKU_KBN "
	w_sSQL = w_sSQL & "		,M05_CLASSMEI as CLASS_NAME "
	w_sSQL = w_sSQL & "		,M05_GAKKA_CD as GAKKA_CD "
	'ADD START 2006/10/13 �V�J 
	w_sSQL = w_sSQL & "		,T27_NENDO as NENDO "
	'ADD END 2006/10/13 �V�J 
	w_sSQL = w_sSQL & " from "
	w_sSQL = w_sSQL & "		T27_TANTO_KYOKAN "
	w_sSQL = w_sSQL & "		,M41_TOKUKATU "
	w_sSQL = w_sSQL & "		,M05_CLASS "
	w_sSQL = w_sSQL & " where "
	'UPD START 2006/10/13 �V�J 
	'w_sSQL = w_sSQL & "		T27_NENDO =" & cint(m_iNendo)
	w_sSQL = w_sSQL & "		T27_NENDO >=" & w_sMinNendo
	'UPD END 2006/10/13 �V�J 
	w_sSQL = w_sSQL & "	and	T27_KYOKAN_CD ='" & m_sKyokanCd & "'"
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_CD = M41_TOKUKATU_CD "
	w_sSQL = w_sSQL & "	and	T27_KAMOKU_BUNRUI = " & C_JIK_TOKUBETU
	w_sSQL = w_sSQL & "	and	T27_SEISEKI_INP_FLG = " & C_SEISEKI_INP_FLG_YES
	'UPD START 2006/10/13 �V�J 
	'w_sSQL = w_sSQL & "	and	M41_NENDO =" & cint(m_iNendo)
	w_sSQL = w_sSQL & "	and	M41_NENDO >=" & w_sMinNendo
	'UPD END 2006/10/13 �V�J 
	w_sSQL = w_sSQL & " and T27_NENDO = M05_NENDO "
	w_sSQL = w_sSQL & " and T27_GAKUNEN = M05_GAKUNEN "
	w_sSQL = w_sSQL & " and T27_CLASS = M05_CLASSNO "
	'UPD START 2006/10/13 �V�J 
	'w_sSQL = w_sSQL & " order by GAKUNEN,CLASS,KAMOKU_KBN "
	w_sSQL = w_sSQL & " order by NENDO DESC,GAKUNEN,CLASS,KAMOKU_KBN "
	'UPD END 2006/10/13 �V�J 
	
	'response.write w_sSQL & "<BR>"
	
'response.write w_sSQL
'response.end		
	If gf_GetRecordset(gRs,w_sSQL) <> 0 Then
		'ں��޾�Ă̎擾���s
		msMsg = Err.description
		Exit function
	End If
	
	f_GetSubject = true
    
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
	
	'UPP 2005/11/22���� CINT()�֐�������
	select case CINT(m_iSikenKbn)
		case CINT(C_SIKEN_ZEN_TYU) : w_FieldName = w_Table & "_KOUSINBI_TYUKAN_Z"
		case CINT(C_SIKEN_ZEN_KIM) : w_FieldName = w_Table & "_KOUSINBI_KIMATU_Z"
		case CINT(C_SIKEN_KOU_TYU) : w_FieldName = w_Table & "_KOUSINBI_TYUKAN_K"
		case CINT(C_SIKEN_KOU_KIM) : w_FieldName = w_Table & "_KOUSINBI_KIMATU_K"
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
<%		'UPD START 2006/10/13 �V�J �p�����[�^�ɔN��ǉ�������
'		document.frm.txtGakuNo.value=vl[0];
'		document.frm.txtClassNo.value=vl[1];
'		document.frm.txtKamokuCd.value=vl[2];
'		document.frm.txtGakkaCd.value=vl[3];
'		document.frm.txtUpdDate.value=vl[4];
'		document.frm.SYUBETU.value=vl[5];
'		document.frm.hidKamokuKbn.value=vl[6];
%>
		document.frm.txtNendo.value=vl[0];
		document.frm.txtGakuNo.value=vl[1];
		document.frm.txtClassNo.value=vl[2];
		document.frm.txtKamokuCd.value=vl[3];
		document.frm.txtGakkaCd.value=vl[4];
		document.frm.txtUpdDate.value=vl[5];
		document.frm.SYUBETU.value=vl[6];
		document.frm.hidKamokuKbn.value=vl[7];
<%		'UPD END 2006/10/13 �V�J �p�����[�^�ɔN��ǉ������� %>
	}
	
	//************************************************************
	//  [�@�\]  �X�V���̃Z�b�g
	//************************************************************
	function f_SetUpdDate(){
		var vl = document.frm.sltSubject.value.split('#@#');
<%		'UPD 2006/10/13 �V�J �p�����[�^�ɔN��ǉ�������
'		document.frm.txtUpdDate.value=vl[4]; %>
		document.frm.txtUpdDate.value=vl[5]; 
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
								<% if not gRs.EOF then %>
									<select name="sltSubject" onChange="f_SetUpdDate();">
									<% 

										do until gRs.EOF
											
											'�ȖڃR���{�\����������
											w_SubjectDisp =""
											'ADD START 2006/10/13 �V�J
											w_SubjectDisp = w_SubjectDisp & gRs("NENDO") & "�@"
											'ADD END 2006/10/13 �V�J
											w_SubjectDisp = w_SubjectDisp & gRs("GAKUNEN") & "�N�@"
											w_SubjectDisp = w_SubjectDisp & gRs("CLASS_NAME") & "�@"
											w_SubjectDisp = w_SubjectDisp & gRs("KAMOKU_NAME") & "�@"
											
											w_TukuName = ""
											
											if cint(gf_SetNull2Zero(gRs("KAMOKU_KBN"))) = 1 then
												w_TukuName = "TOKU"
											else
												w_TukuName = "TUJO"
											end if
											
											'�ȖڃR���{VALUE��������
											w_SubjectValue = ""
											'ADD START 2006/10/13 �V�J
											w_SubjectValue = w_SubjectValue & gRs("NENDO")   & "#@#"
											'ADD END 2006/10/13 �V�J
											w_SubjectValue = w_SubjectValue & gRs("GAKUNEN")   & "#@#"
											w_SubjectValue = w_SubjectValue & gRs("CLASS")     & "#@#"
											w_SubjectValue = w_SubjectValue & gRs("KAMOKU_CD") & "#@#"
											w_SubjectValue = w_SubjectValue & gRs("GAKKA_CD")  & "#@#"
											'UPD START 2006/10/13 �V�J
											'w_SubjectValue = w_SubjectValue & f_GetUpdDate(m_iNendo,gRs("GAKUNEN"),gRs("GAKKA_CD"),gRs("KAMOKU_CD"),cint(gf_SetNull2Zero(gRs("KAMOKU_KBN")))) & "#@#"
											w_SubjectValue = w_SubjectValue & f_GetUpdDate(gRs("NENDO"),gRs("GAKUNEN"),gRs("GAKKA_CD"),gRs("KAMOKU_CD"),cint(gf_SetNull2Zero(gRs("KAMOKU_KBN")))) & "#@#"
											'UPD END 2006/10/13 �V�J
											w_SubjectValue = w_SubjectValue & w_TukuName  & "#@#"
											w_SubjectValue = w_SubjectValue & cint(gf_SetNull2Zero(gRs("KAMOKU_KBN")))
											
									%>
										<option value="<%=w_SubjectValue%>"><%=w_SubjectDisp%>
									<% 
											gRs.movenext
										loop 
									%>
									</select>
								<% 	'' UPD 2007/01/29 EOF �̂Ƃ��̃��b�Z�[�W��ǉ�
									else %>
									�F���̎����敪�̓��͉Ȗڂ͂���܂���
								<% end if %>
							</td>
	                    </tr>
						
						<tr>
							<td align="left" nowrap>�ŏI�X�V��</td>
							<td align="left" colspan="3" nowrap>
								<input type="text" name="txtUpdDate" value="" onFocus="blur();" readonly style="BACKGROUND-COLOR: #E4E4ED">
							</td>
							
							<td colspan="7" align="right">
							<% ' UPD 2007/01/29 EOF �̂Ƃ��̏�����ǉ�
								If Not m_bNoData Then %>
								<input type="button" class="button" value="�@�\�@���@" onclick="javasript:f_Search();">
							<% 	Else %>
							<% 	end if %>
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
	<input type="hidden" name="txtSyoriNendo"     value="<%=g_iNendo%>">
	<input type="hidden" name="txtKyokanCd"  value="<%=m_sKyokanCd%>">
	<input type="hidden" name="txtGakuNo"    value="<%=w_iGakunen_s%>">
	<input type="hidden" name="txtClassNo"   value="">
	<input type="hidden" name="txtKamokuCd"  value="<%=w_sKamokuCd_s%>">
	<input type="hidden" name="txtGakkaCd"   value="<%=w_sGakkaCd_s%>">
	<input type="hidden" name="SYUBETU"      value="">
	<input type="hidden" name="hidKamokuKbn" value="">	
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