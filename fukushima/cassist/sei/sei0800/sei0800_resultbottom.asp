<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���юQ�Ɓi�������j
' ��۸���ID : sei/sei0800/default.asp
' �@      �\: 
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h		��		SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h		��		SESSION���i�ۗ��j
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2003/05/13 �A�c
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////

	Dim  m_iNendo   		'�N�x
	Dim  m_bErrFlg			'�װ�׸�
	Dim  m_Rs				'���R�[�h�Z�b�g
	Dim  m_RecCnt			'���R�[�h�J�E���g�i�Ȗڂ̃J�E���g�j
	Dim  m_sGakuseiNo		'�Ώۊw���̊w���ԍ�
	Dim  m_IppanCnt			'��ʁ@�@�@�����s��
	Dim  m_SenmonCnt		'���@�@�@�����s��
	Dim  m_Ippan_H			'��ʕK�C�@�����s��
	Dim  m_Senmon_H			'���K�C�@�����s��
	Dim  m_Ippan_S			'��ʑI���@�����s��
	Dim  m_Senmon_S			'���I���@�����s��
	Dim  m_bKamokuKBN		'�Ȗڋ敪�^�C�g���\���t���O
	Dim  m_bHissenKBN		'�K�I�敪�^�C�g���\���t���O

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

	Dim w_sWinTitle
	Dim w_sMsgTitle
	Dim w_sMsg
	Dim w_sRetURL
	Dim w_sTarget

	'Message�p�̕ϐ��̏�����
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="���юQ��"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_parent"

	On Error Resume Next
	Err.Clear

	m_bErrFlg = False

	Do
		'// �ް��ް��ڑ�
		If gf_OpenDatabase() <> 0 Then
			'�ް��ް��Ƃ̐ڑ��Ɏ��s
			m_bErrFlg = True
			m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
			Exit Do
		End If

		'// �����`�F�b�N�Ɏg�p
'		Session("PRJ_No") = "SEI0800"

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(Session("PRJ_No"))

		'//���Ұ�SET
		Call s_SetParam()

		'// �Y���w�����уf�[�^�擾
		If Not f_GetGakResult() Then m_bErrFlg = True : Exit Do

		'// �Y���҂����Ȃ��ꍇ
		If m_Rs.EOF Then
			Call gs_showWhitePage("�l���C�f�[�^�����݂��܂���B","���юQ��")
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
    Call gs_CloseDatabase()
End Sub

'********************************************************************************
'*	[�@�\]	�S���ڂɈ����n����Ă����l��ݒ�
'********************************************************************************
Sub s_SetParam()

	m_iNendo     = Session("NENDO")				'�����N�x
	m_sGakuseiNo = Request("hidGakuseiNo")		'�w��NO

End Sub

Function f_GetGakResult()
'********************************************************************************
'*  [�@�\]  ���O�C���������S������N���X�̊w���ꗗ���擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  True / False
'*  [����]  
'********************************************************************************
	On Error Resume Next
	Err.Clear

	Dim w_sSQL
	Dim w_sKamokuCD
	Dim w_lRowCnt
	Dim w_lRowCnt2

	f_GetGakResult = False

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	T16_KAMOKU_KBN       AS KAMOKU_KBN,"
	w_sSQL = w_sSQL & " 	T16_HISSEN_KBN       AS HISSEN_KBN,"
	w_sSQL = w_sSQL & " 	T16_COURSE_CD        AS COURSE_CD, "
	w_sSQL = w_sSQL & " 	T16_SEQ_NO           AS SEQ_NO,    "
	w_sSQL = w_sSQL & " 	T16_HAITOGAKUNEN     AS HAITOGAKUNEN,"
	w_sSQL = w_sSQL & " 	T16_KAMOKU_CD        AS KAMOKU_CD,   "
	w_sSQL = w_sSQL & " 	T16_KAMOKUMEI        AS KAMOKUMEI,   "
	w_sSQL = w_sSQL & " 	T16_HAITOTANI        AS HAITOTANI,   "
	w_sSQL = w_sSQL & " 	T16_KYOKAKEIRETU_KBN AS KYOKAKEIRETU_KBN,"
	w_sSQL = w_sSQL & " 	T16_KYOKAKEIRETU_MEI AS KYOKAKEIRETU_MEI,"
	w_sSQL = w_sSQL & " 	T16_SEI_KIMATU_K     AS SEI_KIMATU_K,    "
	w_sSQL = w_sSQL & " 	T16_HTEN_KIMATU_K    AS HTEN_KIMATU_K,   "
	w_sSQL = w_sSQL & " 	T16_HYOKA_KIMATU_K   AS HYOKA_KIMATU_K,  "
	w_sSQL = w_sSQL & " 	T16_HYOTEI_KIMATU_K  AS HYOTEI_KIMATU_K, "
	w_sSQL = w_sSQL & " 	T16_TANI_SUMI        AS TANI_SUMI,       "
	w_sSQL = w_sSQL & " 	T16_HYOKA_FUKA_KBN   AS HYOKA_FUKA_KBN,  "
	w_sSQL = w_sSQL & " 	T16_OKIKAE_FLG       AS OKIKAE_FLG,      "
	w_sSQL = w_sSQL & " 	T16_SELECT_FLG       AS SELECT_FLG       "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T16_RISYU_KOJIN "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 		T16_NENDO       =  " & m_iNendo
	w_sSQL = w_sSQL & " 	AND T16_GAKUSEI_NO  = '" & m_sGakuseiNo & "' "
	w_sSQL = w_sSQL & " 	AND (T16_HISSEN_KBN =  " & C_HISSEN_HIS & " OR (T16_HISSEN_KBN = " & C_HISSEN_SEN & " AND T16_SELECT_FLG= " & C_SENTAKU_YES & "))"
	w_sSQL = w_sSQL & " 	AND T16_OKIKAE_FLG  <> " & C_TIKAN_KAMOKU_MOTO	'�u�����ȊO
	w_sSQL = w_sSQL & " UNION ALL "
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	T17_KAMOKU_KBN        AS KAMOKU_KBN,"
	w_sSQL = w_sSQL & " 	T17_HISSEN_KBN        AS HISSEN_KBN,"
	w_sSQL = w_sSQL & " 	T17_COURSE_CD         AS COURSE_CD, "
	w_sSQL = w_sSQL & " 	T17_SEQ_NO            AS SEQ_NO,    "
	w_sSQL = w_sSQL & " 	T17_HAITOGAKUNEN      AS HAITOGAKUNEN,"
	w_sSQL = w_sSQL & " 	T17_KAMOKU_CD         AS KAMOKU_CD,   "
	w_sSQL = w_sSQL & " 	T17_KAMOKUMEI         AS KAMOKUMEI,   "
	w_sSQL = w_sSQL & " 	T17_HAITOTANI         AS HAITOTANI,   "
	w_sSQL = w_sSQL & " 	T17_KYOKAKEIRETU_KBN  AS KYOKAKEIRETU_KBN,"
	w_sSQL = w_sSQL & " 	T17_KYOKAKEIRETU_MEI  AS KYOKAKEIRETU_MEI,"
	w_sSQL = w_sSQL & " 	T17_SEI_KIMATU_K      AS SEI_KIMATU_K,    "
	w_sSQL = w_sSQL & " 	T17_HTEN_KIMATU_K     AS HTEN_KIMATU_K,   "
	w_sSQL = w_sSQL & " 	T17_HYOKA_KIMATU_K    AS HYOKA_KIMATU_K,  "
	w_sSQL = w_sSQL & " 	T17_HYOTEI_KIMATU_K   AS HYOTEI_KIMATU_K, "
	w_sSQL = w_sSQL & " 	T17_TANI_SUMI         AS TANI_SUMI,       "
	w_sSQL = w_sSQL & " 	T17_HYOKA_FUKA_KBN    AS HYOKA_FUKA_KBN,  "
	w_sSQL = w_sSQL & " 	T17_OKIKAE_FLG        AS OKIKAE_FLG,      "
	w_sSQL = w_sSQL & " 	T17_SELECT_FLG        AS SELECT_FLG       "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T17_RISYUKAKO_KOJIN, "
	w_sSQL = w_sSQL & " 	(SELECT "
	w_sSQL = w_sSQL & " 		T13_NENDO      AS NENDO, "
	w_sSQL = w_sSQL & " 		T13_GAKUSEI_NO AS GAKUSEI_NO "
	w_sSQL = w_sSQL & " 	 FROM "
	w_sSQL = w_sSQL & " 		T13_GAKU_NEN "
	w_sSQL = w_sSQL & " 	 WHERE "
	w_sSQL = w_sSQL & " 	 		 T13_GAKUSEI_NO = '" & m_sGakuseiNo & "'"
	w_sSQL = w_sSQL & "   		AND (T13_RYUNEN_FLG =  " & C_RYUNEN_OFF & " OR T13_RYUNEN_FLG IS NULL ) "
	w_sSQL = w_sSQL & "   	) T13 "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 		 T17_NENDO      = T13.NENDO "
	w_sSQL = w_sSQL & " 	AND  T17_GAKUSEI_NO = T13.GAKUSEI_NO "
	w_sSQL = w_sSQL & " 	AND (T17_HISSEN_KBN =  " & C_HISSEN_HIS & " OR (T17_HISSEN_KBN = " & C_HISSEN_SEN & " AND T17_SELECT_FLG= " & C_SENTAKU_YES & " )) "
	w_sSQL = w_sSQL & " 	AND  T17_OKIKAE_FLG <> " & C_TIKAN_KAMOKU_MOTO
	w_sSQL = w_sSQL & " ORDER BY "
	w_sSQL = w_sSQL & " 	KAMOKU_KBN, "
	w_sSQL = w_sSQL & " 	HISSEN_KBN, "
	w_sSQL = w_sSQL & " 	COURSE_CD,  "
	w_sSQL = w_sSQL & " 	KYOKAKEIRETU_KBN, "
	w_sSQL = w_sSQL & " 	SEQ_NO,      "
	w_sSQL = w_sSQL & " 	OKIKAE_FLG,  "
	w_sSQL = w_sSQL & " 	HAITOGAKUNEN "

	If gf_GetRecordset(m_Rs,w_sSQL) <> 0 Then Exit Function

	'�ϐ�������
	m_IppanCnt  = 0
	m_SenmonCnt = 0
	m_Ippan_H   = 0
	m_Senmon_H  = 0
	m_Ippan_S   = 0
	m_Senmon_S  = 0
	m_RecCnt    = 0
	w_sKamokuCD = ""

	'1�������R�[�h�����݂��Ȃ��ꍇ
	If m_Rs.EOF then
		f_GetGakResult = True
		Exit Function
	End If

	'�Ȗڋ敪�ʂɃJ�E���g���擾
	Do While Not m_Rs.EOF
		If w_sKamokuCD <> m_Rs("KAMOKU_CD") Then					'�Ȗڕ�
			m_RecCnt = m_RecCnt + 1									'�Ȗڐ�
			w_sKamokuCD = m_Rs("KAMOKU_CD")							'�ȖڃR�[�h�ꎞ�i�[
			If Cint(m_Rs("KAMOKU_KBN")) = C_KAMOKU_IPPAN Then
				m_IppanCnt = m_IppanCnt + 1							'��ʉȖڑ�����
				If Cint(m_Rs("HISSEN_KBN")) = C_HISSEN_HIS Then
					m_Ippan_H = m_Ippan_H + 1						'��ʉȖڕK�C
				ElseIf Cint(m_Rs("HISSEN_KBN")) = C_HISSEN_SEN Then
					m_Ippan_S = m_Ippan_S + 1						'��ʉȖڑI��
				End If
			ElseIf Cint(m_Rs("KAMOKU_KBN")) = C_KAMOKU_SENMON Then
				m_SenmonCnt = m_SenmonCnt + 1						'���Ȗڑ�����
				If Cint(m_Rs("HISSEN_KBN")) = C_HISSEN_HIS Then
					m_Senmon_H = m_Senmon_H + 1						'���ȖڕK�C
				ElseIf Cint(m_Rs("HISSEN_KBN")) = C_HISSEN_SEN Then
					m_Senmon_S = m_Senmon_S + 1						'���ȖڑI��
				End If
			End If
		End If
		m_Rs.MoveNext
	Loop

	'//ں��ރJ�E���g�擾
'	m_RecCnt  = gf_GetRsCount(m_Rs)
	w_lRowCnt  = Cint(m_IppanCnt) + Cint(m_SenmonCnt)
	w_lRowCnt2 = Cint(m_Ippan_H) + Cint(m_Ippan_S) + Cint(m_Senmon_H) + Cint(m_Senmon_S)

	'�\�����錏���ƃ^�C�g�������s�����������ǂ����𔻒�i�Ȗڋ敪�j���O�̂���
	If Cint(m_RecCnt) = Cint(w_lRowCnt) Then
		m_bKamokuKBN = True
	Else
		m_bKamokuKBN = False
	End If

	'�\�����錏���ƃ^�C�g�������s�����������ǂ����𔻒�i�K�I�敪�j���O�̂���
	If Cint(m_RecCnt) = Cint(w_lRowCnt2) Then
		m_bHissenKBN = True
	Else
		m_bHissenKBN = False
	End If

	m_Rs.MoveFirst

	f_GetGakResult = True

End Function

Function f_GetHissen(p_sHissen)
'********************************************************************************
'*  [�@�\]  �K�I�敪�����擾
'*  [����]  p_sHissen - �K�I�敪
'*  [�ߒl]  True / False
'*  [����]  
'********************************************************************************
	On Error Resume Next
	Err.Clear

	Dim w_sSQL
	Dim w_Rs

	f_GetHissen = ""

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUIMEI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	M01_KUBUN "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	M01_NENDO        = " & m_iNendo & " AND "
	w_sSQL = w_sSQL & " 	M01_DAIBUNRUI_CD = " & C_HISSEN & " AND "
	w_sSQL = w_sSQL & " 	M01_SYOBUNRUI_CD = " & p_sHissen

	If gf_GetRecordset(w_Rs,w_sSQL) <> 0 Then Exit Function

	f_GetHissen = w_Rs("M01_SYOBUNRUIMEI")

	w_Rs.Close
	Set w_Rs = Nothing

End Function

'****************************************************************************************
'///////////////////            �p�[�Z���g�`���̐����֐�              ///////////////////
'----------------------------------------------------------------------------------------
'[����]:
'		pValue - �ϊ��Ώ�
'		pUNum  - �����_�ȉ��\������
'  [�߂�l]�F
'		�ϊ���̕�����
'[���l]:
'[�쐬]:2003/04/10 shin
'****************************************************************************************
Function f_FormatPercent(pValue,pUNum)
	Dim wRet , wValue , wUNum
	
	on error resume next
	
	f_FormatPercent = ""

	wValue = trim(pValue)
	wUNum = trim(pUNum)
'	If gf_IsNull(wValue) Then wValue = 0
	If gf_IsNull(wValue) Then exit function
	If gf_IsNull(wUNum) Then wUNum = 0
	
	If Err.number <> 0 Then Exit Function
	wRet = FormatNumber(wValue,wUNum,-1,,-1)
	
	f_FormatPercent = wRet
	
End Function

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
	On Error Resume Next
	Err.Clear

	'// �ϐ���`
	Dim w_sCell						'�Z���̃N���X�ݒ�
	Dim w_lTani						'�ȖڕʏC���P�ʐ�
	Dim w_lTotalTani				'���v�C���P�ʐ�
	Dim w_sKamokuCD					'�ȖڃR�[�h
	Dim w_sKamokuNM					'�Ȗږ�
	Dim w_bLoadFLG					'�������[�h�t���O
	Dim w_sKamokuKBN				'�Ȗڋ敪�R�[�h
	Dim w_bDispFLG					'�Ȗڋ敪�^�C�g���\���ρE���\���t���O
	Dim w_bDispFLG2					'�K�I�敪�^�C�g���\���ρE���\���t���O�i��ʁj
	Dim w_bDispFLG3					'�K�I�敪�^�C�g���\���ρE���\���t���O�i���j
	Dim w_lGakTani(5)				'�w�N�ʒP�ʐ����v
	Dim w_sSei(5)					'����
	Dim w_sTdColor(5)				'�C���E���C���ʃZ���J���[
	Dim i
	Dim w_iGakunen

	'// ������
	w_sCell      = "CELL1"			'�Z���̃N���X�ݒ�i�����l�j
	w_lTani      = 0				'�ȖڕʏC���P�ʐ�
	w_lTotalTani = 0				'���v�C���P�ʐ�
	w_bLoadFLG   = True				'�������[�h�t���O
	w_bDispFLG   = False			'�Ȗڋ敪�^�C�g���\���ρE���\���t���O
	w_bDispFLG2  = False			'�K�I�敪�^�C�g���\���ρE���\���t���O�i��ʁj
	w_bDispFLG3  = False			'�K�I�敪�^�C�g���\���ρE���\���t���O�i���j
	w_bDispFLG4  = True  			'�K�I�敪�^�C�g���\���ρE���\���t���O�i���j
	w_sKamokuCD  = ""

	For i = 1 to 5
		w_sSei(i) = ""
		w_sTdColor(i) = ""
	Next

%>
<html>

<head>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--

	//************************************************************
	//  [�@�\]  �\���{�^������
    //************************************************************
	function jf_Submit(p_i){
		with(document.frm){
			var w_Obj = eval("hidGakNo" + p_i);
			hidGakuseiNo.value = w_Obj.value;
			target = "<%=C_MAIN_FRAME%>";
			action = "sei0800_resultdef.asp";
			submit();
		}
	}

	function jf_Back(){
		with(document.frm){
			target = "<%=C_MAIN_FRAME%>";
			action = "default.asp";
			submit();
		}
	}
	//-->
	</SCRIPT>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">
</head>

<body LANGUAGE="javascript">
	<center>
	<form name="frm" METHOD="post">

	<!-- TABLE���X�g�� -->
	<table border="1" class="hyo" width="630">
<%
	Do While Not m_Rs.EOF

		For i = 1 to 6

			If m_Rs.EOF Then
				Exit For
			End If

			'�Ȗڕω�����HTML�\�����[�v�𔲂����s��
			If w_sKamokuCD <> m_Rs("KAMOKU_CD") Then
				w_sKamokuCD   = m_Rs("KAMOKU_CD")											'�ȖڃR�[�h��ێ�
				If w_bLoadFLG = False then													'�������[�h�̂ݔ�����
					Exit For
				End If
				'�w�N�ʂɃf�[�^���i�[
				w_bLoadFLG   = False														'�������[�h�t���O
			End If

			'// �w�N�ʂɃf�[�^���i�[
			w_iGakunen = Cint(m_Rs("HAITOGAKUNEN"))
			w_sSei(w_iGakunen)     = m_Rs("HYOKA_KIMATU_K")									'����
			w_sTdColor(w_iGakunen) = "style='background : #33CCFF;'"						'���C�J���[
			If Cint(m_Rs("TANI_SUMI")) = 0 Then
				w_sTdColor(w_iGakunen) = "style='background : #FF9900;'"					'���C���J���[
			End If
			w_lGakTani(w_iGakunen) = Cint(w_lGakTani(w_iGakunen)) + Cint(m_Rs("TANI_SUMI"))	'�C���P��

			w_lTani      = w_lTani + Cint(gf_SetNull2Zero(m_Rs("TANI_SUMI")))				'�C���P��
			w_lTotalTani = w_lTotalTani + Cint(gf_SetNull2Zero(m_Rs("TANI_SUMI")))
			w_sKamokuKBN = m_Rs("KAMOKU_KBN")												'�Ȗڋ敪�i��� or ���j
			w_sHissenKBN = m_Rs("HISSEN_KBN")												'�K�I�敪��
			w_sKamokuNM  = m_Rs("KAMOKUMEI")												'�Ȗږ�

			m_Rs.MoveNext
		Next

		w_sCell = gf_IIF(w_sCell="CELL1","CELL2","CELL1")									'�Z���̔w�i�F��ݒ�
%>
		<tr>
<%
'response.write "w_sKamokuKBN = " & w_sKamokuKBN & "<br>"
'response.write "w_sHissenKBN = " & w_sHissenKBN & "<br>"
'response.write "w_bDispFLG3 = " & w_bDispFLG3 & "<br>"
		'�Ȗڋ敪�^�C�g���\����
		If m_bKamokuKBN Then
			'������ AND �Ȗڋ敪 = ���
			If Cint(w_sKamokuKBN) = C_KAMOKU_IPPAN Then
				'�Ȗڋ敪�^�C�g���i��ʉȖځj
				If Not w_bDispFLG Then
					w_bDispFLG = True
					Response.write "<td class='CELL2' width='30' align='center' rowspan=" & m_IppanCnt & " style='writing-mode:tb-rl;' nowrap>��ʉȖ�</td>"
				End If
				'�K�I�敪�^�C�g���i��ʕK�C�j
				If Not w_bDispFLG2 AND Cint(w_sHissenKBN) = C_HISSEN_HIS Then
					w_bDispFLG2 = True
					Response.write "<td class='CELL2' width='30' align='center' rowspan=" & m_Ippan_H & " style='writing-mode:tb-rl;' nowrap>" & f_GetHissen(w_sHissenKBN) & "</td>"
				'�K�I�敪�^�C�g���i��ʑI���j
				ElseIf Not w_bDispFLG3 AND Cint(w_sHissenKBN) = C_HISSEN_SEN Then
					w_bDispFLG3 = True
					Response.write "<td class='CELL2' width='30' align='center' rowspan=" & m_Ippan_S & " style='writing-mode:tb-rl;' nowrap>" & f_GetHissen(w_sHissenKBN) & "</td>"
				End If
			'�������ȊO AND �Ȗڋ敪 = ���
			ElseIf Cint(w_sKamokuKBN) = C_KAMOKU_SENMON Then
				'�Ȗڋ敪�^�C�g���i���Ȗځj
				If w_bDispFLG Then
					w_bDispFLG = False
					Response.write "<td class='CELL2' width='30' align='center' rowspan=" & m_SenmonCnt & " style='writing-mode:tb-rl;' nowrap>���Ȗ�</td>"
				End If
				'�K�I�敪�^�C�g���i���K�C�j
				If w_bDispFLG2 AND Cint(w_sHissenKBN) = C_HISSEN_HIS Then
					w_bDispFLG2 = False
					Response.write "<td class='CELL2' width='30' align='center' rowspan=" & m_Senmon_H & " style='writing-mode:tb-rl;' nowrap>" & f_GetHissen(w_sHissenKBN) & "</td>"
				'�K�I�敪�^�C�g���i���I���j
				ElseIf w_bDispFLG4 AND Cint(w_sHissenKBN) = C_HISSEN_SEN Then

'response.write "m_RecCnt = " & Cint(m_RecCnt) & "<br>"
'response.write "m_IppanCnt = " & Cint(m_IppanCnt) & "<br>"
'response.write "m_SenmonCnt = " & Cint(m_SenmonCnt) & "<br>"
'response.write "m_Ippan_H = " & Cint(m_Ippan_H) & "<br>"
'response.write "m_Ippan_S = " & Cint(m_Ippan_S) & "<br>"
'response.write "m_Senmon_H = " & Cint(m_Senmon_H) & "<br>"
'response.end


					w_bDispFLG4 = False
					Response.write "<td class='CELL2' width='30' align='center' rowspan=" & m_Senmon_S & " style='writing-mode:tb-rl;' nowrap>" & f_GetHissen(w_sHissenKBN) & "</td>"
				End If
			End If
		'�Ȗڋ敪�^�C�g�����\����
		Else
			'������
			If Not w_sDispFLG Then
				w_sDispFLG = True
				Response.write "<td class='CELL2' width='30' align='center' rowspan=" & m_RecCnt & ">&nbsp;&nbsp;&nbsp;&nbsp;</td>"
				Response.write "<td class='CELL2' width='30' align='center' rowspan=" & m_RecCnt & ">&nbsp;&nbsp;&nbsp;&nbsp;</td>"
			End If
		End If
%>
			<td width="250" class=<%=w_sCell%>   align="left"   height="20" nowrap>�@�@�@<%=w_sKamokuNM%></td>
			<td width="70"  class="<%=w_sCell%>" align="center" height="20" nowrap><%=f_FormatPercent(w_lTani,1)%></td>
			<td width="50"  class="<%=w_sCell%>" align="center" height="20" <%=w_sTdColor(1)%> nowrap><%=w_sSei(1)%></td>
			<td width="50"  class="<%=w_sCell%>" align="center" height="20" <%=w_sTdColor(2)%> nowrap><%=w_sSei(2)%></td>
			<td width="50"  class="<%=w_sCell%>" align="center" height="20" <%=w_sTdColor(3)%> nowrap><%=w_sSei(3)%></td>
			<td width="50"  class="<%=w_sCell%>" align="center" height="20" <%=w_sTdColor(4)%> nowrap><%=w_sSei(4)%></td>
			<td width="50"  class="<%=w_sCell%>" align="center" height="20" <%=w_sTdColor(5)%> nowrap><%=w_sSei(5)%></td>
		</tr>
<%
		'// ������
		w_lTani = 0
		For i = 1 to 5
			w_sSei(i) = ""
			w_sTdColor(i) = ""
		Next

	Loop
%>
		<tr>
			<th class="header3" align="center" colspan="3" height="20">���@�@�v</th>
			<th class="header3" align="center"             height="20"><%=f_FormatPercent(w_lTotalTani,1)%></th>
			<th class="header3" align="center"             height="20"><%=f_FormatPercent(w_lGakTani(1),1)%></th>
			<th class="header3" align="center"             height="20"><%=f_FormatPercent(w_lGakTani(2),1)%></th>
			<th class="header3" align="center"             height="20"><%=f_FormatPercent(w_lGakTani(3),1)%></th>
			<th class="header3" align="center"             height="20"><%=f_FormatPercent(w_lGakTani(4),1)%></th>
			<th class="header3" align="center"             height="20"><%=f_FormatPercent(w_lGakTani(5),1)%></th>
		</tr>
	</table>

	<p aling="center"><input type="button" class="button" value="�߁@��" onclick="jf_Back();"></p>

	</center>

	<input type="hidden" name="hidGakuseiNo">

	</form>
</body>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
