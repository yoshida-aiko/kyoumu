<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���шꗗ
' ��۸���ID : sei/sei0300/sei0300_top.asp
' �@      �\: ��y�[�W ���шꗗ�̌������s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h		��		SESSION���i�ۗ��j
'           :�N�x			��		SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h		��		SESSION���i�ۗ��j
'           :�N�x			��		SESSION���i�ۗ��j
' ��      ��:
'           �������\��
'				�R���{�{�b�N�X�͋󔒂ŕ\��
'			���\���{�^���N���b�N��
'				���̃t���[���Ɏw�肵�������ɂ��Ȃ��������̓��e��\��������
'-------------------------------------------------------------------------
' ��      ��: 2001/09/04 �ɓ����q
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�كR���X�g /////////////////////////////

	Public CONST C_KENGEN_SEI0300_FULL = "FULL"	'//�A�N�Z�X����FULL
	Public CONST C_KENGEN_SEI0300_TAN = "TAN"	'//�A�N�Z�X�����S�C
	Public CONST C_KENGEN_SEI0300_GAK = "GAK"	'//�A�N�Z�X�����w��

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_bErrMsg           '�װү����

    Public m_iNendo				'�N�x
    Public m_sKyokanCd			'�����R�[�h
	'�����敪�p��Where����
    Public m_sSikenKBN			'�����敪�R���{�{�b�N�X�ɓ���l
    Public m_sSikenKBNWhere		'�����R���{�{�b�N�X�̏���
	'�N���X�p��Where����
    Public m_sGakuNo			'�N���X�̊w�N�R���{�{�b�N�X�ɓ���l
    Public m_sGakuNoWhere		'�N���X�̊w�N�R���{�{�b�N�X�̏���
    Public m_sClassNo			'�N���X�̊w�ȃR���{�{�b�N�X�ɓ���l
    Public m_sClassNoWhere		'�N���X�̊w�ȃR���{�{�b�N�X�̏���
    Public m_sGakusei			'�N���X�̊w�N�R���{�{�b�N�X�ɓ���l
    Public m_sGakuseiWhere		'�N���X�̊w�N�R���{�{�b�N�X�̏���
    Public m_sGakkaNo			'�w�ȃR�[�h

    Public m_sKengen			'����
'    Public m_bTannin
'    Public m_bGakka

    Public m_sOption			'�N���X�̊w�ȃR���{�{�b�N�X�̎g�p�A�s�̔���
    Public m_sGakuNoOption
    Public m_sClassNoOption
    Public m_sGakuseiOption


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
	w_sMsgTitle="�l�ʐ��шꗗ"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"     
	w_sTarget="_top"


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

'	m_iNendo	= session("NENDO")
'	m_sKyokanCd	= session("KYOKAN_CD")

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

		'//�p�����[�^�Z�b�g
		Call s_SetParam

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

		'�����敪�p��Where����
        Call f_SikenKBNWhere()

		'�N���X�̊w�N�p��Where����
        Call f_GakuNoWhere()

	If w_sKengen <> C_KENGEN_SEI0300_GAK Then
		'�N���X�̑g�p��Where����
		Call  f_ClassNoWhere()
	End If
	
		'�敪�p��Where����
		Call f_GakuseiWhere()
		
	   '// �y�[�W��\��
	   Call showPage()
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
'*  [�@�\]  �S���ڂɈ����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

	m_iNendo	= session("NENDO")
	m_sKyokanCd	= session("KYOKAN_CD")

	m_sSikenKBN = trim(Request("txtSikenKBN"))
	m_sGakuNo   = Replace(trim(Request("txtGakuNo")),"@@@","")
	m_sClassNo  = Replace(trim(Request("txtClassNo")),"@@@","")
	m_sGakusei  = Replace(trim(Request("txtGakusei")),"@@@","")

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
	p_bTannin = False

	Do 

		'// �S�C�N���X���
		w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & "  M05_CLASS.M05_GAKUNEN "
		w_sSQL = w_sSQL & vbCrLf & "  ,M05_CLASS.M05_CLASSNO "
		w_sSQL = w_sSQL & vbCrLf & " FROM M05_CLASS"
		w_sSQL = w_sSQL & vbCrLf & " WHERE "
		w_sSQL = w_sSQL & vbCrLf & "      M05_CLASS.M05_NENDO=" & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M05_CLASS.M05_TANNIN='" & m_sKyokanCd & "'"

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
		Else
			p_sKengen = C_KENGEN_SEI0300_TAN 
			m_sGakuNo  = rs("M05_GAKUNEN")
			m_sClassNo = rs("M05_CLASSNO")

			'//�������S�C�̏ꍇ�́A�S�C�N���X�ȊO�͑I���ł��Ȃ�
			m_sGakuNoOption = " DISABLED "
			m_sClassNoOption = " DISABLED "
		End If

		f_GetClassInfo = 0
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
		w_sSQL = w_sSQL & vbCrLf & "      M04_NENDO=" & m_iNendo
		w_sSQL = w_sSQL & vbCrLf & "  AND M04_KYOKAN_CD='" & m_sKyokanCd & "'"
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
			m_sGakkaNo  = rs("M04_GAKKA_CD")
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

Function f_SikenKBNWhere()
'********************************************************************************
'*	[�@�\]	�����敪�R���{�Ɋւ���WHERE���쐬����
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************

	m_sSikenKBNWhere=""
	m_sSikenKBNWhere = " M01_NENDO = " & m_iNendo & " "
	m_sSikenKBNWhere = m_sSikenKBNWhere & " AND M01_DAIBUNRUI_CD = " & cint(C_SIKEN) & " "
	m_sSikenKBNWhere = m_sSikenKBNWhere & " AND M01_SYOBUNRUI_CD < " & cint(C_SIKEN_JITURYOKU) & " "

End Function

Function f_GakuNoWhere()
'********************************************************************************
'*	[�@�\]	�N���X�̊w�N�R���{�Ɋւ���WHERE���쐬����
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************

	m_sGakuNoWhere=""
	m_sGakuNoWhere = " M05_NENDO = " & m_iNendo & " "
	m_sGakuNoWhere = m_sGakuNoWhere & " GROUP BY M05_GAKUNEN "

End Function

Sub f_ClassNoWhere()
'********************************************************************************
'*	[�@�\]	�N���X�̑g�R���{�Ɋւ���WHERE���쐬����
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************

	m_sClassNoWhere=""
	m_sOption=""
	If m_sGakuNo <> "" Then
		m_sClassNoWhere = " M05_NENDO = " & m_iNendo & " AND "
		m_sClassNoWhere = m_sClassNoWhere & " M05_GAKUNEN = " & m_sGakuNo & " "
	Else
		m_sClassNoOption = " DISABLED "
		m_sClassNoWhere  = " M05_GAKUNEN = 99 "
	End IF

End Sub

Sub f_GakuseiWhere()
'********************************************************************************
'*  [�@�\]  �����R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_sGakuseiWhere=""

 	If m_sClassNo <> "" Then
	    m_sGakuseiWhere = " T11_GAKUSEI_NO = T13_GAKUSEI_NO "
'	    m_sGakuseiWhere = m_sGakuseiWhere & " AND T11_NYUNENDO = T13_NENDO - T13_GAKUNEN + 1"
	    m_sGakuseiWhere = m_sGakuseiWhere & " AND T13_GAKUNEN = " & m_sGakuNo
	    m_sGakuseiWhere = m_sGakuseiWhere & " AND T13_CLASS = " & m_sClassNo
	    m_sGakuseiWhere = m_sGakuseiWhere & " AND T13_NENDO = " & m_iNendo
		m_sGakuseiWhere = m_sGakuseiWhere & " AND T13_ZAISEKI_KBN < " & C_ZAI_SOTUGYO

	ElseIf m_sKengen = C_KENGEN_SEI0300_GAK AND m_sGakuNo <> "" Then
	    m_sGakuseiWhere = " T11_GAKUSEI_NO = T13_GAKUSEI_NO "
'	    m_sGakuseiWhere = m_sGakuseiWhere & " AND T11_NYUNENDO = T13_NENDO - T13_GAKUNEN + 1"
	    m_sGakuseiWhere = m_sGakuseiWhere & " AND T13_GAKUNEN = " & m_sGakuNo
	    m_sGakuseiWhere = m_sGakuseiWhere & " AND T13_GAKKA_CD = " & m_sGakkaNo
	    m_sGakuseiWhere = m_sGakuseiWhere & " AND T13_NENDO = " & m_iNendo
		m_sGakuseiWhere = m_sGakuseiWhere & " AND T13_ZAISEKI_KBN < " & C_ZAI_SOTUGYO

	Else
		m_sGakuseiOption = " DISABLED "
		m_sGakuseiWhere  = " T13_GAKUNEN = 99 "
	End IF

End Sub

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
%>
	<html>
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--
	//************************************************************
	//	[�@�\]	�N�x���C�����ꂽ�Ƃ��A�ĕ\������
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_ReLoadMyPage(){

		document.frm.action="sei0300_top.asp";
		document.frm.target="top";
		document.frm.submit();

	}

	//************************************************************
	//	[�@�\]	�N���A�{�^���������ꂽ�ꍇ
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_Clear(){

		document.frm.txtGakuNo.value = "";
		document.frm.txtClassNo.value = "";
		document.frm.txtGakusei.value = "";

	}

	//************************************************************
	//	[�@�\]	�\���{�^���N���b�N���̏���
	//	[����]	�Ȃ�
	//	[�ߒl]	�Ȃ�
	//	[����]
	//
	//************************************************************
	function f_Search(){

		// ������NULL����������
		// ���w�N
		if( f_Trim(document.frm.txtGakuNo.value) == "<%=C_CBO_NULL%>" ){
			window.alert("�w�N�̑I�����s���Ă�������");
			document.frm.txtGakuNo.focus();
			return ;
		}
	<%If m_sKengen <> C_KENGEN_SEI0300_GAK Then%>
		// ���N���X
		if( f_Trim(document.frm.txtClassNo.value) == "<%=C_CBO_NULL%>" ){
			window.alert("�N���X�̑I�����s���Ă�������");
			document.frm.txtClassNo.focus();
			return ;
		}
	<%End If%>
		// ���w��
		if( f_Trim(document.frm.txtGakusei.value) == "<%=C_CBO_NULL%>" ){
		    window.alert("�w���̑I�����s���Ă�������");
		    document.frm.txtGakusei.focus();
		    return ;
		}

		// ���w�N
		if( f_Trim(document.frm.txtGakuNo.value) == "" ){
			window.alert("�w�N�̑I�����s���Ă�������");
			document.frm.txtGakuNo.focus();
			return ;
		}
	<%If m_sKengen <> C_KENGEN_SEI0300_GAK Then%>
		// ���N���X
		if( f_Trim(document.frm.txtClassNo.value) == "" ){
			window.alert("�N���X�̑I�����s���Ă�������");
			document.frm.txtClassNo.focus();
			return ;
		}
	<%End If%>
		// ���w��
		if( f_Trim(document.frm.txtGakusei.value) == "<%=C_CBO_NULL%>" ){
		    window.alert("�w���̑I�����s���Ă�������");
		    document.frm.txtGakusei.focus();
		    return ;
		}

		document.frm.action="sei0300_main.asp";
		document.frm.target="<%=C_MAIN_FRAME%>";
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
	<br>
	<table border="0">
		<tr><td valign="bottom">

			<table border="0">
				<tr><td class="search">

					<table border="0">
						<tr>
							<td align="left" nowrap>�����敪</td>
							<td nowrap><%call gf_ComboSet("txtSikenKBN",C_CBO_M01_KUBUN,m_sSikenKBNWhere, "style='width:150px;' ",False,m_sSikenKBN) %></td>
						<td></td>
						<td></td>
						</tr>
						<tr>
							<td align="left" nowrap>�N���X</td>
							<td nowrap><%
								call gf_ComboSet("txtGakuNo",C_CBO_M05_CLASS_G,m_sGakuNoWhere,"style='width:40px;' onchange='javascript:f_ReLoadMyPage()' " & m_sGakuNoOption,True,m_sGakuNo)%>�N�@<%
							If m_sKengen <> C_KENGEN_SEI0300_GAK then 
								call gf_ComboSet("txtClassNo",C_CBO_M05_CLASS,m_sClassNoWhere,"style='width:80px;'  onchange='javascript:f_ReLoadMyPage()' "& m_sClassNoOption,True,m_sClassNo)
							else %>
								<%=gf_GetGakkaNm(m_iNendo,m_sGakkaNo)%>
						 <% end if %>
							</td>
					            <td Nowrap align="center">�@���@���@
								<%call gf_PluComboSet("txtGakusei",C_CBO_T11_GAKUSEKI_N,m_sGakuseiWhere, "style='width:250px;'"& m_sGakuseiOption,True,m_sGakusei)%>
							</td>
						</tr>
						<tr>
					        <td colspan="4" align="right" valign="bottom"  nowrap>
					        <input type="button" class="button" value=" �N�@���@�A " onclick="javasript:f_Clear();">
					        <input type="button" class="button" value="�@�\�@���@" onclick="javasript:f_Search();">
					        </td>
						</tr>
					</table>
				</td></tr>
			</table>
		</tr>
	</table>

	<%If m_sKengen=C_KENGEN_SEI0300_TAN Then%>
		<input type="hidden" name="txtGakuNo" value="<%=m_sGakuNo%>">
		<input type="hidden" name="txtClassNo" value="<%=m_sClassNo%>">
	<%ElseIf m_sKengen=C_KENGEN_SEI0300_GAK Then%>
		<input type="hidden" name="txtGakkaNo" value="<%=m_sGakkaNo%>">
	<%End If%>

	</form>
	</center>
	</body>
	</html>
<%
End Sub
%>