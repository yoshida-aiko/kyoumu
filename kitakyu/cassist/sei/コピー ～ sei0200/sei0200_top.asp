<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���шꗗ
' ��۸���ID : sei/sei0100/sei0100_top.asp
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
' ��      ��: 2001/08/08 �O�c �q�j
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
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

    Public m_sGakkaNo			'�w��

	'�Ȗڗp��Where����
    Public m_sKBN				'�敪���R���{�{�b�N�X�ɓ���l
    Public m_sKBNWhere			'�敪���R���{�{�b�N�X�̏���

    Public m_sOption			'�N���X�̊w�ȃR���{�{�b�N�X�̎g�p�A�s�̔���
    Public m_sKengen			'�\������

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
	w_sTarget="_top"


    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

	m_iNendo	= session("NENDO")
	m_sKyokanCd	= session("KYOKAN_CD")

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

		'//�������擾
		w_iRet = f_GetKengen_Sei0200(m_sKengen)
		If w_iRet <> 0 Then
            m_bErrFlg = True
            m_sErrMsg = "�������擾�ł��܂���ł���"
			Exit Do
		End If

		if m_sKengen = C_SEI0200_ACCESS_TANNIN then
	        '//�w�N�̑Ώۂ̃f�[�^�擾
	        w_iRet = f_getData()
	        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do

		ElseIf m_sKengen = C_SEI0200_ACCESS_GAKKA then
			w_iRet = f_GetGakkaInfo(m_sKengen)
	        If w_iRet <> 0 Then m_bErrFlg = True : Exit Do
		End if

		'�����敪�p��Where����
        Call f_SikenKBNWhere()

		'�N���X�̊w�N�p��Where����
        Call f_GakuNoWhere()

	If m_sKengen <> C_SEI0200_ACCESS_GAKKA then
		'�N���X�̑g�p��Where����
		Call  f_ClassNoWhere()
	End If

		'�敪�p��Where����
		Call f_KBNWhere()

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
'*  [�@�\]  �l���C�I���Ȗڌ��菈���̌������擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  p_sKengen
'*  [����]  
'********************************************************************************
Function f_GetKengen_Sei0200(p_sKengen)
	Dim wLevRs

    On Error Resume Next
    Err.Clear

    gf_GetKengen_web0340 = 1

    Do
        w_sSQL = ""
		w_sSQL = w_sSQL & vbCrLf & " SELECT "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID  "
		w_sSQL = w_sSQL & vbCrLf & " FROM "
		w_sSQL = w_sSQL & vbCrLf & " 	T51_SYORI_LEVEL T51 "
		w_sSQL = w_sSQL & vbCrLf & " WHERE  "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_ID In ('SEI0200','SEI0201','SEI0202') AND "
		w_sSQL = w_sSQL & vbCrLf & " 	T51.T51_LEVEL" & session("LEVEL") & " = 1 "
		w_sSQL = w_sSQL & vbCrLf & "ORDER BY T51.T51_ID "

        iRet = gf_GetRecordset(wLevRs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_GetKengen_Sei0200 = 99
            Exit Do
        End If
		if wLevRs.Eof then
            msMsg = "�������擾�ł��܂���ł���"
            f_GetKengen_Sei0200 = 99
            Exit Do
		End if

		Select Case wLevRs("T51_ID")
			Case "SEI0200" : p_sKengen = C_SEI0200_ACCESS_FULL	  		'//�A�N�Z�X����FULL�A�N�Z�X��
			Case "SEI0201" : p_sKengen = C_SEI0200_ACCESS_TANNIN        '//�A�N�Z�X�S�C�A�N�Z�X
			Case "SEI0202" : p_sKengen = C_SEI0200_ACCESS_GAKKA        '//�A�N�Z�X�S�C�A�N�Z�X
		End Select

		'== ���� ==
	    Call gf_closeObject(wLevRs)

        f_GetKengen_Sei0200 = 0
        Exit Do
    Loop

End Function

Function f_getData()
'********************************************************************************
'*  [�@�\]  �w�N�̑Ώۂ̃f�[�^�擾
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    On Error Resume Next
    Err.Clear
    f_getData = 1

    Do

        w_sSQL = ""
        w_sSQL = w_sSQL & " SELECT "
        w_sSQL = w_sSQL & "     M05_GAKUNEN,M05_CLASSNO,M05_CLASSMEI "
        w_sSQL = w_sSQL & " FROM "
        w_sSQL = w_sSQL & "     M05_CLASS "
        w_sSQL = w_sSQL & " WHERE"
        w_sSQL = w_sSQL & "     M05_NENDO  = " & session("NENDO")
        w_sSQL = w_sSQL & " AND M05_TANNIN = '" & session("KYOKAN_CD") & "' "

        iRet = gf_GetRecordset(wClsRs, w_sSQL)
        If w_iRet <> 0 Then
            'ں��޾�Ă̎擾���s
            msMsg = Err.description
            f_getData = 99
            Exit Do
        End If

		m_sGakuNo  = wClsRs("M05_GAKUNEN")
		m_sClassNo = wClsRs("M05_CLASSNO")

        f_getData = 0
        Exit Do
    Loop

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
		w_sSQL = w_sSQL & vbCrLf & "      M04_NENDO=" & session("NENDO")
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
			p_sKengen = C_SEI0200_ACCESS_GAKKA 
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

	m_sSikenKBN = request("txtSikenKBN")

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

	if gf_IsNull(m_sGakuNo) then
		m_sGakuNo = request("txtGakuNo")
		If request("txtGakuNo") = C_CBO_NULL Then m_sGakuNo = ""
	End if

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

		if gf_IsNull(m_sClassNo) then
			m_sClassNo = request("txtClassNo")
		End if

	Else
		m_sOption = " DISABLED "
		m_sClassNoWhere  = " M05_GAKUNEN = 99 "
	End IF

End Sub

Sub f_KBNWhere()
'********************************************************************************
'*	[�@�\]	�Ȗږ��R���{�Ɋւ���WHERE���쐬����
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************

	m_sKBNWhere=""

		m_sKBNWhere = " M01_NENDO = " & m_iNendo & " AND "
		m_sKBNWhere = m_sKBNWhere & " M01_DAIBUNRUI_CD = " & Cint(C_HISSEN) & " "

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

	document.frm.action="sei0200_top.asp";
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

	<% if Cint(m_sKengen) = 0 then %>
        document.frm.txtGakuNo.value = "";
        document.frm.txtClassNo.value = "";
	<% End if %>
        document.frm.txtKBN.value = "";

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
<% If m_sKengen <> C_SEI0200_ACCESS_GAKKA then %>
	// ���N���X
	if( f_Trim(document.frm.txtClassNo.value) == "<%=C_CBO_NULL%>" ){
		window.alert("�N���X�̑I�����s���Ă�������");
		document.frm.txtClassNo.focus();
		return ;
	}
<% End If %>

	// ���敪��
	if( f_Trim(document.frm.txtKBN.value) == "<%=C_CBO_NULL%>" ){
		window.alert("�敪�̑I�����s���Ă�������");
		document.frm.txtKBN.focus();
		return ;
	}
	// ���w�N
	if( f_Trim(document.frm.txtGakuNo.value) == "" ){
		window.alert("�w�N�̑I�����s���Ă�������");
		document.frm.txtGakuNo.focus();
		return ;
	}

<% If m_sKengen <> C_SEI0200_ACCESS_GAKKA then %>
	// ���N���X
	if( f_Trim(document.frm.txtClassNo.value) == "" ){
		window.alert("�N���X�̑I�����s���Ă�������");
		document.frm.txtClassNo.focus();
		return ;
	}
<% End If %>

	// ���敪��
	if( f_Trim(document.frm.txtKBN.value) == "" ){
		window.alert("�敪�̑I�����s���Ă�������");
		document.frm.txtKBN.focus();
		return ;
	}

	document.frm.action="sei0200_main.asp";
	document.frm.target="main";
	document.frm.submit();

}
//-->
</SCRIPT>
<link rel=stylesheet href="../../common/style.css" type=text/css>
</head>

<body>
<center>
<form name="frm" METHOD="post" onClick="return false;">

<% call gs_title(" ���шꗗ "," ��@�� ") %>
<br>
<table border="0">
	<tr><td valign="bottom">

		<table border="0">
			<tr><td class="search">

				<table border="0">
					<tr>
						<td align="left">�����敪</td>
						<td><%call gf_ComboSet("txtSikenKBN",C_CBO_M01_KUBUN,m_sSikenKBNWhere, "style='width:150px;' ",False,m_sSikenKBN) %></td>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
					</tr>
					<tr>
						<td align="left">�N���X</td>
						<td><% if Cint(m_sKengen) = 1 then wDISABLED = "DISABLED" %>
							<%call gf_ComboSet("txtGakuNo",C_CBO_M05_CLASS_G,m_sGakuNoWhere,"style='width:40px;' onchange='javascript:f_ReLoadMyPage()' " & wDISABLED,True,m_sGakuNo)%>�N
<% If m_sKengen <> C_SEI0200_ACCESS_GAKKA then %>
							<%call gf_ComboSet("txtClassNo",C_CBO_M05_CLASS,m_sClassNoWhere,"style='width:80px;' "& m_sOption & wDISABLED ,True,m_sClassNo)%>
<%Else%>
							�@<%=gf_GetGakkaNm(m_iNendo,m_sGakkaNo)%>
<%End If%>
						</td>
						<td align="right">�@��@��</td>
						<td><%call gf_ComboSet("txtKBN",C_CBO_M01_KUBUN,m_sKBNWhere,"style='width:96px;' ",True,m_sKBN)%></td>
					</tr>
					<tr>
				        <td colspan="4" align="right">
				        <input type="button" class="button" value=" �N�@���@�A " onclick="javasript:f_Clear();">
				        <input type="button" class="button" value="�@�\�@���@" onclick="javasript:f_Search();">
				        </td>
					</tr>
				</table>
			</td></tr>
		</table>
	</tr>
</table>
<% if m_sKengen = C_SEI0200_ACCESS_TANNIN then %>
<input type="hidden" name="txtGakuNo"  value="<%=m_sGakuNo %>">
<input type="hidden" name="txtClassNo" value="<%=m_sClassNo%>">
<% ElseIf m_sKengen = C_SEI0200_ACCESS_GAKKA then %>
<input type="hidden" name="txtGakkaNo" value="<%=m_sGakkaNo%>">
<% End if %>
<input type="hidden" name="txtNendo" value="<%=m_iNendo%>">
<input type="hidden" name="txtKyokanCd" value="<%=m_sKyokanCd%>">
<input type="hidden" name="txtKengen" value="<%=m_sKengen%>">

</form>
</center>
</body>
</html>
<%
End Sub
%>