<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w����񌟍��ڍ�
' ��۸���ID : gak/gak0310/kojin_sita1.asp
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
' ��      ��: 2001/12/01 ���c
' ��      �X: 
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

	Public m_HyoujiFlg		'�\���׸�

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

		'//�\�����ڂ��擾
		w_iRet = f_GetDetailKojin()
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
'*  [�@�\]  �\�����ڂ��擾
'*  [����]  �Ȃ�
'*  [�ߒl]  0:����I��	1:�C�ӂ̃G���[  99:�V�X�e���G���[
'*  [����]  
'********************************************************************************
Function f_GetDetailKojin()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetDetailKojin = 1

	Do
		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & " 	A.T11_NYUNENDO,  "	'���w�N�x
		w_sSql = w_sSql & " 	A.T11_SEIBETU,  "	'����
		w_sSql = w_sSql & "		A.T11_NYUGAKU_KBN, " '���w�敪
		w_sSql = w_sSql & " 	A.T11_SEINENBI,  "	'���N����
		w_sSql = w_sSql & " 	A.T11_KETUEKI,  "	'���t�^
		w_sSql = w_sSql & " 	A.T11_RH,  "		'RH
		w_sSql = w_sSql & "		A.T11_SYUSSINKO,  "	'�o�g�Z
		w_sSql = w_sSql & "		A.T11_SYUSSINKOKU,  "	'�o�g��
		w_sSql = w_sSql & "		A.T11_RYUGAKU_KBN,  "	'���w�敪
		w_sSql = w_sSql & " 	A.T11_HOG_SIMEI,  "		'�ی�Ҏ���
		w_sSql = w_sSql & " 	A.T11_HOG_SIMEI_K,  "
		w_sSql = w_sSql & " 	A.T11_HOG_ZOKU,  "
		w_sSql = w_sSql & " 	A.T11_HOG_ZIP,  "
		w_sSql = w_sSql & " 	A.T11_HOG_JUSYO,  "
		w_sSql = w_sSql & " 	A.T11_HOGO_TEL,  "
		w_sSql = w_sSql & " 	A.T11_HOS_SIMEI,  "
		w_sSql = w_sSql & " 	A.T11_HOS_SIMEI_K,  "
		w_sSql = w_sSql & " 	A.T11_HOS_ZOKU,  "
		w_sSql = w_sSql & " 	A.T11_HOS_ZIP,  "
		w_sSql = w_sSql & " 	A.T11_HOS_JUSYO,  "
		w_sSql = w_sSql & " 	A.T11_HOS_TEL, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_1, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_ZOKU_1, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_2, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_ZOKU_2, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_3, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_ZOKU_3, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_4, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_ZOKU_4,"
		w_sSql = w_sSql & "		A.T11_KAZOKU_5,	"
		w_sSql = w_sSql & "		A.T11_KAZOKU_ZOKU_5, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_6,  "
		w_sSql = w_sSql & "		A.T11_KAZOKU_ZOKU_6, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_7,  "
		w_sSql = w_sSql & "		A.T11_KAZOKU_ZOKU_7, "
		w_sSql = w_sSql & "		A.T11_KAZOKU_8,  "
		w_sSql = w_sSql & "		A.T11_KAZOKU_ZOKU_8,"
		w_sSql = w_sSql & " 	A.T11_HOG_SEINEIBI,"
		w_sSql = w_sSql & " 	A.T11_HOS_SEINEIBI,"
		w_sSql = w_sSql & " 	A.T11_KAZOKU_SEINEIBI_1,"
		w_sSql = w_sSql & " 	A.T11_KAZOKU_SEINEIBI_2,"
		w_sSql = w_sSql & " 	A.T11_KAZOKU_SEINEIBI_3,"
		w_sSql = w_sSql & " 	A.T11_KAZOKU_SEINEIBI_4,"
		w_sSql = w_sSql & " 	A.T11_KAZOKU_SEINEIBI_5,"
		w_sSql = w_sSql & " 	A.T11_KAZOKU_SEINEIBI_6,"
		w_sSql = w_sSql & " 	A.T11_KAZOKU_SEINEIBI_7,"
		w_sSql = w_sSql & " 	A.T11_KAZOKU_SEINEIBI_8 "
		w_sSql = w_sSql & " FROM  "
		w_sSql = w_sSql & " 	T11_GAKUSEKI A "
		w_sSql = w_sSql & " WHERE "
		w_sSql = w_sSql & "  	A.T11_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' "

		iRet = gf_GetRecordset(m_Rs, w_sSql)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_GetDetailKojin = 99
			Exit Do
		End If

		'// ���ʂ��擾
		if Not gf_GetKubunName(C_SEIBETU,m_Rs("T11_SEIBETU"),Session("HyoujiNendo"),m_SEIBETU) then Exit Do

		'// ���t�^���擾
		if Not gf_GetKubunName(C_BLOOD,m_Rs("T11_KETUEKI"),Session("HyoujiNendo"),m_BLOOD) then Exit Do

		'// RH���擾
		if Not gf_GetKubunName(C_RH,m_Rs("T11_RH"),Session("HyoujiNendo"),m_RH) then Exit Do

		'// �ی�ґ������擾
		if Not gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_HOG_ZOKU"),Session("HyoujiNendo"),m_HOG_ZOKU) then Exit Do

		'// �ۏؐl�������擾
		if Not gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_HOS_ZOKU"),Session("HyoujiNendo"),m_HOS_ZOKU) then Exit Do

		'// ���w�敪���擾
		if Not gf_GetKubunName(C_RYUGAKU,m_Rs("T11_RYUGAKU_KBN"),Session("HyoujiNendo"),m_RYUGAKU_KBN) then Exit Do

		'//����I��
		f_GetDetailKojin = 0
		Exit Do
	Loop

End Function

'********************************************************************************
'*  [�@�\]  �o�g�Z���擾
'*  [����]  �w�ZCD
'*  [�ߒl]  �w�Z��
'********************************************************************************
Function f_GetSyussinko(p_Nendo,p_GakkouCd)

	On Error Resume Next
	Err.Clear

	f_GetSyussinko = ""

	'// �N�x�E�w�ZCD��NULL�������甲����
	if gf_IsNull(p_Nendo) or gf_IsNull(p_GakkouCd) then
		Exit Function
	End if

	w_sSql = "" 
	w_sSql = w_sSql & " SELECT "
	w_sSql = w_sSql & " 	M31_GAKKOMEI "
	w_sSql = w_sSql & " FROM "
	w_sSql = w_sSql & " 	M31_SYUSSINKO "
	w_sSql = w_sSql & " WHERE "
	w_sSql = w_sSql & " 	M31_NENDO    = '" & p_Nendo & "'"
	w_sSql = w_sSql & " AND M31_GAKKO_CD = '" & p_GakkouCd & "'"

	iRet = gf_GetRecordset(w_Rs, w_sSql)
	If iRet = 0 Then
		if Not w_Rs.Eof then
			f_GetSyussinko = w_Rs("M31_GAKKOMEI")
		End if
	End If

	p_oRecordset.Close
	Set p_oRecordset = Nothing

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

	m_SEINENBI 		= ""
	m_HOG_SIMEI		= ""
	m_HOG_SIMEI_K	= ""
	m_HOG_ZIP 		= ""
	m_HOG_JUSYO		= ""
	m_HOGO_TEL 		= ""
	m_HOS_SIMEI		= ""
	m_HOS_SIMEI_K	= ""
	m_HOS_ZIP		= ""
	m_HOS_JUSYO		= ""
	m_HOS_TEL		= ""
	m_KAZOKU_1      = ""
	m_KAZOKU_ZOKU_1 = ""
	m_KAZOKU_2      = ""
	m_KAZOKU_ZOKU_2 = ""
	m_KAZOKU_3      = ""
	m_KAZOKU_ZOKU_3 = ""
	m_KAZOKU_4      = ""
	m_KAZOKU_ZOKU_4 = ""
	m_KAZOKU_5      = ""
	m_KAZOKU_ZOKU_5 = ""
	m_KAZOKU_6      = ""
	m_KAZOKU_ZOKU_6 = ""
	m_KAZOKU_7      = ""
	m_KAZOKU_ZOKU_7 = ""
	m_KAZOKU_8      = ""
	m_KAZOKU_ZOKU_8 = ""
	m_SYUSSINKO    = ""
	m_SYUSSINKOKU   = ""
'	m_RYUGAKU_KBN   = ""

	m_Ken = ""
	m_SITYOSONCD = ""
	m_SITYOSONMEI = ""
	m_TYOIKIMEI = ""
	m_Ken2 = ""
	m_SITYOSONCD2 = ""
	m_SITYOSONMEI2 = ""
	m_TYOIKIMEI2 = ""

	m_HOG_SEINEIBI      = ""
	m_HOS_SEINEIBI      = ""
	m_KAZOKU_SEINEIBI_2 = ""
	m_KAZOKU_SEINEIBI_3 = ""
	m_KAZOKU_SEINEIBI_4 = ""
	m_KAZOKU_SEINEIBI_5 = ""
	m_KAZOKU_SEINEIBI_6 = ""
	m_KAZOKU_SEINEIBI_7 = ""
	m_KAZOKU_SEINEIBI_8 = ""

	'// �Ƒ������P�`�W���擾
	Call gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_KAZOKU_ZOKU_1"),Session("HyoujiNendo"),m_KAZOKU_ZOKU_1)
	Call gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_KAZOKU_ZOKU_2"),Session("HyoujiNendo"),m_KAZOKU_ZOKU_2)
	Call gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_KAZOKU_ZOKU_3"),Session("HyoujiNendo"),m_KAZOKU_ZOKU_3)
	Call gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_KAZOKU_ZOKU_4"),Session("HyoujiNendo"),m_KAZOKU_ZOKU_4)
	Call gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_KAZOKU_ZOKU_5"),Session("HyoujiNendo"),m_KAZOKU_ZOKU_5)
	Call gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_KAZOKU_ZOKU_6"),Session("HyoujiNendo"),m_KAZOKU_ZOKU_6)
	Call gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_KAZOKU_ZOKU_7"),Session("HyoujiNendo"),m_KAZOKU_ZOKU_7)
	Call gf_GetKubunName(C_ZOKUGARA,m_Rs("T11_KAZOKU_ZOKU_8"),Session("HyoujiNendo"),m_KAZOKU_ZOKU_8)

	Call gf_ComvZip(m_Rs("T11_HOG_ZIP"),m_Ken,m_SITYOSONCD,m_SITYOSONMEI,m_TYOIKIMEI,Session("HyoujiNendo"))
	Call gf_ComvZip(m_Rs("T11_HOS_ZIP"),m_Ken2,m_SITYOSONCD2,m_SITYOSONMEI2,m_TYOIKIMEI2,Session("HyoujiNendo"))


 	if Not m_Rs.EOF then
		m_SEINENBI 		= m_Rs("T11_SEINENBI") 
		m_HOG_SIMEI		= m_Rs("T11_HOG_SIMEI") 
		m_HOG_SIMEI_K	= m_Rs("T11_HOG_SIMEI_K")
		m_HOG_ZIP 		= m_Rs("T11_HOG_ZIP")

	'/* �Z���Ɍ��A�s�����������݂��Ă����ꍇ�폜���čēx�t�������B*/ Add 2002.04.30 okada
		m_HOG_JUSYO		= m_Rs("T11_HOG_JUSYO")
		m_HOG_JUSYO     = Replace(m_HOG_JUSYO,m_Ken,"")
		m_HOG_JUSYO     = Replace(m_HOG_JUSYO,m_SITYOSONMEI,"")

		m_HOG_JUSYO		= m_Ken & m_SITYOSONMEI & m_HOG_JUSYO 'm_SITYOSONMEI & Replace(m_Rs("T11_HOG_JUSYO"),m_SITYOSONMEI,"")

		
		m_HOGO_TEL 		= m_Rs("T11_HOGO_TEL") 
		m_HOS_SIMEI		= m_Rs("T11_HOS_SIMEI") 
		m_HOS_SIMEI_K	= m_Rs("T11_HOS_SIMEI_K") 
		m_HOS_ZIP		= m_Rs("T11_HOS_ZIP")
		m_HOS_JUSYO		= m_Rs("T11_HOS_JUSYO")
	
	'/* �Z���Ɍ��A�s�����������݂��Ă����ꍇ�폜���čēx�t�������B*/ Add 2002.04.30 okada
		m_HOS_JUSYO		= m_Rs("T11_HOS_JUSYO")
		m_HOS_JUSYO		= Replace(m_HOS_JUSYO,m_Ken2,"")
		m_HOS_JUSYO		= Replace(m_HOS_JUSYO,m_SITYOSONMEI2,"")

		m_HOS_JUSYO		= m_Ken2 & m_SITYOSONMEI2 & m_HOS_JUSYO'm_SITYOSONMEI2 & Replace(m_Rs("T11_HOS_JUSYO"),m_SITYOSONMEI2,"")

		'm_HOS_TEL		= m_Rs("T11_HOS_TEL")
		'_KAZOKU_ZOKU_1 = m_Rs("T11_KAZOKU_ZOKU_1")
		'_KAZOKU_ZOKU_2 = m_Rs("T11_KAZOKU_ZOKU_2")
		'_KAZOKU_ZOKU_3 = m_Rs("T11_KAZOKU_ZOKU_3")
		'_KAZOKU_ZOKU_4 = m_Rs("T11_KAZOKU_ZOKU_4")
		'_KAZOKU_ZOKU_5 = m_Rs("T11_KAZOKU_ZOKU_5")
		'_KAZOKU_ZOKU_6 = m_Rs("T11_KAZOKU_ZOKU_6")
		'_KAZOKU_ZOKU_7 = m_Rs("T11_KAZOKU_ZOKU_7")
		'_KAZOKU_ZOKU_8 = m_Rs("T11_KAZOKU_ZOKU_8")
		'm_RYUGAKU_KBN  = m_Rs("T11_RYUGAKU_KBN")
		m_KAZOKU_1      = m_Rs("T11_KAZOKU_1")
		m_KAZOKU_2      = m_Rs("T11_KAZOKU_2")
		m_KAZOKU_3      = m_Rs("T11_KAZOKU_3")
		m_KAZOKU_4      = m_Rs("T11_KAZOKU_4")
		m_KAZOKU_5      = m_Rs("T11_KAZOKU_5")
		m_KAZOKU_6      = m_Rs("T11_KAZOKU_6")
		m_KAZOKU_7      = m_Rs("T11_KAZOKU_7")
		m_KAZOKU_8      = m_Rs("T11_KAZOKU_8")
		m_SYUSSINKO     = f_GetSyussinko(Session("HyoujiNendo"),m_Rs("T11_SYUSSINKO"))
		m_SYUSSINKOKU   = m_Rs("T11_SYUSSINKOKU")

		m_HOG_SEINEIBI      =   m_Rs("T11_HOG_SEINEIBI")
		m_HOS_SEINEIBI      =   m_Rs("T11_HOS_SEINEIBI")
		m_KAZOKU_SEINEIBI_1 =   m_Rs("T11_KAZOKU_SEINEIBI_1")
		m_KAZOKU_SEINEIBI_2 =   m_Rs("T11_KAZOKU_SEINEIBI_2")
		m_KAZOKU_SEINEIBI_3 =   m_Rs("T11_KAZOKU_SEINEIBI_3")
		m_KAZOKU_SEINEIBI_4 =   m_Rs("T11_KAZOKU_SEINEIBI_4")
		m_KAZOKU_SEINEIBI_5 =   m_Rs("T11_KAZOKU_SEINEIBI_5")
		m_KAZOKU_SEINEIBI_6 =   m_Rs("T11_KAZOKU_SEINEIBI_6")
		m_KAZOKU_SEINEIBI_7 =   m_Rs("T11_KAZOKU_SEINEIBI_7")
		m_KAZOKU_SEINEIBI_8 =   m_Rs("T11_KAZOKU_SEINEIBI_8")

	End if

%>

	<html>
	<head>
	<title>�w�Ѓf�[�^�Q��</title>
	<meta http-equiv="Content-Type" content="text/html; charset=x-sjis">
    <link rel=stylesheet href=../../common/style.css type=text/css>
	<script language="javascript">
	<!--
		function sbmt(m,i) {
			document.forms[0].mode.value = m;
			document.forms[0].id.value = i;
			document.forms[0].submit();
		}
	//-->
	</script>
	<style type="text/css">
	<!--
		a:link { color:#cc8866; text-decoration:none; }
		a:visited { color:#cc8866; text-decoration:none; }
		a:active { color:#888866; text-decoration:none; }
		a:hover { color:#888866; text-decoration:underline; }
		b { color:#88bbbb; font-weight: bold; font-size:14px}
	//-->
	</style>
	</head>

	<body>
	<form action="main.asp" method="post" name="frm" target="fMain">
	<div align="center">

	<br><br>
	<table border="0" cellpadding="0" cellspacing="0" width="600">
		<tr>
			<td nowrap><a href="kojin_sita0.asp">����{���</a></td>
			<td nowrap><b>���l���</b></td>
			<td nowrap><a href="kojin_sita2.asp">�����w���</a></td>
			<td nowrap><a href="kojin_sita3.asp">���w�N���</a></td>
			<td nowrap><a href="kojin_sita4.asp">�����̑��\�����</a></td>
			<td nowrap><a href="kojin_sita5.asp">���ٓ����</a></td>
		</tr>
	</table>
	<br>

	<table border="0" cellpadding="1" cellspacing="1">
		<tr>
			<td valign="top" align="left">

					<table border="1" width="220" class="disp">
						<% if gf_empItem(C_T11_SEIBETU) then %>
							<tr>
								<td width="100" height="16" class="disph">���@�@��</td>
								<td class="disp"><%= m_SEIBETU %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_SEINENBI) then %>
							<tr>
								<td height="16" class="disph">���N����</td>
								<td class="disp"><%= m_SEINENBI %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_KETUEKI) then %>
							<tr>
								<td height="16" class="disph">�� �t �^</td>
								<td class="disp"><%= m_BLOOD %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_RH) then %>
							<tr>
								<td height="16" class="disph">�q�@�@�g</td>
								<td class="disp"><%= m_RH %>&nbsp</td>
							</tr>
						<% End if %>
						<% if Cint(gf_SetNull2Zero(m_Rs("T11_NYUGAKU_KBN"))) = C_NYU_RYUGAKU then %>
							<% if gf_empItem(C_T11_SYUSSINKO) then %>
								<tr>
									<td height="16" class="disph">�o �g �Z</td>
									<td class="disp"><%= m_SYUSSINKO %>&nbsp</td>
								</tr>
							<% End if %>
							<% if gf_empItem(C_T11_SYUSSINKOKU) then %>
								<tr>
									<td height="16" class="disph">�o �g ��</td>
									<td class="disp"><%= m_SYUSSINKOKU %>&nbsp</td>
								</tr>
							<% End if %>
							<% if gf_empItem(C_T11_RYUGAKU_KBN) then %>
								<tr>
									<td height="16" class="disph">�� �w �� ��</td>
									<td class="disp"><%= m_RYUGAKU_KBN %>&nbsp</td>
								</tr>
							<% End if %>
						<% End if %>
					</table>
			</td>

			<td valign="top" align="left">

					<table border="1" width="220" class="disp">
						<% if gf_empItem(C_T11_HOG_SIMEI) then %>
							<tr>
								<td class="disph" width="110" height="16"><font color="white">�ی�Ҏ���</font></td>
								<td class="disp"  width="110"><%= m_HOG_SIMEI %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_SIMEI_K) then %>
							<tr>
								<td class="disph" height="16"><font color="white">�ی�҃J�i</font></td>
								<td class="disp"><%= m_HOG_SIMEI_K %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_ZOKU) then %>
							<tr>
								<td class="disph" height="16"><font color="white">�ی�ґ���</font></td>
								<td class="disp"><%= m_HOG_ZOKU %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_SEINEIBI) then %>
							<tr>
								<td class="disph" height="16"><font color="white">�ی�Ґ��N����</font></td>
								<td class="disp"><%= m_HOG_SEINEIBI %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_ZIP) then %>
							<tr>
								<td class="disph" height="16"><font color="white">�ی�ҁ�</font></td>
								<td class="disp"><%= m_HOG_ZIP %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_JUSYO) then %>
							<tr>
								<td class="disph" height="16"><font color="white">�ی�ҏZ��</font></td>
								<td class="disp"><%= m_HOG_JUSYO %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_TEL) then %>
							<tr>
								<td class="disph" height="16"><font color="white">�ی��TEL</font></td>
								<td class="disp"><%= m_HOGO_TEL %>&nbsp</td>
							</tr>
						<% End if %>
					</table>
			<br>

			</td>
			<td valign="top" align="left">
					<table border="1" width="220" class="disp">
						<% if gf_empItem(C_T11_HOS_SIMEI) then %>
							<tr>
								<td class="disph" width="110" height="16">�ۏؐl����</td>
								<td class="disp"  width="110"><%= m_HOS_SIMEI %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_SIMEI_K) then %>
							<tr>
								<td class="disph" height="16">�ۏؐl�J�i</td>
								<td class="disp"><%= m_HOS_SIMEI_K %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_ZOKU) then %>
							<tr>
								<td class="disph" height="16">�ۏؐl����</td>
								<td class="disp"><%= m_HOS_ZOKU %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_SEINEIBI) then %>
							<tr>
								<td class="disph" height="16">�ۏؐl���N����</td>
								<td class="disp"><%= m_HOS_SEINEIBI %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_ZIP) then %>
							<tr>
								<td class="disph" height="16">�ۏؐl��</td>
								<td class="disp"><%= m_HOS_ZIP %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_JUSYO) then %>
							<tr>
								<td class="disph" height="16">�ۏؐl�Z��</td>
								<td class="disp"><%= m_HOS_JUSYO %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_TEL) then %>
							<tr>
								<td class="disph" height="16">�ۏؐlTEL</td>
								<td class="disp"><%= m_HOS_TEL %>&nbsp</td>
							</tr>
						<% End if %>
					</table>
					
			</td>

		</tr>
	<tr><td colspan=3>

					<table border="1" class="disp">
						<% if gf_empItem(C_T11_KAZOKU_1) then %>
							<tr>
								<td class="disph" width="100" height="16">�Ƒ����̂P</td>
								<td class="disp" width="120" ><%= m_KAZOKU_1 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_ZOKU_1) then %>
								<td class="disph" width="60" height="16"> �� �� �P</td>
								<td class="disp" width="60"><%= m_KAZOKU_ZOKU_1 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_SEINEIBI_1) then %>
								<td class="disph" width="100" height="16">���N�����P</td>
								<td class="disp" width="100"><%= m_KAZOKU_SEINEIBI_1 %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_2) then %>
							<tr>
								<td class="disph" height="16">�Ƒ����̂Q</td>
								<td class="disp"><%= m_KAZOKU_2 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_ZOKU_2) then %>
								<td class="disph" height="16"> �� �� �Q</td>
								<td class="disp"><%= m_KAZOKU_ZOKU_2 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_SEINEIBI_2) then %>
								<td class="disph" height="16">���N�����Q</td>
								<td class="disp"><%= m_KAZOKU_SEINEIBI_2 %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_3) then %>
							<tr>
								<td class="disph" height="16">�Ƒ����̂R</td>
								<td class="disp"><%= m_KAZOKU_3 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_ZOKU_3) then %>
								<td class="disph" height="16"> �� �� �R</td>
								<td class="disp"><%= m_KAZOKU_ZOKU_3 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_SEINEIBI_3) then %>
								<td class="disph" height="16">���N�����R</td>
								<td class="disp"><%= m_KAZOKU_SEINEIBI_3 %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_4) then %>
							<tr>
								<td class="disph" height="16">�Ƒ����̂S</td>
								<td class="disp"><%= m_KAZOKU_4 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_ZOKU_4) then %>
								<td class="disph" height="16"> �� �� �S</td>
								<td class="disp"><%= m_KAZOKU_ZOKU_4 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_SEINEIBI_4) then %>
								<td class="disph" height="16">���N�����S</td>
								<td class="disp"><%= m_KAZOKU_SEINEIBI_4 %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_5) then %>
							<tr>
								<td class="disph" height="16">�Ƒ����̂T</td>
								<td class="disp"><%= m_KAZOKU_5 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_ZOKU_5) then %>
								<td class="disph" height="16"> �� �� �T</td>
								<td class="disp"><%= m_KAZOKU_ZOKU_5 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_SEINEIBI_5) then %>
								<td class="disph" height="16">���N�����T</td>
								<td class="disp"><%= m_KAZOKU_SEINEIBI_5 %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_6) then %>
							<tr>
								<td class="disph" height="16">�Ƒ����̂U</td>
								<td class="disp"><%= m_KAZOKU_6 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_ZOKU_6) then %>
								<td class="disph" height="16"> �� �� �U</td>
								<td class="disp"><%= m_KAZOKU_ZOKU_6 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_SEINEIBI_6) then %>
								<td class="disph" height="16">���N�����U</td>
								<td class="disp"><%= m_KAZOKU_SEINEIBI_6 %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_7) then %>
							<tr>
								<td class="disph" height="16">�Ƒ����̂V</td>
								<td class="disp"><%= m_KAZOKU_7 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_ZOKU_7) then %>
								<td class="disph" height="16"> �� �� �V</td>
								<td class="disp"><%= m_KAZOKU_ZOKU_7 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_SEINEIBI_7) then %>
								<td class="disph" height="16">���N�����V</td>
								<td class="disp"><%= m_KAZOKU_SEINEIBI_7 %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_8) then %>
							<tr>
								<td class="disph" height="16">�Ƒ����̂W</td>
								<td class="disp"><%= m_KAZOKU_8 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_ZOKU_8) then %>
								<td class="disph" height="16"> �� �� �W</td>
								<td class="disp"><%= m_KAZOKU_ZOKU_8 %>&nbsp</td>
						<% End if %>
						<% if gf_empItem(C_T11_KAZOKU_SEINEIBI_8) then %>
								<td class="disph" height="16">���N�����W</td>
								<td class="disp"><%= m_KAZOKU_SEINEIBI_8 %>&nbsp</td>
							</tr>
						<% End if %>
					</table>

	</td></tr>
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