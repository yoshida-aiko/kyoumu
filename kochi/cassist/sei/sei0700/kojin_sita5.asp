<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w����񌟍��ڍ�
' ��۸���ID : gak/gak0300/kojin_sita5.asp
' �@      �\: �������ꂽ�w���̏ڍׂ�\������(�ٓ����)
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
		w_iRet = f_GetDetailIdo()
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
Function f_GetDetailIdo()
	Dim w_iRet
	Dim w_sSQL

	On Error Resume Next
	Err.Clear

	f_GetDetailIdo = 1

	Do

		w_sSql = ""
		w_sSql = w_sSql & " SELECT "
		w_sSql = w_sSql & " 	A.T13_IDOU_NUM, "
		w_sSql = w_sSql & " 	A.T13_NENDO, "
		w_sSql = w_sSql & " 	A.T13_GAKUNEN, "
		w_sSql = w_sSql & " 	A.T13_IDOU_KBN_1, 	A.T13_IDOU_BI_1, 	A.T13_IDOU_BIK_1, "
		w_sSql = w_sSql & " 	A.T13_IDOU_KBN_2, 	A.T13_IDOU_BI_2, 	A.T13_IDOU_BIK_2, "
		w_sSql = w_sSql & " 	A.T13_IDOU_KBN_3, 	A.T13_IDOU_BI_3, 	A.T13_IDOU_BIK_3, "
		w_sSql = w_sSql & " 	A.T13_IDOU_KBN_4, 	A.T13_IDOU_BI_4, 	A.T13_IDOU_BIK_4, "
		w_sSql = w_sSql & " 	A.T13_IDOU_KBN_5, 	A.T13_IDOU_BI_5, 	A.T13_IDOU_BIK_5, "
		w_sSql = w_sSql & " 	A.T13_IDOU_KBN_6, 	A.T13_IDOU_BI_6, 	A.T13_IDOU_BIK_6, "
		w_sSql = w_sSql & " 	A.T13_IDOU_KBN_7, 	A.T13_IDOU_BI_7, 	A.T13_IDOU_BIK_7, "
		w_sSql = w_sSql & " 	A.T13_IDOU_KBN_8, 	A.T13_IDOU_BI_8, 	A.T13_IDOU_BIK_8 "
		w_sSql = w_sSql & " FROM "
		w_sSql = w_sSql & " 	T13_GAKU_NEN A "
		w_sSql = w_sSql & " WHERE "
		w_sSql = w_sSql & " 	A.T13_GAKUSEI_NO = '" & Session("GAKUSEI_NO") & "' "
		w_sSql = w_sSql & " AND A.T13_NENDO = " & Session("HyoujiNendo")

		iRet = gf_GetRecordset(m_Rs, w_sSQL)
		If iRet <> 0 Then
			'ں��޾�Ă̎擾���s
			msMsg = Err.description
			f_GetDetailIdo = 99
			Exit Do
		End If


		'//����I��
		f_GetDetailIdo = 0
		Exit Do
	Loop

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

	m_IDOU_NUM		= ""
	m_NENDO			= ""
	m_GAKUNEN		= ""

	if Not m_Rs.Eof then
		m_IDOU_NUM		= m_Rs("T13_IDOU_NUM")
		m_NENDO			= m_Rs("T13_NENDO")
		m_GAKUNEN		= m_Rs("T13_GAKUNEN")
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
		function sbmt(m,i) {
			document.forms[0].mode.value = m;
			document.forms[0].id.value = i;
			document.forms[0].submit();
		}
	//-->
	</script>
	</head>

	<body>
	<form action="main.asp" method="post" name="frm" target="fMain">
	<div align="center">

	<br><br>
	<table border="0" cellpadding="0" cellspacing="0" width="600">
		<tr>
			<td nowrap><a href="kojin_sita0.asp">����{���</a></td>
			<td nowrap><a href="kojin_sita1.asp">���l���</a></td>
			<td nowrap><a href="kojin_sita2.asp">�����w���</a></td>
			<td nowrap><a href="kojin_sita3.asp">���w�N���</a></td>
			<td nowrap><a href="kojin_sita4.asp">�����l�E����</a></td>
			<td nowrap><b>���ٓ����</b></td>
		</tr>
	</table>
	<br>

	<table border="1" class=hyo>
		<tr>
			<% if gf_empItem(C_T13_IDOU_KBN) then %>
				<th width="80"  class="header">���R</th>
			<% End if %>
			<% if gf_empItem(C_T13_NENDO) then %>
				<th width="100" class="header">�N�x</th>
			<% End if %>
			<% if gf_empItem(C_T13_GAKUNEN) then %>
				<th width="100" class="header">�w�N</th>
			<% End if %>
			<% if gf_empItem(C_T13_IDOU_BI) then %>
				<th width="160" class="header">���t�܂��͊���</th>
			<% End if %>
			<% if gf_empItem(C_T13_IDOU_BIK) then %>
				<th width="140" class="header">���l</th>
			<% End if %>
		</tr>
		<%
			'// �ړ��񐔕�,��
			if Cint(gf_SetNull2Zero(m_IDOU_NUM)) > 0 then
				i_line = 1
				Do Until i_line > Cint(m_IDOU_NUM)

					'// ���R���擾
					m_IDOU_KBN = ""
					w_IDOU_KBN_NO = m_Rs("T13_IDOU_KBN_" & i_line)
					Call gf_GetKubunName(C_IDO,w_IDOU_KBN_NO,Session("HyoujiNendo"),m_IDOU_KBN)
					Call gs_cellPtn(w_cell)
				%>
					<tr>
						<% if gf_empItem(C_T13_IDOU_KBN) then %>
							<td class="<%=w_cell%>"><%= m_IDOU_KBN %></td>
						<% End if %>
						<% if gf_empItem(C_T13_NENDO) then %>
							<td class="<%=w_cell%>"><%= m_NENDO %></td>
						<% End if %>
						<% if gf_empItem(C_T13_GAKUNEN) then %>
							<td class="<%=w_cell%>"><%= m_GAKUNEN %></td>
						<% End if %>
						<% if gf_empItem(C_T13_IDOU_BI) then %>
							<td class="<%=w_cell%>"><%= m_Rs("T13_IDOU_BI_" & i_line ) %></td>
						<% End if %>
						<% if gf_empItem(C_T13_IDOU_BIK) then %>
							<td class="<%=w_cell%>"><%= m_Rs("T13_IDOU_BIK_" & i_line ) %></td>
						<% End if %>
					</tr>
				<%
					i_line = i_line + 1
				Loop
			End if
		%>
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