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
		w_sSql = w_sSql & " 	A.T11_SEIBETU,  "
		w_sSql = w_sSql & " 	A.T11_SEINENBI,  "
		w_sSql = w_sSql & " 	A.T11_KETUEKI,  "
		w_sSql = w_sSql & " 	A.T11_RH,  "
		w_sSql = w_sSql & " 	A.T11_HOG_SIMEI,  "
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
		w_sSql = w_sSql & " 	A.T11_HOS_TEL "
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

		'//����I��
		f_GetDetailKojin = 0
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

 	if Not m_Rs.EOF then
		m_SEINENBI 		= m_Rs("T11_SEINENBI") 
		m_HOG_SIMEI		= m_Rs("T11_HOG_SIMEI") 
		m_HOG_SIMEI_K	= m_Rs("T11_HOG_SIMEI_K") 
		m_HOG_ZIP 		= m_Rs("T11_HOG_ZIP") 
		m_HOG_JUSYO		= m_Rs("T11_HOG_JUSYO") 
		m_HOGO_TEL 		= m_Rs("T11_HOGO_TEL") 
		m_HOS_SIMEI		= m_Rs("T11_HOS_SIMEI") 
		m_HOS_SIMEI_K	= m_Rs("T11_HOS_SIMEI_K") 
		m_HOS_ZIP		= m_Rs("T11_HOS_ZIP")
		m_HOS_JUSYO		= m_Rs("T11_HOS_JUSYO")
		m_HOS_TEL		= m_Rs("T11_HOS_TEL")
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
			<td nowrap><a href="kojin_sita4.asp">�����l�E����</a></td>
			<td nowrap><a href="kojin_sita5.asp">���ٓ����</a></td>
		</tr>
	</table>
	<br>

	
	<table border="0" cellpadding="1" cellspacing="1">
		<tr>
			<td valign="top" align="left">
				<br>

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
					</table>

			</td>

			<td valign="top" align="left">
				�y �ی�ҏ�� �z

					<table border="1" width="220" class="disp">
						<% if gf_empItem(C_T11_HOG_SIMEI) then %>
							<tr>
								<td class="disph" width="100" height="16"><font color="white">���@�@��</font></td>
								<td class="disp"><%= m_HOG_SIMEI %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_SIMEI_K) then %>
							<tr>
								<td class="disph" height="16"><font color="white">�J�@�@�i</font></td>
								<td class="disp"><%= m_HOG_SIMEI_K %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_ZOKU) then %>
							<tr>
								<td class="disph" height="16"><font color="white">���@�@��</font></td>
								<td class="disp"><%= m_HOG_ZOKU %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_ZIP) then %>
							<tr>
								<td class="disph" height="16"><font color="white">��</font></td>
								<td class="disp"><%= m_HOG_ZIP %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_JUSYO) then %>
							<tr>
								<td class="disph" height="16"><font color="white">�Z�@�@��</font></td>
								<td class="disp"><%= m_HOG_JUSYO %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOG_TEL) then %>
							<tr>
								<td class="disph" height="16"><font color="white">�s �d �k</font></td>
								<td class="disp"><%= m_HOGO_TEL %>&nbsp</td>
							</tr>
						<% End if %>
					</table>

			</td>

			<td valign="top" align="left">
				�y �ۏؐl��� �z

					<table border="1" width="220" class="disp">
						<% if gf_empItem(C_T11_HOS_SIMEI) then %>
							<tr>
								<td class="disph" width="100" height="16">���@�@��</td>
								<td class="disp"><%= m_HOS_SIMEI %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_SIMEI_K) then %>
							<tr>
								<td class="disph" height="16">�J�@�@�i</td>
								<td class="disp"><%= m_HOS_SIMEI_K %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_ZOKU) then %>
							<tr>
								<td class="disph" height="16">���@�@��</td>
								<td class="disp"><%= m_HOS_ZOKU %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_ZIP) then %>
							<tr>
								<td class="disph" height="16">��</td>
								<td class="disp"><%= m_HOS_ZIP %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_JUSYO) then %>
							<tr>
								<td class="disph" height="16">�Z�@�@��</td>
								<td class="disp"><%= m_HOS_JUSYO %>&nbsp</td>
							</tr>
						<% End if %>
						<% if gf_empItem(C_T11_HOS_TEL) then %>
							<tr>
								<td class="disph" height="16">�s �d �k</td>
								<td class="disp"><%= m_HOS_TEL %>&nbsp</td>
							</tr>
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