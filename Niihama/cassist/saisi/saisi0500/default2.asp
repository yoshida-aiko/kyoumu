<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �w���ꗗ
' ��۸���ID : saisi/saisi�~�~�~/default.asp
' �@      �\: �S�C�����������k�̈ꗗ���Q�Ƃ���
'-------------------------------------------------------------------------
' ��      ��:
' ��      ��:
' ��      �n:
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2003/02/24
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

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
	w_sWinTitle = "�L�����p�X�A�V�X�g"
	w_sMsgTitle = "�s���i�w���ꗗ"
	w_sMsg = ""
	w_sRetURL= C_RetURL & C_ERR_RETURL
	w_sTarget = "fTopMain"

	On Error Resume Next
	Err.Clear

	m_bErrFlg = False

	Do
		'// �����`�F�b�N�Ɏg�p
		session("PRJ_No") = C_LEVEL_NOCHK '(�����`�F�b�N�����Ȃ�)

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

		'// �����y�[�W��\��
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

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
	dim w_item
	'---------- HTML START  ----------
	On Error Resume Next
	Err.Clear
%>

<html>
<head>
<title><%= w_sMsgTitle %></title>
</head>

<frameset rows=120,1,* frameborder="no">
	<frame src="saisi0500_top2.asp" scrolling="auto" name="_TOP" noresize>
	<frame src="../../common/bar.html" scrolling="auto" name="bar" noresize>
	<frame src="saisi0500_lower.asp?mode=new" scrolling="auto" name="_LOWER" noresize>
</frameset>

</html>

<% End Sub %>