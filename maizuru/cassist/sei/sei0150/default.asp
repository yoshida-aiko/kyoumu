<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ѓo�^
' ��۸���ID : sei/sei0150/default.asp
' �@      �\: 
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h		��		SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h		��		SESSION���i�ۗ��j
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/07/26 �O�c�@�q�j
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
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="���ѓo�^"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"     
	w_sTarget="_parent"


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

		'// �����`�F�b�N�Ɏg�p
		session("PRJ_No") = "SEI0100"

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

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
<!--#include file="../../Common/scroll.js"-->
<title>���ѓo�^</title>
</head>

<frameset rows=245px,1,* frameborder="0" framespacing="0">
	<frame src="sei0150_top.asp" scrolling="yes"  name="topFrame" noresize>
    <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
	<frame src="default2.asp" scrolling="yes"  name="main" noresize>
</frameset>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
