<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �������{�Ȗړo�^
' ��۸���ID : skn/skn0130/default.asp
' �@      �\: �t���[���y�[�W �����\��}�X�^�̎Q�Ƃ��s��
'-------------------------------------------------------------------------
' ��      ��:
' ��      ��:
' ��      �n:
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/06/18
' ��      �X: 2001/06/26
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
    w_sMsgTitle = "�������{�Ȗړo�^"
    w_sMsg = ""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget = ""

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
		session("PRJ_No") = "SKN0130"

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
    
	w_item = ""
    w_item = w_item & "txtMode="&request("txtMode")
    w_item = w_item & "&txtSikenKbn="&request("txtSikenKbn")

%>
<html>
<head>
<% '�^�C�g���������FireFox�ŕ����������邽�ߍ폜 --2019/06/24 Del Fujibayashi <title>�������{�Ȗړo�^</title> %>
</head>

<frameset rows=120,1,* frameborder="no">
	<frame src="SKN0130_top.asp?<%=request.form.item%>" scrolling="auto" noresize>
    <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
	<frame src="SKN0130_main.asp?<%=w_item%>" scrolling="auto" noresize name=main>
</frameset>

</html>
<% End Sub %>