<%@Language=VBScript %>
<%
'******************************************************************
'�V�X�e����     �F���������V�X�e��
'���@���@��     �F�l���C�I���Ȗڌ���
'�v���O����ID   �Fweb/web0340/default.asp
'�@�@�@�@�\     �F�t���[���y�[�W �l���C�I���Ȗڌ���̕\�����s��
'------------------------------------------------------------------
'���@�@�@��     �F
'�ρ@�@�@��     �F
'���@�@�@�n     �F
'���@�@�@��     �F
'------------------------------------------------------------------
'��@�@�@��     �F2001.07.23    �O�c�@�q�j
'�ρ@�@�@�X     �F
'
'******************************************************************
'*******************�@ASP���ʃ��W���[���錾�@**********************
%>
<!--#include file="../../common/com_All.asp"-->
<%
'******�@�� �W �� �[ �� �� ���@********
Public m_iNendo
Public m_sKyokanCd
'******�@���C�������@********

    'Ҳ�ٰ�ݎ��s
    Call Main()

'******�@�d�@�m�@�c�@********

Sub Main()
'******************************************************************
'�@�@�@�\�F�{ASP��Ҳ�ٰ��
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************

    '******���ʊ֐�******
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�A�E�}�X�^"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False
    m_iNendo = session("NENDO")
    m_sKyokanCd = session("KYOKAN_CD")

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
		session("PRJ_No") = "WEB0340"

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

        '// �S�C�`�F�b�N
'	  If gf_Tannin(m_iNendo,m_sKyokanCd,1) <> 0 Then
'	            m_bErrFlg = True
'	            m_sErrMsg = "�S�C�ȊO�̓��͂͂ł��܂���B"
'	            Exit Do
'	  End If


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
    <title>�l���C�I���Ȗڌ���</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <script language=javascript>
    </script>
    <frameset rows=170,1,* frameborder="0">
        <frame src="web0340_top.asp" scrolling="auto" noresize name="top">
        <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
        <frame src="default2.asp" scrolling="auto" noresize name="main">
    </frameset>
    </head>
</html>
<%
End Sub
%>