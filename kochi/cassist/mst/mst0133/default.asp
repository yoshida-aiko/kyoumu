<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �A�E��}�X�^
' ��۸���ID : mst/mst0133/default.asp
' �@      �\: �t���[���y�[�W �A�E��}�X�^�̎Q�Ƃ��s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' �@      �@:session("PRJ_No")      '���������̃L�[ '/2001/07/31�ǉ�
'           :txtSinroCD             :�i�H�R�[�h     '/2001/07/31�ǉ�
'           :txtSingakuCD           :�i�w�R�[�h     '/2001/07/31�ǉ�
'           :txtSyusyokuName        :�A�E�於�́i�ꕔ�j '/2001/07/31�ǉ�
'           :txtPageCD              :�\���Ő�           '/2001/07/31�ǉ�
'           :txtMode                :���[�h             '/2001/07/31�ǉ�
'           :txtSentakuSinroCD
'           :txtFLG
'           :txtSNm
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/06/18 �≺�@�K��Y
' ��      �X: 2001/07/31 ���{ ����  �����E���n�ǉ�
'           :                       �ϐ��������K���Ɋ�ύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_sMode             '���[�h
    Public  m_iSinroCD          '�i�H�敪           '/2001/07/31�ύX
    Public  m_iSingakuCd        '�i�w�敪           '/2001/07/31�ύX
    Public  m_sSyusyokuName     '�A�E�於�́i�ꕔ�j
    Public  m_sPageCD           '�\���Ő�
    Public  m_sSentakuSinroCD   
    Public  m_iFLG
    Public  m_sSNm

    Public  m_sArg              '����   '/2001/07/31�ύX
    Public  m_sArg_top          '����   '/2001/07/31�ύX

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
    w_sMsgTitle="�i�H���񌟍�"
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

        '// �����`�F�b�N�Ɏg�p
        session("PRJ_No") = "MST0133"

        '// �s���A�N�Z�X�`�F�b�N
        Call gf_userChk(session("PRJ_No"))

        '// ���Ұ�SET
        Call s_SetParam()

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
'*  [�@�\]  �C�ӂ̃y�[�W�փp�����[�^��n��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_SetParam()

    m_iSinroCD = Request("txtSinroCD")      ':�i�H�敪
    '�R���{���I����
    If m_iSinroCD="@@@" Then
        m_iSinroCD=""
    End If

    m_iSingakuCd = Request("txtSingakuCD")      ':�i�w�R�[�h
    '�R���{���I����
    If m_iSingakuCd="@@@" Then
        m_iSingakuCd=""
    End If

    m_sSyusyokuName = Request("txtSyusyokuName")':�A�E�於�́i�ꕔ�j

    m_sPageCD = Request("txtPageCD")        ':�\�������y�[�W

    m_sMode = Request("txtMode")            ':���[�h

    m_iFLG = request("txtFLG")
    m_sSNm = request("txtSNm")

    m_sArg = "?"
    m_sArg = m_sArg & "txtMode=" & m_sMode 
    m_sArg = m_sArg & "&txtSinroCD=" & m_iSinroCD 
    m_sArg = m_sArg & "&txtSingakuCD=" & m_iSingakuCd 
    m_sArg = m_sArg & "&txtSyusyokuName=" & Server.URLEncode(m_sSyusyokuName) 
    m_sArg = m_sArg & "&txtPageCD=" & m_sPageCD 
    m_sArg = m_sArg & "&txtSentakuSinroCD=" & m_sSentakuSinroCD 

    m_sArg_top = "?"
    m_sArg_top = m_sArg_top & "txtSinroCD=" & m_iSinroCD 
    m_sArg_top = m_sArg_top & "&txtSingakuCD=" & m_iSingakuCd 
    m_sArg_top = m_sArg_top & "&txtSyusyokuName=" & Server.URLEncode(m_sSyusyokuName) 
    m_sArg_top = m_sArg_top & "&txtFLG=" & m_iFLG 
    m_sArg_top = m_sArg_top & "&txtSNm=" & Server.URLEncode(m_sSNm) 

End Sub


Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

	If m_sMode = "" Then
	    w_FrameUrl = "default2.asp"
	Else
	    w_FrameUrl = "main.asp" & m_sArg
	End If

%>
<html>
<head>
<title>�i�H���񌟍�</title>
</head>

<frameset rows=190,1,* frameborder="0">
	<frame src="top.asp<%=m_sArg_top%>" scrolling="auto" noresize name="top">
    <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
	<frame src="<%=w_FrameUrl%>" scrolling="auto" noresize name="main">
</frameset>

</html>
<%
    '---------- HTML END   ----------
End Sub
%>
