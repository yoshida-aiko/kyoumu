<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����w�Z��񌟍�
' ��۸���ID : mst/mst0123/default.asp
' �@      �\: �t���[���y�[�W �����w�Z�}�X�^�̎Q�Ƃ��s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
'           :txtKenCd       :�s���{���R�[�h     '/2001/07/31�ǉ�
'           :txtSityoCd     :�s�����R�[�h       '/2001/07/31�ǉ�
'           :txtSyuName     :�����w�Z��           '/2001/07/31�ǉ�
'           :txtPageSyu     :�\���ϕ\���Ő�     '/2001/07/31�ǉ�
'           :txtMode        :���[�h             '/2001/07/31�ǉ�
'           :txtTyuKbn      :���w�Z�敪         '/2001/07/31�ǉ�
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'           :�����N�x       ��      SESSION���i�ۗ��j
' �@      �@:session("PRJ_No")      '���������̃L�[ '/2001/07/31�ǉ�
'           :txtKenCd       :�s���{���R�[�h     '/2001/07/31�ǉ�
'           :txtSityoCd     :�s�����R�[�h       '/2001/07/31�ǉ�
'           :txtSyuName     :�����w�Z��           '/2001/07/31�ǉ�
'           :txtPageSyu     :�\���ϕ\���Ő�     '/2001/07/31�ǉ�
'           :txtMode        :���[�h             '/2001/07/31�ǉ�
'           :txtSyuKbn      :�����w�Z�敪         '/2001/07/31�ǉ�
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/06/20 �≺�@�K��Y
' ��      �X: 2001/07/27 ���{�@����(DB�ύX�ɔ����C��)
'           : 2001/07/31 ���{ ����  �ϐ��������K���Ɋ�ύX
'           :                       �����E���n�ǉ�
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�

    Public  m_sKenCd            '�s���{���R�[�h
    Public  m_sSityoCd          '�s�����R�[�h
    'Public  m_sSinroName       
    Public  m_sPageSyu          '�\���ϕ\���Ő�
    'Public  m_sSentakuSinroCD  
    Public  m_sMode             '���[�h
    Public  m_sSyuName          '�����w�ZM��

    Public  m_iSyuKbn       ':�����w�Z�敪

    Public  m_sArg          ':����'/2001/07/31�ύX
    Public  m_sArg_top      ':����'/2001/07/31�ύX

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
    w_sMsgTitle="�����w�Z��񌟍�"
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
        session("PRJ_No") = "MST0123"

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

    m_sKenCd = Request("txtKenCd")      ':�i�H�R�[�h
    m_sSityoCd = Request("txtSityoCd")  ':�i�w�R�[�h
    m_sSyuName = Request("txtSyuName")  ':�����w�Z��
    m_sPageSyu = Request("txtPageSyu")  ':�\���ϕ\���Ő�
    m_sMode = Request("txtMode")        ':���[�h
    
    m_iSyuKbn = Request("txtSyuKbn")        ':�����w�Z�敪

    m_sArg = "?"
    m_sArg = m_sArg & "txtMode=" & m_sMode 
    m_sArg = m_sArg & "&txtKenCd=" & m_sKenCd 
    m_sArg = m_sArg & "&txtSityoCd=" & m_sSityoCD 
    m_sArg = m_sArg & "&txtSyuName=" & m_sSyuName 
    m_sArg = m_sArg & "&txtPageSyu=" & m_sPageSyu
    m_sArg = m_sArg & "&txtSyuKbn=" & m_iSyuKbn

    m_sArg_top = "?"
    m_sArg_top = m_sArg_top & "txtMode=" & m_sMode 
    m_sArg_top = m_sArg_top & "&txtKenCd=" & m_sKenCd 
    m_sArg_top = m_sArg_top & "&txtSityoCd=" & m_sSityoCD 
    m_sArg_top = m_sArg_top & "&txtPageSyu=" & m_sPageSyu 
    m_sArg_top = m_sArg_top & "&txtSyuName=" & m_sSyuName 
    m_sArg_top = m_sArg_top & "&txtSyuKbn=" & m_iSyuKbn

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

<title>�����w�Z��񌟍�</title>

<frameset rows=180,1,* frameborder="0">

<frame src="./top.asp<%= m_sArg_top %>" scrolling="auto" noresize name="top">
<frame src="../../common/bar.html" scrolling="auto" noresize name="bar">

<frame src="
<%If m_sMode = "" Then%>
    default2.asp
<%Else%>
    main.asp<%= m_sArg %>
<%End If%>
" scrolling="auto" noresize name="main">
</frameset>


</head>

</html>
<%
    '---------- HTML END   ----------
End Sub
%>