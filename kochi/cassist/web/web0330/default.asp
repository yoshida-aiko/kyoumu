<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �A���f����
' ��۸���ID : web/web0330/default.asp
' �@      �\: �t���[���y�[�W �A���f�����s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/07/10 �O�c
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_stxtMode
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
    w_sMsgTitle="�A���f����"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    m_stxtMode = request("txtMode")

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
		session("PRJ_No") = "WEB0330"

		'session("NENDO") = 2001

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

        If m_stxtMode = "NEW" or m_sTxtMode = "UPD" Then
            '// �y�[�W��\��
            Call TOUROKU_showpage()
            Exit Do
        ElseIf m_stxtMode = "" Then
            '// �y�[�W��\��
            Call showPage()
            Exit Do
        End If

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

<title>�A���f����</title>

<frameset rows="110,1,*" frameborder="no">
    <frame src="web0330_top.asp" scrolling="auto" noresize name="top">
    <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
    <frame src="web0330_main.asp" scrolling="auto" noresize name="main">
</frameset>

</head>

</html>
<%
End Sub

Sub TOUROKU_showpage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Dim w_sKenmei
Dim w_sNaiyou
Dim w_sKaisibi
Dim w_sSyuryoubi
Dim w_stxtMode
Dim w_sNendo
Dim w_sKyokanCd
Dim w_stxtNo

    w_sKenmei    = request("Kenmei")
    w_sNaiyou    = request("Naiyou")
    w_sKaisibi   = request("Kaisibi")
    w_sSyuryoubi = request("Syuryoubi")
    w_stxtMode   = request("txtMode")
    w_sNendo     = request("txtNendo")
    w_sKyokanCd  = request("txtKyokanCd")
    w_stxtNo     = request("txtNo")

        sArg = ""
        sArg = sArg & "txtKenmei=" & Server.URLEncode(w_sKenmei)
        sArg = sArg & "&txtNaiyou=" & Server.URLEncode(w_sNaiyou)
        sArg = sArg & "&txtKaisibi=" & w_sKaisibi 
        sArg = sArg & "&txtSyuryoubi=" & w_sSyuryoubi 
        sArg = sArg & "&txtMode=" & w_stxtMode 
        sArg = sArg & "&txtNendo=" & w_sNendo 
        sArg = sArg & "&txtKyokanCd=" & w_sKyokanCd 
        sArg = sArg & "&txtNo=" & w_stxtNo 

%>
<html>

<head>

<title>�A���f����</title>

<frameset rows=250,1,* frameborder="no">
    <frame src="sousin_top.asp?<%=sArg %>" scrolling="auto" noresize name="top">
    <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
    <frame src="default2.asp?<%=sArg %>" scrolling="auto" noresize name="main">
</frameset>

</head>

</html>
<%
End Sub
%>