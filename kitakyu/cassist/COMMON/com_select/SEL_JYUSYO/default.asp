<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �s�����������
' ��۸���ID : Common/com_select/SEL_JYUSYO/default.asp
' �@      �\: �t���[����`����
'-------------------------------------------------------------------------
' ��      ��:	
' 	           	JUSYO1	= ���s��
'   	        JUSYO2	= ��
' 
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/30 ���i
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../com_All.asp"-->
<%

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////

	Public  m_bErrFlg           '�װ�׸�
	Public  m_JUSYO1			'�Z��1
	Public  m_JUSYO2			'�Z��2

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

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�A�������o�^"
    w_sMsg=""
    w_sRetURL="../../login/top.asp"
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
            Call gs_SetErrMsg("�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B")
            Exit Do
        End If

		'// ���Ұ��擾
		m_JUSYO1 = request("txtJUSYO1")
		m_JUSYO2 = request("txtJUSYO2")

		'// ����݊i�[
'		Session("m_JUSYO1") = m_JUSYO1
'		Session("m_JUSYO2") = m_JUSYO2

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
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()

    On Error Resume Next
    Err.Clear

	m_JUSYO1 = Server.URLEncode(m_JUSYO1)
	m_JUSYO2 = Server.URLEncode(m_JUSYO2)

%>

<html>

<head>
<title>�s��������</title>
<link rel=stylesheet href="../../style.css" type=text/css>
</head>

<frameset rows=230,1,* frameborder="no" onload="window.focus();">
	<frame src="Jyusyo_top.asp?JUSYO1=<%=m_JUSYO1%>&JUSYO2=<%=m_JUSYO2%>" scrolling="auto" noresize name="top">
        <frame src="bar.html" scrolling="auto" noresize name="bar">
	<frame src="Jyusyo_dow.asp" scrolling="auto" noresize name="dow">
</frameset>

</html>
<%
End Sub
%>