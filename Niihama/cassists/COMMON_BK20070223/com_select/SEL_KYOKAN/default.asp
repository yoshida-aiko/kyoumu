<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �����Q�ƑI�����
' ��۸���ID : Common/com_select/SEL_KYOKAN/default.asp
' �@      �\: �t���[���y�[�W �����̎Q�ƁA�I�����s��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h		��		SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h		��		SESSION���i�ۗ��j
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/07/19 �O�c �q�j
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
	'�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    Public  m_bErrMsg           '�װү����
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
	w_sMsgTitle="�����Q�ƑI�����"
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
            m_bErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

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
Dim w_iI
Dim w_sKNM
Dim w_sGakkaCd
	w_iI	 = request("txtI")
	w_sKNM	 = request("txtKNm")
	w_sGakkaCd = request("txtGakka")
	sArg = ""
	sArg = sArg & "txtI=" & w_iI 
	sArg = sArg & "&txtKNm=" & Server.URLEncode(w_sKNM)
	sArg = sArg & "&txtGakka=" & Server.URLEncode(w_sGakkaCd)

%>
<html>

<head>

<title>�����Q�ƑI�����</title>

<frameset rows=175px,1,* frameborder="no">
	<frame src="SEL_KYOKAN_top.asp?<%=sArg %>" scrolling="auto" noresize name="top">
    <frame src="bar.html" scrolling="auto" noresize name="bar">
	<frame src="default2.asp" scrolling="auto" noresize name="main">
</frameset>

</head>

</html>
<%
End Sub
%>