<%@Language=VBScript %>
<%
'******************************************************************
'�V�X�e����     �F���������V�X�e��
'���@���@��     �F�g�p���ȏ��o�^
'�v���O����ID   �Fweb/web0320/default.asp
'�@�@�@�@�\     �F�t���[���y�[�W �g�p���ȏ��o�^�̕\�����s��
'------------------------------------------------------------------
'���@�@�@��     �F
'�ρ@�@�@��     �F
'���@�@�@�n     �F
'���@�@�@��     �F
'------------------------------------------------------------------
'��@�@�@��     �F2001.08.01    �O�c�@�q�j
'�ρ@�@�@�X     �F
'
'******************************************************************
'*******************�@ASP���ʃ��W���[���錾�@**********************
%>
<!--#include file="../../common/com_All.asp"-->
<%
'******�@�� �W �� �[ �� �� ���@********

	Public m_iNendo
	Public m_iGakunen
	Public m_iClassNo
	Public m_iPage

'******�@���C�������@********

    'Ҳ�ٰ�ݎ��s
    Call Main()

'******�@�d�@�m�@�c�@********

'******************************************************************
'�@�@�@�\�F�{ASP��Ҳ�ٰ��
'�ԁ@�@�l�F�Ȃ�
'���@�@���F�Ȃ�
'�@�\�ڍׁF
'���@�@�l�F���ɂȂ�
'******************************************************************
Sub Main()
    Dim w_iRet              '// �߂�l
    Dim w_sSQL              '// SQL��
    Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	
    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�g�p���ȏ��o�^"
    w_sMsg=""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget=""
	
    On Error Resume Next
    Err.Clear
	
    m_bErrFlg = False
	
    Do
        '// �ް��ް��ڑ�
        If gf_OpenDatabase() <> 0 Then
            '�ް��ް��Ƃ̐ڑ��Ɏ��s
            m_bErrFlg = True
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If
		
		'// �����`�F�b�N�Ɏg�p
		session("PRJ_No") = "WEB0320"
		
		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
		
		'm_iNendo = Request("txtNendo")
		m_iNendo = Request("hidYear")
		
		If m_iNendo <> "" Then
	        '// �y�[�W��\��
	        Call showPage_Reload()
		Else
	        '// �y�[�W��\��
	        Call showPage()
		End If

        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
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
    <title>�g�p���ȏ��o�^</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <script language=javascript>
    </script>
    <frameset rows=140,1,* frameborder="0">
        <frame src="web0320_top.asp" scrolling="auto" noresize name="top">
        <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
        <frame src="default2.asp" scrolling="auto" noresize name="main">
    </frameset>
    </head>
</html>
<%
End Sub

Sub showPage_Reload()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

	sArg = ""
	sArg = sArg & "?txtNendo=" & m_iNendo
%>
<html>
    <head>
    <title>�g�p���ȏ��o�^</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <script language=javascript>
    <!--
    /*
    window.onload = init;
    
    function init(){
		alert("<%=request("hidYear")%>");
		alert("<%=request("hidGakunen")%>");
		alert("<%=request("hidGakka")%>");
	}
	*/
    //-->
    </script>
    <frameset rows=140,1,* frameborder="0">
        <frame src="web0320_top.asp?<%=Request.Form.Item%>" scrolling="auto" noresize name="top">
        <frame src="../../common/bar.html" scrolling="auto" noresize name="bar">
        
        <% if request("txtPageCD") <> "" then %>
        	<frame src="web0320_main.asp?<%=Request.Form.Item%>" scrolling="auto" noresize name="main">
        <% else %>
        	<frame src="default2.asp" scrolling="auto" noresize name="main">
        <% end if %>
    </frameset>
    </head>
</html>
<%
End Sub
%>