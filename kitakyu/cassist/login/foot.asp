<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���O�C���I�������
' ��۸���ID : login/menu.asp
' �@      �\: ���O�C���I�����̃��j���[���
'-------------------------------------------------------------------------
' ��      ��    
'               
' ��      ��
' ��      �n
'           
'           
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/02 
' ��      �X: 2001/07/26    ���`�i�K
'*************************************************************************/
%>
<!--#include file="../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
Dim m_MenuMode		'//�ƭ�Ӱ��

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

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    'Message�p�̕ϐ��̏�����
    w_sWinTitle="�L�����p�X�A�V�X�g"
    w_sMsgTitle="�t�b�^�[�f�[�^"
    w_sMsg=""
    w_sRetURL="../default.asp"
    w_sTarget="_parent"

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
		session("PRJ_No") = C_LEVEL_NOCHK

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

		'//�ƭ�Ӱ��
		m_MenuMode = request("hidMenuMode")

        '//�����\��
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

    %>
    <html>
    <head>
    <title>�t�b�^�[</title>
    <meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
	<link rel=stylesheet href="../common/style.css" type=text/css>
	    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [�@�\]  �g�b�v�֖߂�B
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function fj_BackTop() {

		document.frm.action="menu.asp";
		document.frm.target="menu";
		document.frm.submit();
		
    }
    //-->
    </SCRIPT>
    </head>
    <body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="images/back.gif" style="background-repeat: repeat-y;">
<table border=0 cellpadding=0 cellspacing=0 width="100%">
		<form action="menu.asp" method="post" name="frm">
            <tr><td width="152" class="info" align="center" valign="center" nowrap><span color="#ffffff"><a class="menu" href="http://www.infogram.co.jp/" target="_blank"><img src="images/logo.gif" border="0"></a></span></td>
				<td>�@</td>
				<td width="125" align="right" nowrap><a href="../web/web0380/default.asp" target="<%=C_MAIN_FRAME%>" onClick="">�m�ٓ��󋵈ꗗ�n</a></td>
				<td width="120" align="right" nowrap><a href="../web/web0370/default.asp" target="<%=C_MAIN_FRAME%>" onClick="">�m�w�����ꗗ�n</a></td>
				<td width="125" align="right" nowrap><a href="top.asp" target="<%=C_MAIN_FRAME%>" onClick="javascript:fj_BackTop()">�m�g�b�v�֖߂�n</a></td>
				<td width="120" align="right" nowrap><a href="../default.asp" target="_top">�m���O�A�E�g�n</a></td>
			</tr>
		</form>
		</table></body>
</html>
<%
End Sub%>
