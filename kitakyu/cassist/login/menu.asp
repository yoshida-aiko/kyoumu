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
    w_sMsgTitle="�w�b�_�[�f�[�^"
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
    <title>�w�Ѓf�[�^����</title>
    <meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
	<link rel=stylesheet href="../common/style.css" type=text/css>
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
    //************************************************************
    //  [�@�\]  �����[�h���ă��j���[�̕\����������
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function fj_CaseMenu(pMode){

		document.frm.hidMenuMode.value = pMode;
		document.frm.action="menu.asp";
		document.frm.target="menu";
		document.frm.submit();
		
    }
    //-->
    </SCRIPT>

    </head>

    <body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="images/back.gif">
	<form name="frm" method="post">

    <table border="0" cellspacing="0" cellpadding="0" width="150" height="100%">
        <tr>
            <td align="center" valign="top">

                <table border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <td class=home>

                            <table bordercolor="#222268" border="1" cellspacing="0" cellpadding="0" width="140">
                                <tr>
					<% if m_MenuMode = "" then %>
                                    <td class=home align="center"><font color="#ffff00">�s�@�n�@�o</font></td>
					<% Else %>
                                    <td class=home align="center"><font color="#ffffff"><a class=menu href="top.asp" target="<%=C_MAIN_FRAME%>" onClick="javascript:fj_CaseMenu('');">�s�@�n�@�o</a></font></td>
					<% End if %>
                                </tr>
                            </table>

                        </td>
                    </tr>
                    <tr><td><img src="../image/sp.gif"></td></tr>
<!--
					<% if m_MenuMode = "REGIST" then %>
							<tr><td class=category><font color="#ffff00">�e����̓t�H�[��<img src="images/sankaku_dow.gif" border="0"></font></td></tr>
					<% Else %>
							<tr><td class=category><a href="javascript:fj_CaseMenu('REGIST');"><font color="#ffffff">�e����̓t�H�[��<img src="images/sankaku.gif" border="0"></font></a></td></tr>
					<% End if %>
					<% if m_MenuMode = "REGIST" then Call s_MenuDateRegist() %>
					<tr><td><img src="../image/sp.gif"></td></tr>

					<% if m_MenuMode = "REFER" then %>
	                    <tr><td class=category><font color="#ffff00">�e�팟��<img src="images/sankaku_dow.gif" border="0"></font></td></tr>
					<% Else %>
	                    <tr><td class=category><a href="javascript:fj_CaseMenu('REFER');"><font color="#ffffff">�e�팟��<img src="images/sankaku.gif" border="0"></font></a></td></tr>
					<% End if %>
					<% if m_MenuMode = "REFER" then Call s_MenuDateRefer() %>
					<tr><td><img src="../image/sp.gif"></td></tr>

					<% if m_MenuMode = "ETC" then %>
	                    <tr><td class=category><font color="#ffff00">���̑�<img src="images/sankaku_dow.gif" border="0"></font></a></td></tr>
					<% Else %>
	                    <tr><td class=category><a href="javascript:fj_CaseMenu('ETC');"><font color="#ffffff">���̑�<img src="images/sankaku.gif" border="0"></font></a></td></tr>
					<% End if %>
					<% if m_MenuMode = "ETC" then Call s_MenuDateETC() %>
					<tr><td><img src="../image/sp.gif"></td></tr>
//-->
					<% if m_MenuMode = "SYUKETU" then %>
							<tr><td class=category><font color="#ffff00">�o������<img src="images/sankaku_dow.gif" border="0"></font></td></tr>
					<% Else %>
							<tr><td class=category><a href="javascript:fj_CaseMenu('SYUKETU');"><font color="#ffffff">�o������<img src="images/sankaku.gif" border="0"></font></a></td></tr>
					<% End if %>
					<% if m_MenuMode = "SYUKETU" then Call s_MenuData("SYUKETU") %>
					<tr><td><img src="../image/sp.gif"></td></tr>

					<% if m_MenuMode = "SHIKEN" then %>
							<tr><td class=category><font color="#ffff00">�����E����<img src="images/sankaku_dow.gif" border="0"></font></td></tr>
					<% Else %>
							<tr><td class=category><a href="javascript:fj_CaseMenu('SHIKEN');"><font color="#ffffff">�����E����<img src="images/sankaku.gif" border="0"></font></a></td></tr>
					<% End if %>
					<% if m_MenuMode = "SHIKEN" then Call s_MenuData("SHIKEN") %>
					<tr><td><img src="../image/sp.gif"></td></tr>

					<% if m_MenuMode = "SCHE" then %>
							<tr><td class=category><font color="#ffff00">�X�P�W���[��<img src="images/sankaku_dow.gif" border="0"></font></td></tr>
					<% Else %>
							<tr><td class=category><a href="javascript:fj_CaseMenu('SCHE');"><font color="#ffffff">�X�P�W���[��<img src="images/sankaku.gif" border="0"></font></a></td></tr>
					<% End if %>
					<% if m_MenuMode = "SCHE" then Call s_MenuData("SCHE") %>
					<tr><td><img src="../image/sp.gif"></td></tr>

					<% if m_MenuMode = "OTHERS" then %>
							<tr><td class=category><font color="#ffff00">���̑�����<img src="images/sankaku_dow.gif" border="0"></font></td></tr>
					<% Else %>
							<tr><td class=category><a href="javascript:fj_CaseMenu('OTHERS');"><font color="#ffffff">���̑�����<img src="images/sankaku.gif" border="0"></font></a></td></tr>
					<% End if %>
					<% if m_MenuMode = "OTHERS" then Call s_MenuData("OTHERS") %>
					<tr><td><img src="../image/sp.gif"></td></tr>

					<% if m_MenuMode = "INFO" then %>
							<tr><td class=category><font color="#ffff00">��񌟍�<img src="images/sankaku_dow.gif" border="0"></font></td></tr>
					<% Else %>
							<tr><td class=category><a href="javascript:fj_CaseMenu('INFO');"><font color="#ffffff">��񌟍�<img src="images/sankaku.gif" border="0"></font></a></td></tr>
					<% End if %>
					<% if m_MenuMode = "INFO" then Call s_MenuData("INFO") %>
					<tr><td><img src="../image/sp.gif"></td></tr>

					<% if m_MenuMode = "SUPPORT" then %>
							<tr><td class=category><font color="#ffff00">�x���@�\<img src="images/sankaku_dow.gif" border="0"></font></td></tr>
					<% Else %>
							<tr><td class=category><a href="javascript:fj_CaseMenu('SUPPORT');"><font color="#ffffff">�x���@�\<img src="images/sankaku.gif" border="0"></font></a></td></tr>
					<% End if %>
					<% if m_MenuMode = "SUPPORT" then Call s_MenuData("SUPPORT") %>
					<tr><td><img src="../image/sp.gif"></td></tr>

<!--
                    <tr><td class=info align="center"><font color="#ffffff"><a class=menu href="http://www.infogram.co.jp/" target="_blank"><img src="images/logo.gif" border="0"></a></font></td></tr>
//-->
                </table>

            </td>
        </tr>
    </table>

	<input type="hidden" name="hidMenuMode">
	</form>
    </body>

    </html>
<% End Sub



'********************************************************************************
'*  [�@�\]  �f�[�^�o�^���j���[
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]   
'********************************************************************************
Sub s_MenuData(p_menu) 
Select Case p_menu
	Case "SYUKETU" %>
		<% if gf_empMenu("KKS0110") then %><tr><td><a class=menu href="../kks/kks0110/" target="<%=C_MAIN_FRAME%>">���Əo������</a></td></tr><% End if %>
		<% if gf_empMenu("KKS0140") then %><tr><td><a class=menu href="../kks/kks0140/" target="<%=C_MAIN_FRAME%>">�s���o������</a></td></tr><% End if %>
		<% if gf_empMenu("KKS0170") then %><tr><td><a class=menu href="../kks/kks0170/" target="<%=C_MAIN_FRAME%>">�����o������</a></td></tr><% End if %>
	<% Case "SHIKEN" %>
		<% if gf_empMenu("SKN0130") then %><tr><td><a class=menu href="../skn/skn0130/" target="<%=C_MAIN_FRAME%>">�������{�Ȗړo�^</a></td></tr><% End if %>
		<% if gf_empMenu("SKN0120") then %><tr><td><a class=menu href="../skn/skn0120/" target="<%=C_MAIN_FRAME%>">�����ēƏ��\���o�^</a></td></tr><% End if %>
		<% if gf_empMenu("SEI0100") then %><tr><td><a class=menu href="../sei/sei0100/" target="<%=C_MAIN_FRAME%>">���ѓo�^</a></td></tr><% End if %>
		<% if gf_empMenu("SEI0500") then %><tr><td><a class=menu href="../sei/sei0500/" target="<%=C_MAIN_FRAME%>">���͎������ѓo�^</a></td></tr><% End if %>
		<% if gf_empMenu("SEI0500") then %><tr><td><a class=menu href="../sei/sei0800/" target="<%=C_MAIN_FRAME%>">�Ď������ѓo�^</a></td></tr><% End if %>
		<% if gf_empMenu("SEI0400") then %><tr><td><a class=menu href="../sei/sei0400/" target="<%=C_MAIN_FRAME%>">�����������o�^</a></td></tr><% End if %>
		<% if gf_empMenu("SEI0600") then %><tr><td><a class=menu href="../sei/sei0600/" target="<%=C_MAIN_FRAME%>">���ȓ����o�^</a></td></tr><% End if %>
		<% if gf_empMenu("SKN0170") then %><tr><td><a class=menu href="../skn/skn0170/" target="<%=C_MAIN_FRAME%>">�������Ԋ�(�N���X��)</a></td></tr><% End if %>
		<% if gf_empMenu("SKN0180") then %><tr><td><a class=menu href="../skn/skn0180/" target="<%=C_MAIN_FRAME%>">�������ԋ����\��ꗗ</a></td></tr><% End if %>
		<% if gf_empMenu("SEI0200") then %><tr><td><a class=menu href="../sei/sei0700/default.asp?p_mode=P_HAN0100" target="<%=C_MAIN_FRAME%>">���шꗗ</a></td></tr><% End if %>
		<% if gf_empMenu("SEI0200") then %><tr><td><a class=menu href="../sei/sei0700/default.asp?p_mode=P_KKS0200" target="<%=C_MAIN_FRAME%>">���ۈꗗ</a></td></tr><% End if %>
		<% if gf_empMenu("SEI0200") then %><tr><td><a class=menu href="../sei/sei0700/default.asp?p_mode=P_KKS0210" target="<%=C_MAIN_FRAME%>">�x���ꗗ</a></td></tr><% End if %>
		<% if gf_empMenu("SEI0200") then %><tr><td><a class=menu href="../sei/sei0700/default.asp?p_mode=P_HAN0111" target="<%=C_MAIN_FRAME%>">�]�_�ꗗ</a></td></tr><% End if %>
		<% if gf_empMenu("SEI0200") then %><tr><td><a class=menu href="../sei/sei0700/default.asp?p_mode=P_HAN0400_48" target="<%=C_MAIN_FRAME%>">���͎������шꗗ</a></td></tr><% End if %>
		<% if gf_empMenu("SEI0300") then %><tr><td><a class=menu href="../sei/sei0300/" target="<%=C_MAIN_FRAME%>">�l�ʐ��шꗗ</a></td></tr><% End if %>
		<% if gf_empMenu("HAN0121") then %><tr><td><a class=menu href="../han/han0121/" target="<%=C_MAIN_FRAME%>">���N�Y���҈ꗗ</a></td></tr><% End if %>
		
	<% Case "SCHE" %>
		<% if gf_empMenu("GYO0200") then %><tr><td><a class=menu href="../gyo/gyo0200/" target="<%=C_MAIN_FRAME%>">�s�������ꗗ</a></td></tr><% End if %>
		<% if gf_empMenu("JIK0210") then %><tr><td><a class=menu href="../jik/jik0210/" target="<%=C_MAIN_FRAME%>">�N���X�ʎ��Ǝ��Ԉꗗ</a></td></tr><% End if %>
		<% if gf_empMenu("JIK0200") then %><tr><td><a class=menu href="../jik/jik0200/" target="<%=C_MAIN_FRAME%>">�����ʎ��Ǝ��Ԉꗗ</a></td></tr><% End if %>
		<% if gf_empMenu("WEB0310") then %><tr><td><a class=menu href="../web/web0310/" target="<%=C_MAIN_FRAME%>">���Ԋ������A��</a></td></tr><% End if %>

	<% Case "OTHERS" %>
		<% if gf_empMenu("MST0144") then %><tr><td><a class=menu href="../mst/mst0144/" target="<%=C_MAIN_FRAME%>">�i�H����o�^</a></td></tr><% End if %>
		<% if gf_empMenu("WEB0320") then %><tr><td><a class=menu href="../web/web0320/" target="<%=C_MAIN_FRAME%>">�g�p���ȏ��o�^</a></td></tr><% End if %>
		<% if gf_empMenu("GAK0460") then %><tr><td><a class=menu href="../gak/gak0460/" target="<%=C_MAIN_FRAME%>">�w���v�^�������o�^</a></td></tr><% End if %>
		<% if gf_empMenu("GAK0461") then %><tr><td><a class=menu href="../gak/gak0461/" target="<%=C_MAIN_FRAME%>">�������������o�^</a></td></tr><% End if %>
		<% if gf_empMenu("GAK0470") then %><tr><td><a class=menu href="../gak/gak0470/" target="<%=C_MAIN_FRAME%>">�e��ψ��o�^</a></td></tr><% End if %>
		<% if gf_empMenu("WEB0340") then %><tr><td><a class=menu href="../web/web0340/" target="<%=C_MAIN_FRAME%>">�l���C�I���Ȗڌ���</a></td></tr><% End if %>
		<% if gf_empMenu("WEB0390") then %><tr><td><a class=menu href="../web/web0390/" target="<%=C_MAIN_FRAME%>">���x���ʉȖڌ���</a></td></tr><% End if %>
		<% if gf_empMenu("WEB0360") then %><tr><td><a class=menu href="../web/web0360/" target="<%=C_MAIN_FRAME%>">�����������ꗗ</a></td></tr><% End if %>

	<% Case "INFO" %>
		<% if gf_empMenu("GAK0300") then %><tr><td><a class=menu href="../gak/gak0310/" target="<%=C_MAIN_FRAME%>">�w����񌟍�</a></td></tr><% End if %>
		<% if gf_empMenu("MST0113") then %><tr><td><a class=menu href="../mst/mst0113/" target="<%=C_MAIN_FRAME%>">���w�Z��񌟍�</a></td></tr><% End if %>
		<% if gf_empMenu("MST0123") then %><tr><td><a class=menu href="../mst/mst0123/" target="<%=C_MAIN_FRAME%>">�����w�Z��񌟍�</a></td></tr><% End if %>
		<% if gf_empMenu("MST0133") then %><tr><td><a class=menu href="../mst/mst0133/" target="<%=C_MAIN_FRAME%>">�i�H���񌟍�</a></td></tr><% End if %>
		<% if gf_empMenu("WEB0350") then %><tr><td><a class=menu href="../web/web0350/" target="<%=C_MAIN_FRAME%>">�󂫎��ԏ�񌟍�</a></td></tr><% End if %>

	<% Case "SUPPORT" %>
		<% if gf_empMenu("WEB0300") then %><tr><td><a class=menu href="../web/web0300/" target="<%=C_MAIN_FRAME%>">���ʋ����\��</a></td></tr><% End if %>
		<% if gf_empMenu("WEB0330") then %><tr><td><a class=menu href="../web/web0330/" target="<%=C_MAIN_FRAME%>">�A�������o�^</a></td></tr><% End if %>
		<% if gf_empMenu("WEB0330") then %><tr><td><a class=menu href="../login/top.asp" target="<%=C_MAIN_FRAME%>">�A���f����</a></td></tr><% End if %>

<% End Select %>
		<tr><td> </td></tr>
<%
End Sub
%>