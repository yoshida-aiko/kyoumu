<%@ Language=VBScript %>
<%Response.Expires = 0%>
<%Response.AddHeader "Pragma", "No-Cache"%>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �p�X���[�h�ύX
' ��۸���ID : web/web0400/default.asp
' �@      �\: ���O�C���p�X���[�h��ύX���܂��B
'-------------------------------------------------------------------------
' ��      ��:SESSION(""):�����R�[�h     ��      SESSION���
' ��      ��:�Ȃ�
' ��      �n:SESSION(""):�����R�[�h     ��      SESSION���
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2001/10/04 �J�e
' ��      �X: 2019/03/18 ���� �p�X���[�h�̃G���[�`�F�b�N�𔼊p�p���L���`�F�b�N�ɕύX
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public m_bErrFlg           '�װ�׸�
    Public m_iNendo				'�����N�x
    Public m_sUser				'���O�C�����[�U�h�c
    Public m_sPass				'�Â��p�X���[�h
    Public m_sPassN1			'�V�����p�X���[�h�P
    Public m_sPassN2			'�V�����p�X���[�h�Q
	Public m_loginF 			'���O�C��ID���͂����邩�ǂ���

'///////////////////////////���C������/////////////////////////////

    'Ҳ�ٰ�ݎ��s
    Call Main()
response.end
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
    w_sMsgTitle="�p�X���[�h�ύX"
    w_sMsg=""
    w_sRetURL="../../login/default.asp"
    w_sTarget="_top"

    On Error Resume Next
    Err.Clear

    m_bErrFlg = False

    Do

		'// �ϐ�������
		call f_paraSet()

			'// �����`�F�b�N�Ɏg�p
	'		session("PRJ_No") = "WEB0400"

			'// �s���A�N�Z�X�`�F�b�N
	'		Call gf_userChk(session("PRJ_No"))
			
        '// �ύX�y�[�W��\��
        Call showPage()
        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If m_bErrFlg = True Then
        w_sMsg = gf_GetErrMsg()
'response.write w_sMsg
'        Call gs_showMsgPage(w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget)
    End If
    
End Sub

Sub f_paraSet()
'*******************************************************************************
' �@�@�@�\�F�ϐ��̏������Ƒ��
' ���@�@���F�Ȃ�
' �@�\�ڍׁF
' ���@�@�l�F�Ȃ�
' ��@�@���F2001/08/29�@�J�e
'*******************************************************************************
m_sUser = Request("txtUser")
m_sPass = Request("txtPass")
m_sPassN1 = Request("txtPassN1")
m_sPassN2 = Request("txtPassN2")

if m_sUser = "" then m_sUser = session("LOGIN_ID")
'm_iNendo = 2001

m_loginF = true '���O�C���h�c����͂�����Ƃ���true
	
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
    <title>�p�X���[�h�ύX</title>
    <link rel=stylesheet href=../../common/style.css type=text/css>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
	<!--
		function jf_check(){
		var w_str //��Ɨp
		var w_nm //��Ɨp
		var w_msg //��Ɨp
		var w_err //��Ɨp
		
		w_err = true;
		while(1) {
			//���O�C��ID�̓��̓`�F�b�N
			w_nm = eval(document.frm.txtUser);
			w_str = w_nm.value;
			if (w_str == "") {w_str = "���O�C��ID����͂��Ă��������B";break;}
			if (!IsHankakuEisu(w_str)) {w_str = "���O�C��ID���A���p�p�������ł͂���܂���B";break;}
			if (getLengthB(w_str) > 16 ) {w_str = "���O�C��ID���A16�����ȉ��ł͂���܂���B";break;}

			//�p�X���[�h�̓��̓`�F�b�N
			w_nm = eval(document.frm.txtPass);
			w_str = w_nm.value;
			if (w_str == "") {w_str = "�p�X���[�h����͂��Ă��������B";break;}
			//if (!IsHankakuEisu(w_str)) {w_str = "�p�X���[�h���A���p�p�������ł͂���܂���B";break;}
			if (getLengthB(w_str) > 10 ) {w_str = "�p�X���[�h���A10�����ȉ��ł͂���܂���B";break;}

			//�p�X���[�h�̓��̓`�F�b�N
			w_nm = eval(document.frm.txtPassN1);
			w_str = w_nm.value;
			if (w_str == "") {w_str = "�V�����p�X���[�h(1)����͂��Ă��������B";break;}
			if (!IsHankakuEisuKigo(w_str)) {w_str = "�V�����p�X���[�h(1)���A���p�p���L���i�V���O���R�[�e�[�V�����������j�����ł͂���܂���B";break;}
			if (getLengthB(w_str) > 10 ) {w_str = "�V�����p�X���[�h(1)���A10�����ȉ��ł͂���܂���B";break;}

			//�p�X���[�h�̓��̓`�F�b�N
			w_nm = eval(document.frm.txtPassN2);
			w_str = w_nm.value;
			if (w_str == "") {w_str = "�V�����p�X���[�h(2)����͂��Ă��������B";break;}
			if (!IsHankakuEisuKigo(w_str)) {w_str = "�V�����p�X���[�h(2)���A���p�p���L���i�V���O���R�[�e�[�V�����������j�����ł͂���܂���B";break;}
			if (getLengthB(w_str) > 10 ) {w_str = "�V�����p�X���[�h(2)���A10�����ȉ��ł͂���܂���B";break;}
			if (document.frm.txtPassN1.value != w_str) {w_str = "�V�����p�X���[�h(2)���A�V�����p�X���[�h(1)�Ɠ����ł͂���܂���B";break;}

		w_err = false;
		break;
		}

		//�G���[�L��
		if (w_err) {
			jf_errMsg(w_nm,w_str);
			return false;
		}
		
		//����I��
		return true;
//		return false; //�e�X�g�p

		}

		function jf_errMsg(p_nm,p_str) {
			alert(p_str);
			p_nm.select();
			return;
		}

	//-->
	</script>
</head>
<body>
<center>
    <%call gs_title("���O�C���p�X���[�h�ύX","�X�@�V")%>

<BR>
���O�C���p�X���[�h�̕ύX���s�����Ƃ��ł��܂��B<br>
�Z�L�����e�B�Ǘ��̂��߁A����I�Ƀp�X���[�h�̕ύX���s���悤�ɂ��Ă��������B<BR><BR>
	<table class="hyo" border="0" width="70%">
		<FORM action="web0400_upd.asp" name="frm" method="post" target="_self" onSubmit="return jf_check();">
			<tr><td colspan="2" height="30" class="detail"></td></tr>
	        <tr>
	            <td nowrap class="detail" align="right" width="45%">���O�C��ID�F</td>
<% If m_loginF = true then %>
	            <td nowrap class="detail" width="55%">�@<input type="text" name="txtUser" value="<%=m_sUser%>" maxlength="16" size="25"></td>
	        <tr>
	            <td nowrap class="detail" align="right"></td>
	            <td nowrap class="detail">�@<span class="CAUTION" style="text-align:left;">�����p�p��16�����ȓ�</span></td>
	        </tr>
<% else %>
	            <td nowrap class="detail" width="55%">�@<%=m_sUser%><input type="hidden" name="txtUser" value="<%=m_sUser%>" maxlength="16" size="25"></td></tr>
			<tr><td colspan="2" height="15" class="detail"></td>
<% End If %>
	        </tr>
			<tr><td colspan="2" height="15" class="detail"></td></tr>
	        <tr>
	            <td nowrap class="detail" align="right">�p�X���[�h�F</td>
	            <td nowrap class="detail">�@<input type="password" name="txtPass" value="<%=m_sPass%>" maxlength="10"></td>
	        </tr>
	        <tr>
	            <td nowrap class="detail" align="right"></td>
	            <td nowrap class="detail">�@<span class="CAUTION" style="text-align:left;">�����p�p��10�����ȓ�</span></td>
	        </tr>
			<tr><td colspan="2" height="15" class="detail"></td></tr>
	        <tr>
	            <td nowrap class="detail" align="right">�V�����p�X���[�h�F</td>
	            <td nowrap class="detail">�@<input type="password" name="txtPassN1" value="<%=m_sPassN1%>" maxlength="10"></td>
	        </tr>
	        <tr>
	            <td nowrap class="detail" align="right"></td>
	            <td nowrap class="detail">�@<span class="CAUTION" style="text-align:left;">�����p�p��10�����ȓ�</span></td>
	        </tr>
			<tr><td colspan="2" height="15" class="detail"></td></tr>
	        <tr>
	            <td nowrap class="detail" align="right">�V�����p�X���[�h�F</td>
	            <td nowrap class="detail">�@<input type="password" name="txtPassN2" value="<%=m_sPassN2%>" maxlength="10"></td>
	        </tr>
	        <tr>
	            <td nowrap class="detail" align="right"></td>
	            <td nowrap class="detail">�@<span class="CAUTION" style="text-align:left;">���m�F�̂��߂ɂ�����x���͂��Ă�������</span></td>
	        </tr>
			<tr><td colspan="2" height="30" class="detail"></td></tr>
			<tr><td colspan="2" height="30" class="detail" align="center">
				<input type="submit" name="submit" value=" �� �X " maxlength="10">
	            �@<input type="reset" name="can" value="��ݾ�" maxlength="10" onclick="history.back()">
	            </td>
	        </tr>
		</FORM>
	</table>
</center>
</body>
</head>
</html>
<%
End Sub
%>