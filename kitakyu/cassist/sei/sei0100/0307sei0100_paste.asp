<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ѓo�^
' ��۸���ID : sei/sei0100/sei0100_paste.asp
' �@      �\: ���ѓ\��t���p
'-------------------------------------------------------------------------
' ��      ��:�w�������E�\��t���Ώ�(���сE�x���E����)
' ��      ��:�Ȃ�
' ��      ��:�N���b�v�{�[�h���琬�уf�[�^���擾����
'-------------------------------------------------------------------------
' ��      ��: 2002/02/04 ���� ���
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////

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
	w_sMsgTitle="���ѓo�^"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"     
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
            m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
            Exit Do
        End If

		'// �����`�F�b�N�Ɏg�p
		session("PRJ_No") = "SEI0100"

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))

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
	Dim w_sTitle

	Select Case request("PasteType")
	Case "Seiseki"
		w_sTitle = "���ѓ\��t��"
	Case "Chikai"
		w_sTitle = "�x���\��t��"
	Case "Kekka"
		w_sTitle = "���ۑΏۓ\��t��"
	Case "KekkaGai"
		w_sTitle = "���ۑΏۊO�\��t��"
	Case Else
		w_sTitle = ""
	End Select
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    '---------- HTML START ----------
    %>
<html>
<head>
<title><%=w_sTitle%></title>
<link rel=stylesheet href="../../common/style.css" type="text/css">
<script language=javascript>
<!--
        //************************************************************
        //  [�@�\]  �N���A�{�^�����N���b�N�����ꍇ
        //  [����]
        //  [�ߒl]
        //  [����]
        //************************************************************
        function f_Clear(p_No) {

            document.frm.paste.value = "";

            return true;    
        }

        //************************************************************
        //  [�@�\]  �\��t���{�^�����N���b�N�����ꍇ
        //  [����]
        //  [�ߒl]
        //  [����]
        //************************************************************
        function f_Paste() {
			var str
			var i;
			var textbox;
			var strLen;

			//�����̓`�F�b�N
			if (document.frm.paste.value=="") {
				alert("�\��t���Ώۃf�[�^������܂���B");
				return false;
			}

			//�\��t��������̎擾
			str = (document.frm.paste.value).split("\r");
			strLen = str.length;

			//�w�����ł̃��[�v
			for(i=1;i<=<%=request("i_Max")%>;i++) {

				//�e�E�B���h�E�ɑ��݂��邩�ǂ���
				textbox = eval("opener.parent.main.document.frm.<%=request("PasteType")%>" + i);

				//(�擾�ł����f�[�^���Ɋ֌W�Ȃ��S�f�[�^����U�N���A����)
				//if (textbox && i<=strLen + 1) {
				if (textbox){

					//������
					textbox.value = "";

					if (str[i-1] != "") {
						//�����łȂ��͖̂�������
						if (!isNaN(str[i-1])) {
							textbox.value = str[i-1];
						}
					}
				}
			}

			//���v�E���ς̌v�Z
			eval("opener.parent.main").f_GetTotalAvg();

			window.close();
        }
    //-->
    </script>

</head>

<body>
<form name="frm">
<center>
<%call gs_title(w_sTitle,"�o�@�^")%>

<br>

<table border="0" cellpadding="1" cellspacing="1">
	<tr>
		<td align="center" colspan="2">

			<span class="msg">��Excel�t�@�C������R�s�[�����f�[�^��<BR>�\��t���Ă��������B</span><br>

		</td>
	</tr>
	<tr>
		<td align="center" width="250" valign="top">
			<textarea name="paste" COLS="20" ROWS="27"></textarea>
		</td>
	</tr>
	<tr>
		<td align="center" colspan="2">
			<br>
		    <input type="button" value=" �\��t�� " class="button" onclick="javascript:f_Paste('<%=m_iI%>');">�@
		    <input type="button" value=" �N���A " class="button" onclick="javascript:f_Clear('<%=m_iI%>');">�@
		    <input type="button" value="����" class="button" onclick="javascript:window.close();">
		</td>
	</tr>
</table>

<INPUT TYPE="HIDDEN" NAME="GAKUNEN" VALUE="<%=request("m_sGakunen") %>">
<INPUT TYPE="HIDDEN" NAME="CLASS"   VALUE="<%=request("m_sClass") %>">
<INPUT TYPE="HIDDEN" NAME="IINNM"   VALUE="<%=request("m_sIinNm") %>">
<INPUT TYPE="HIDDEN" NAME="i"       VALUE="<%=request("m_iI") %>">

</center>
</form>
</center>
</body>
</html>

<%
    '---------- HTML END   ----------
End Sub
%>