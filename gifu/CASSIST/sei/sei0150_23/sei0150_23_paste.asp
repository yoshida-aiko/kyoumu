<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���ѓo�^
' ��۸���ID : sei/sei0100/sei0150_paste.asp
' �@      �\: ���ѓ\��t���p
'-------------------------------------------------------------------------
' ��      ��:�w�������E�\��t���Ώ�(���сE�x���E����)
' ��      ��:�Ȃ�
' ��      ��:�N���b�v�{�[�h���琬�уf�[�^���擾����
'-------------------------------------------------------------------------
' ��      ��: 2002/02/04 ���� ���
' 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
Dim m_iSeisekiInpType
Dim m_KekkaGaiDispFlg	'//���ۊO�\���t���O
Dim m_bKekkaNyuryokuFlg	'//���ۓ��͊��ԃt���O

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
		
		m_iSeisekiInpType = cint(request("hidSeisekiInpType"))
		m_KekkaGaiDispFlg = request("hidKekkaGaiDispFlg")
		m_bKekkaNyuryokuFlg	= request("hidKekkaNyuryokuFlg")
		
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

    '---------- HTML START ----------
    %>
<html>
<head>
<title>���ѓ\��t���]��</title>
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
			var str;
			var i;
			var textbox;
			var strLen;
			var strMaxLen = "3";
			var w_sObj = "<%=request("PasteType")%>";

			if(w_sObj == "txtKekka"){
				strMaxLen = "3"
			}
			if(w_sObj == "txtKibi"){
				strMaxLen = "2"
			}
			if(w_sObj == "txtTeisi"){
				strMaxLen = "3"
			}
			if(w_sObj == "txtHaken"){
				strMaxLen = "2"
			}

			//�����̓`�F�b�N
			if (document.frm.paste.value=="") {
				alert("�]���Ώۃf�[�^������܂���B");
				return false;
			}

			//�\��t��������̎擾
			str = (document.frm.paste.value).split("\n");
			strLen = str.length;
			
			//�w�����ł̃��[�v
			for(i=1;i<=<%=request("i_Max")-1%>;i++) {
				//�e�E�B���h�E�ɑ��݂��邩�ǂ���
				textbox = eval("opener.parent.main.document.frm.<%=request("PasteType")%>" + i);
				
				//(�擾�ł����f�[�^���Ɋ֌W�Ȃ��S�f�[�^����U�N���A����)
				//if (textbox && i<=strLen + 1) {
				if (textbox){
					
					//÷���ޯ����ۯ����������ĂȂ�������
					if(textbox.readOnly == false && textbox.disabled == false){
						
						//������
						textbox.value = "";
						
						if (str[i-1] != "") {
							//�����łȂ��͖̂�������
							if (!isNaN(str[i-1])) {
								textbox.value = jf_Left(str[i-1],strMaxLen);
							}
						}
					}
				}
			}
			
			<% if m_iSeisekiInpType = cint(C_SEISEKI_INP_TYPE_NUM) then %>
				//���v�E���ς̌v�Z
				eval("opener.parent.main").f_GetTotalAvg();
			<% end if %>
			
			<% if m_bKekkaNyuryokuFlg then %>
				//�x���̍��v
//				eval("opener.parent.main").f_CalcSum("Chikai");
				
				//���ۂ̍��v
//				eval("opener.parent.main").f_CalcSum("Kekka");
				
				<% if m_KekkaGaiDispFlg then %>
					//���ۊO�̍��v
//					eval("opener.parent.main").f_CalcSum("KekkaGai");
				<% end if %>
			<% end if %>
			
			window.close();
        }

	//**************************************************************************************
	//////////////////////////   ������w�萔�̕����𔲂����   ////////////////////////////
	//--------------------------------------------------------------------------------------
	// Arguments: String length
	// Return:  String
	//**************************************************************************************
	function jf_Left(str,len){
	 if(str==null) return "";
	 if(len==null) return "";
	 if(isNaN(len)) return "";
	 str = str.substr(0,len);
	 return str;
	}

    //-->
    </script>

</head>

<body>
<form name="frm">
<center>
<%call gs_title("���ѓ\��t���]��","�o�@�^")%>

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
		    <input type="button" value=" �]�@�� " class="button" onclick="javascript:f_Paste();">�@
		    <input type="button" value=" �N���A " class="button" onclick="javascript:f_Clear();">�@
		    <input type="button" value="����" class="button" onclick="javascript:window.close();">
		</td>
	</tr>
</table>

</center>
</form>
</center>
</body>
</html>

<%
    '---------- HTML END   ----------
End Sub
%>