<%@ Language=VBScript %>
<%
'*************************************************************************
'* �V�X�e����: ���������V�X�e��
'* ��  ��  ��: ���N�Y���҈ꗗ
'* ��۸���ID : han/han0121/header.asp
'* �@      �\: ��y�[�W ���N�Y���҈ꗗ�̌������s��
'*-------------------------------------------------------------------------
'* ��      ��:�����R�[�h     ��      SESSION���i�ۗ��j
'*           :�����N�x       ��      SESSION���i�ۗ��j
'*           :session("PRJ_No")      '���������̃L�[
'* ��      ��:�Ȃ�
'* ��      �n:�����R�[�h     ��      SESSION���i�ۗ��j
'*           :�����N�x       ��      SESSION���i�ۗ��j
'*           cboGakunenCd      :�w�N�R�[�h
'* ��      ��:
'*           �������\��
'*               �R���{�{�b�N�X�͊w�N��\��
'*           ���\���{�^���N���b�N��
'*               ���̃t���[���Ɏw�肵�������̗��N�Y���҈ꗗ��\��������
'*-------------------------------------------------------------------------
'* ��      ��: 2001/08/08 �O�c�@�q�j
'* ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->

<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    '�G���[�n
    Public  m_bErrFlg           '�װ�׸�
    
    '�I��p��Where����
    Public m_sGakunenWhere      '�w�N�̏���
    
    '�擾�����f�[�^�����ϐ�
    Public  m_iNendo         ':�����N�x
    Public  m_iKyokanCd         ':�����R�[�h
    
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
    w_sWinTitle= "�L�����p�X�A�V�X�g"
    w_sMsgTitle= "���N�Y���҈ꗗ"
    w_sMsg= ""
    w_sRetURL= C_RetURL & C_ERR_RETURL
    w_sTarget= ""


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

        '// �s���A�N�Z�X�`�F�b�N
        Call gf_userChk(session("PRJ_No"))

        '// ���Ұ�SET
        Call s_SetParam()

        '�w�N�R���{�Ɋւ���WHERE���쐬����
        Call s_MakeGakunenWhere() 

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

Sub s_SetParam()
'********************************************************************************
'*  [�@�\]  �����n����Ă����l��ݒ�
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

    m_iNendo = Session("NENDO")
    m_iKyokanCd = Session("KYOKAN_CD")

End Sub

Sub s_MakeGakunenWhere()
'********************************************************************************
'*  [�@�\]  �w�N�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
    
    m_sGakunenWhere = ""
    
    m_sGakunenWhere = m_sGakunenWhere & " M05_NENDO = " & m_iNendo
    m_sGakunenWhere = m_sGakunenWhere & " GROUP BY M05_GAKUNEN"
    
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
<link rel=stylesheet href="../../common/style.css" type=text/css>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    //************************************************************
    //  [�@�\]  �y�[�W���[�h������
    //  [����]
    //  [�ߒl]
    //  [����]
    //************************************************************
    function window_onload() {

    }
    //************************************************************
    //  [�@�\]  �\���{�^���N���b�N���̏���
    //  [����]  �Ȃ�
    //  [�ߒl]  �Ȃ�
    //  [����]
    //
    //************************************************************
    function f_Search(){

        if( f_Trim(document.frm.cboGakunenCd.value) == "<%=C_CBO_NULL%>" ){
            window.alert("�w�N�̑I�����s���Ă�������");
            document.frm.cboGakunenCd.focus();
            return ;
		}

        document.frm.action="ichiran.asp";
        document.frm.target="main";
        document.frm.submit();
    
    }

    //-->
    </SCRIPT>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<center>
<form name="frm" method="post">
<table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
	<tr>
		<td valign="top" align="center">
    	<%call gs_title("���N�Y���҈ꗗ","��@��")%>
<br>
    		<table border="0">
    			<tr>
    				<td class=search>
				        <table border="0" cellpadding="1" cellspacing="1">
					        <tr>
						        <td align="left">
						            <table border="0" cellpadding="1" cellspacing="1">
							            <tr>
								            <td align="left">�w�N</td>
								            <td align="left">
								            <% call gf_ComboSet("cboGakunenCd",C_CBO_M05_CLASS_G,m_sGakunenWhere," style='width:40px;' ",True,"") %>
								            �N</td>
										    <td valign="bottom">
									        <input type="button" value="�@�\�@���@" onClick = "javascript:f_Search()" class=button>
										    </td>
							            </tr>
						            </table>
						        </td>
					        </tr>
				        </table>
				    </td>
			    </tr>
		    </table>
		</td>
	</tr>
</table>

	<input type=hidden name=txtMode value="Hyouji">
	<input type=hidden name=txtNendo value="<%=m_iNendo%>">
	<input type=hidden name=txtKyokanCd value="<%=m_iKyokanCd%>">

</form>
</center>
</body>
</html>
<%
    '---------- HTML END   ----------
End Sub
%>
