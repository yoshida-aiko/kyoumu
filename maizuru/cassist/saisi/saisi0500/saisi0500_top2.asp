<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��:
' ��۸���ID :
' �@      �\:
'-------------------------------------------------------------------------
' ��      ��:
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2003/02/24 hirota
'*************************************************************************/

%>
<!--#include file="../../Common/com_All.asp"-->
<%

'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////

	Public msURL
	Public m_bErrFlg
	Public m_sGakunenWhere		'//�w�N�R���{�Z�b�g����
	Public m_sClassWhere		'//�N���X�R���{�Z�b�g����
	Public m_sClassOption       '//�N���X�R���{�̃I�v�V����

	Public m_iGakunen			'//�w�N
	Public m_iClassNo			'//�N���XNO
	Public m_iSyoriNen			'//�N�x
	Public m_iKyokanCd			'//��������
	Public m_iGakka				'//�w��
	Public m_sClassNM			'//�N���X��

	Const C_FIRST_DISP_GAKUNEN = 1

'///////////////////////////���C������/////////////////////////////

	'Ҳ�ٰ�ݎ��s
	Call Main()

'///////////////////////////�@�d�m�c�@/////////////////////////////

'********************************************************************************
'*  [�@�\]  �{ASP��Ҳ�ٰ��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub Main()

	Dim w_sWinTitle, w_sMsgTitle, w_sMsg, w_sRetURL, w_sTarget
	Dim w_iRet

	On Error Resume Next
	Err.Clear

	'Message�p�̕ϐ��̏�����
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="�s���i�w���ꗗ"
	w_sMsg=""
	w_sRetURL = C_RetURL & C_ERR_RETURL
	w_sTarget = "fTopMain"

	m_bErrFlg = False

	Do
		'//�ް��ް��ڑ�
		w_iRet = gf_OpenDatabase()
		If w_iRet <> 0 Then
			'�ް��ް��Ƃ̐ڑ��Ɏ��s
			m_sErrMsg = "�f�[�^�x�[�X�Ƃ̐ڑ��Ɏ��s���܂����B"
			Exit Do
		End If

		'//�l�̏�����
        Call s_ClearParam()

		'//�p�����[�^�擾
		Call s_GetParameter()

		'//�w�N�R���{�Z�b�g���̏���
		Call s_MakeGakunenWhere()

		'//�N���X�R���{�Ɋւ���WHERE���쐬����
		Call s_MakeClassWhere()

		'//�y�[�W��\��
		Call showPage()

		m_bErrFlg = True
        Exit Do
    Loop

    '// �װ�̏ꍇ�ʹװ�߰�ނ�\��
    If Not m_bErrFlg Then
        w_sMsg = gf_GetErrMsg()
        Call gs_showMsgPage(w_sWinTitle,w_sMsgTitle,w_sMsg,w_sRetURL,w_sTarget)
    End If

	'// �I������
	Call gs_CloseDatabase()

End Sub

'********************************************************************************
'*  [�@�\]  �ϐ�������
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_ClearParam()

    m_iSyoriNen = ""
    m_iKyokanCd = ""
    m_iGakunen  = ""
    m_iClassNo  = ""

End Sub

'********************************************************************************
'*	[�@�\]	�p�����[�^�擾
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub s_GetParameter()

    m_iSyoriNen = Session("NENDO")
    m_iKyokanCd = Session("KYOKAN_CD")

End Sub

'********************************************************************************
'*  [�@�\]  �w�N�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeGakunenWhere()

    m_sGakunenWhere = ""
    m_sGakunenWhere = m_sGakunenWhere & " M05_NENDO = " & m_iSyoriNen
    m_sGakunenWhere = m_sGakunenWhere & " GROUP BY M05_GAKUNEN"

End Sub

'********************************************************************************
'*  [�@�\]  �N���X�R���{�Ɋւ���WHERE���쐬����
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub s_MakeClassWhere()

    m_sClassWhere = ""
    m_sClassWhere = m_sClassWhere & " M05_NENDO = " & m_iSyorinen

    If gf_IsNull(Trim(m_iGakunen)) Then
        '//�����\������1�N1�g��\������
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & C_FIRST_DISP_GAKUNEN
    Else
        m_sClassWhere = m_sClassWhere & " AND M05_GAKUNEN = " & cint(m_iGakunen)
    End If

End Sub

'********************************************************************************
'*	[�@�\]	HTML���o��
'*	[����]	�Ȃ�
'*	[�ߒl]	�Ȃ�
'*	[����]	
'********************************************************************************
Sub showPage()

	'---------- HTML START ----------
%>
<html>
<head>
    <link rel="stylesheet" href="../../common/style.css" type="text/css">
    <title>�s���i�w���ꗗ</title>
    <!--#include file="../../Common/jsCommon.htm"-->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--

    //************************************************************
    //  [�@�\]  �t�H�[�����[�h��
    //  [����]  
    //  [�ߒl]  
    //  [����]
    //************************************************************
	function jf_winload(){

	}

    //************************************************************
    //  [�@�\]  �\���{�^��������
    //  [����]  
    //  [�ߒl]  
    //  [����]
    //************************************************************
	function jf_Search(){
		document.body.style.cursor = "wait";
		with(document.frm){
			f_SetVal();
			target = "_LOWER";
			action = "Wait.asp";
			submit();
		}
	}

    //************************************************************
    //  [�@�\]  ���M�f�[�^��ϐ��ɃZ�b�g
    //  [����]  
    //  [�ߒl]  
    //  [����]
    //************************************************************
	function f_SetVal(){
		with(document.frm){
			hidGakunen.value = cboGakunenCd.value;
			hidClass.value   = cboClassCd.value;
			hidClassNM.value = cboClassCd.options[cboClassCd.selectedIndex].text;
		}
	}
	window.onload = jf_winload;
	//-->
	</SCRIPT>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<center>

<form name="frm" action="" target="main" Method="POST">

	<table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
	<tr>
		<td valign="top" align="center">

	<table cellspacing="0" cellpadding="0" border="0" width="98%">
	<tr>
	<td height="27" width="100%" align="left"
	>

	<DIV class=title>�s���i�w���ꗗ</DIV>

	</td
	>
	</tr
	>

	<tr
	><td height="4" width="5%" background="/cassist/image/table_sita.gif"
	><img src="/cassist/image/sp.gif"
	></td
	></tr
	>

	<tr
	><td height="10" class=title_Sub width="5%" align="right" valign="top"
	>

	<table class=title_Sub cellspacing="0" cellpadding="0" bgcolor=#393976 height="10" border="0"
	><tr
	><td align="center" valign="middle"
	><DIV class=title_Sub
	><img src="/cassist/image/sp.gif" width=8
        ><font color="#ffffff"
	>�ꗗ</font
	><img src="/cassist/image/sp.gif" width=8
	></DIV
	></td
	></tr
	></table
	>
	</td
	></tr
	></table>

	<br>

    <table border="0">
	    <tr>
	    	<td class="search">
				<table border="0" cellpadding="1" cellspacing="1">
					<tr>
						<td nowrap align="left">�N���X</td>
						<td align="left">
<%
		Call gf_ComboSet("cboGakunenCd",C_CBO_M05_CLASS_G,m_sGakunenWhere," style='width:40px;' ",False,m_iGakunen)
%>
						</td>
						<td align="left" width="20">�N</td>
						<td align="left" width="90">
<%
		Call gf_ComboSet("cboClassCd",C_CBO_M05_CLASS,m_sClassWhere," style='width:80px;' " & m_sClassOption,False,m_iClassNo)
%>
						</td>
						<td valign="bottom" align="right">
							<input class="button" type="button" onclick="javascript:jf_Search();" value="�@�\�@���@" <%= w_sDisabled %>>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

	</td>
	</tr>
	</table>

<input type="hidden" name="hidGakunen" value="<%= m_iGakunen %>">
<input type="hidden" name="hidClass"   value="<%= m_iClassNo %>">
<input type="hidden" name="hidClassNM" value="<%= m_sClassNM %>">
</form>
</center>

</body>
</html>
<%
'---------- HTML END   ----------
End Sub
%>