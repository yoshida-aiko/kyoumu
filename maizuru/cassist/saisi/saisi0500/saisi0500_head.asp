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

	Public m_iGakunen			'//�w�N
	Public m_iClassNo			'//�N���XNO
	Public m_iSyoriNen			'//�N�x
	Public m_iKyokanCd			'//��������
	Public m_iGakka				'//�w��
	Public m_sClass				'//�N���X
	Public m_sClassNM			'//�N���X��

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
		'//�l�̏�����
        Call s_ClearParam()

		'//�p�����[�^�擾
		Call s_GetParameter()

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
	m_sClass    = Request("hidClass")
	m_iGakunen  = Request("hidGakunen")
	m_sClassNM  = Request("hidClassNM")

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
			></table
			>
		</tr
		><tr>
			<td align="center">
				<table class="hyo" border="1" width="260" height="20">
				    <tr>
				        <th class="header" width="80"  align="center" nowrap>�N���X</th>
				        <td class="detail" width="100" align="center" nowrap><%= m_iGakunen %> �N</td>
				        <td class="detail" width="180" align="center" nowrap><%= m_sClassNM %></td>
				    </tr>
				</table>
			</td>
		</tr>
		<tr>
			<td valign="bottom" align="center">
				<table class="hyo" border="1">
					<tr>
						<th width="60"  height="30" class="header3" nowrap>�o�Ȕԍ�</th>
						<th width="150" height="30" class="header3" nowrap>����</th>
						<th width="150" height="30" class="header3" nowrap>�Ȗ�</th>
						<th width="50"  height="30" class="header3" nowrap>�N�x</th>
						<th width="70"  height="30" class="header3" nowrap>����/����</th>
						<th width="40"  height="30" class="header3" nowrap>���]��</th>
						<th width="40"  height="30" class="header3" nowrap>�V�]��</th>
						<th width="100" height="30" class="header3" nowrap>�S������</th>
					</tr>
				</table>
			</td>
		</tr>
	</table>

</form>
</center>

</body>
</html>
<%
'---------- HTML END   ----------
End Sub
%>