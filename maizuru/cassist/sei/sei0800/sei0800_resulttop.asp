<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: ���юQ�Ɓi�������j
' ��۸���ID : sei/sei0800/default.asp
' �@      �\: 
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h		��		SESSION���i�ۗ��j
' ��      ��:�Ȃ�
' ��      �n:�����R�[�h		��		SESSION���i�ۗ��j
' ��      ��:
'           ���t���[���y�[�W
'-------------------------------------------------------------------------
' ��      ��: 2003/05/13 �A�c
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////

	Public  m_iNendo   			'�N�x
	Public  m_sKyokanCd			'���O�C������
	Public  m_bErrFlg			'�װ�׸�
	Dim     m_iGakunen
    Dim     m_iClass
	Dim     m_sGakName
    Dim     m_sGakNo

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

	Dim w_sWinTitle
	Dim w_sMsgTitle
	Dim w_sMsg
	Dim w_sRetURL
	Dim w_sTarget

	'Message�p�̕ϐ��̏�����
	w_sWinTitle="�L�����p�X�A�V�X�g"
	w_sMsgTitle="���юQ��"
	w_sMsg=""
	w_sRetURL="../../login/default.asp"
	w_sTarget="_parent"

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
'		Session("PRJ_No") = "SEI0800"

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(Session("PRJ_No"))

		'//���Ұ�SET
		Call s_SetParam()

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

'********************************************************************************
'*	[�@�\]	�S���ڂɈ����n����Ă����l��ݒ�
'********************************************************************************
Sub s_SetParam()

	m_iNendo    = Session("NENDO")
	m_sKyokanCd = Session("KYOKAN_CD")
	m_iGakunen  = Request("hidGakunen")
	m_iClass    = Request("hidClass")
	m_sGakName  = Request("hidGakuseiNM")
	m_sGakNo    = Request("hidGakuseiNo")

End Sub

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

	On Error Resume Next
	Err.Clear

%>
<html>

<head>
	<!--#include file="../../Common/jsCommon.htm"-->
	<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
	<!--
	//-->
	</SCRIPT>
	<link rel="stylesheet" href="../../common/style.css" type="text/css">
</head>

<body LANGUAGE="javascript">
	<center>
	<form name="frm" METHOD="post">
	<% call gs_title(" ���юQ�� "," �Q�@�� ") %>

	<table width="630" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td width="250" align="center" nowrap>�@<%=m_iGakunen%>�@�N�@�@<%=gf_GetClassName(m_iNendo,m_iGakunen,m_iClass)%>�@�@<%=m_sGakName%></td>
			<td width="380" align="right"  nowrap>
				<table width="380" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td width="180" align="right" nowrap>
							<table border="1" class="hyo" cellspacing="0" cellpadding="0">
								<tr>
									<td width="30" class="CELL1" height="20" style="background : #33CCFF;" nowrap></td>
								</tr>
							</table>
						</td>
						<td align="left" nowrap>= �C����</td>
						<td align="right" nowrap>
							<table border="1" class="hyo" cellspacing="0" cellpadding="0">
								<tr>
									<td width="30" class="CELL1" height="20" style="background : #FF9900;" nowrap></td>
								</tr>
							</table>
						</td>
						<td align="left" nowrap>= ���C��</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

	<br>

	<!-- TABLE�w�b�_�� -->
	<table border="1" class="hyo" width="630">
		<tr>
			<th width="30"  class="header3" align="center" height="  " rowspan="2" nowrap>&nbsp;</th>
			<th width="30"  class="header3" align="center" height="  " rowspan="2" nowrap>&nbsp;</th>
			<th width="250" class="header3" align="center" height="  " rowspan="2" nowrap>�ȁ@��</th>
			<th width="70"  class="header3" align="center" height="  " rowspan="2" nowrap>�C���P��</th>
			<th width="250" class="header3" align="center" height="20" colspan="5" nowrap>���@��</th>
		</tr>
		<tr>
			<th width="50" class="header2" align="center" height="20" nowrap>1�N</th>
			<th width="50" class="header2" align="center" height="20" nowrap>2�N</th>
			<th width="50" class="header2" align="center" height="20" nowrap>3�N</th>
			<th width="50" class="header2" align="center" height="20" nowrap>4�N</th>
			<th width="50" class="header2" align="center" height="20" nowrap>5�N</th>
		</tr>
	</table>

	</form>
	</cinter>
</body>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
