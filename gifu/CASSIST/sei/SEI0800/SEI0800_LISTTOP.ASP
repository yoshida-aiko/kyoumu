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

	<table>
		<tr>
			<td width="510" align="center"><%=m_iGakunen%>�@�N�@�@<%=gf_GetClassName(m_iNendo,m_iGakunen,m_iClass)%></td>
		</tr>
	</table>

	<br>

	<!-- TABLE�w�b�_�� -->
	<table border="1" class="hyo">
		<tr>
			<th width="20"  class="header3" align="center" height="20"></th>
			<th width="100" class="header3" align="center" height="20">�w���ԍ�</th>
			<th width="100" class="header3" align="center" height="20">�o�Ȕԍ�</th>
			<th width="250" class="header3" align="center" height="20">���@�@��</th>
			<th width="60"  class="header3" align="center" height="20">���ѕ\��</th>
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
