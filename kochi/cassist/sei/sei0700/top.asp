<%@ Language=VBScript %>
<%
'/************************************************************************
' �V�X�e����: ���������V�X�e��
' ��  ��  ��: �A���f����
' ��۸���ID : web/web0330/web0330_top.asp
' �@      �\: ��y�[�W �\������\��
'-------------------------------------------------------------------------
' ��      ��:�����R�[�h     ��      SESSION("KYOKAN_CD")
'            �N�x           ��      SESSION("NENDO")
' ��      ��:
' ��      �n:
' ��      ��:
'-------------------------------------------------------------------------
' ��      ��: 2001/07/10 �O�c
' ��      �X: 
'*************************************************************************/
%>
<!--#include file="../../Common/com_All.asp"-->
<%
'/////////////////////////// Ӽޭ�ٕϐ� /////////////////////////////
    Public m_sNendo             '�N�x
    Public m_PgMode             '�����ʃt���O
    Public m_sMsgTitle          '����
	
	Public m_Rs
	
	'�G���[�n
	Public  m_bErrFlg           '�װ�׸�
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
	
	On Error Resume Next
	Err.Clear
	
	m_PgMode=request("p_mode")
	
	Select Case m_PgMode
		Case "P_HAN0100"
		    m_sMsgTitle="���шꗗ�\"
		Case "P_KKS0200"
		    m_sMsgTitle="���ۈꗗ�\"
		Case "P_KKS0210"
		    m_sMsgTitle="�x���ꗗ�\"
		Case "P_KKS0220"
		    m_sMsgTitle="�s�����ۈꗗ�\"
		Case "P_HAN0111"
		    m_sMsgTitle="�]�_�ꗗ�\"
		Case Else
	End Select
	
	m_bErrFlg = False
	
	m_sNendo    = session("NENDO")
	
	Do
		'// �����`�F�b�N�Ɏg�p
		session("PRJ_No") = C_LEVEL_NOCHK
		
		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(session("PRJ_No"))
		
		'// �y�[�W��\��
		Call showPage()
		Exit Do
	Loop
	
End Sub

'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************
Sub showPage()
%>
	<html>
	<head>
		<link rel="stylesheet" href=../../common/style.css type="text/css">
		<title><%=m_sMsgTitle%></title>
	</head>
	
	<body>
	<form>
		<center><% call gs_title(m_sMsgTitle,"��@��") %><br></center>
		<INPUT TYPE=HIDDEN NAME=txtNendo    VALUE="<%=m_sNendo%>">
	</form>
	</body>
	</html>
<%
End Sub
%>
