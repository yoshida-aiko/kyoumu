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
	Public  m_Rs
	Public  m_RecCnt			'���R�[�h�J�E���g

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
		Session("PRJ_No") = "SEI0800"

		'// �s���A�N�Z�X�`�F�b�N
		Call gf_userChk(Session("PRJ_No"))

		'//���Ұ�SET
		Call s_SetParam()

		'// �w���ꗗ�擾�A�N�Z�X�`�F�b�N
		If Not f_GetStudent() Then m_bErrFlg = True : Exit Do

		'// �Y���҂����Ȃ��ꍇ
		If m_Rs.EOF Then
			Call gs_showWhitePage("�l���C�f�[�^�����݂��܂���B","���юQ��")
			Exit Do
		End If

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

End Sub

Function f_GetStudent()
'********************************************************************************
'*  [�@�\]  ���O�C���������S������N���X�̊w���ꗗ���擾����
'*  [����]  �Ȃ�
'*  [�ߒl]  True / False
'*  [����]  
'********************************************************************************

	On Error Resume Next
	Err.Clear

	Dim w_sSQL

	f_GetStudent = False

	w_sSQL = ""
	w_sSQL = w_sSQL & " SELECT "
	w_sSQL = w_sSQL & " 	T13_GAKUSEI_NO,  "
	w_sSQL = w_sSQL & " 	T13_GAKUSEKI_NO, "
	w_sSQL = w_sSQL & " 	T13_GAKUNEN, "
	w_sSQL = w_sSQL & " 	T11_SIMEI,   "
	w_sSQL = w_sSQL & " 	M05_CLASSMEI "
	w_sSQL = w_sSQL & " FROM "
	w_sSQL = w_sSQL & " 	T11_GAKUSEKI, "
	w_sSQL = w_sSQL & " 	T13_GAKU_NEN, "
	w_sSQL = w_sSQL & " 	M05_CLASS "
	w_sSQL = w_sSQL & " WHERE "
	w_sSQL = w_sSQL & " 	M05_NENDO      =  " & m_iNendo & " AND "
	w_sSQL = w_sSQL & " 	M05_TANNIN     = '" & m_sKyokanCd & "' AND"
	w_sSQL = w_sSQL & " 	T13_NENDO      =  M05_NENDO      AND "
	w_sSQL = w_sSQL & " 	T13_GAKUNEN    =  M05_GAKUNEN    AND "
	w_sSQL = w_sSQL & " 	T13_GAKKA_CD   =  M05_GAKKA_CD   AND "
	w_sSQL = w_sSQL & " 	T13_GAKUSEI_NO =  T11_GAKUSEI_NO     "

	If gf_GetRecordset(m_Rs,w_sSQL) <> 0 Then Exit Function

	'//ں��ރJ�E���g�擾
	m_RecCnt = gf_GetRsCount(m_Rs)

	f_GetStudent = True

End Function

Sub showPage()
'********************************************************************************
'*  [�@�\]  HTML���o��
'*  [����]  �Ȃ�
'*  [�ߒl]  �Ȃ�
'*  [����]  
'********************************************************************************

	On Error Resume Next
	Err.Clear

	Dim w_sCell

	w_sCell = "CELL1"

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

	<table  border="1" class="hyo">
		<tr>
			<th class="header2" width="550" align="center"><%=m_Rs("T13_GAKUNEN")%>�@�N�@<%=m_Rs("M05_CLASSMEI")%></th>
		</tr>
	<table>

	<br>

	<!-- TABLE�w�b�_�� -->
	<table border="1" class="hyo">
		<tr>
			<th width="100" class="header3" align="center" height="20">�w���ԍ�</th>
			<th width="100" class="header3" align="center" height="20">�o�Ȕԍ�</th>
			<th width="250" class="header3" align="center" height="20">���@�@��</th>
			<th width="100" class="header3" align="center" height="20">���ѕ\��</th>
		</tr>
	</table>

	<!-- TABLE���X�g�� -->
	<table class="hyo" align="center" border="1">

<%
	Do While Not m_Rs.EOF
		w_sCell = gf_IIF(w_sCell="CELL1","CELL2","CELL1")
%>
						<tr>
							<td width="100" class="<%=w_sCell%>" align="center" nowrap><%=m_Rs("T13_GAKUSEI_NO")%></td>
							<td width="100" class="<%=w_sCell%>" align="center" nowrap><%=m_Rs("T13_GAKUSEKI_NO")%></td>
							<td width="250" class="<%=w_sCell%>" align="left"   nowrap>�@<%=m_Rs("T11_SIMEI")%></td>
							<td width="100" class="<%=w_sCell%>" align="center" nowrap><input type="button" name="btnDisp" value="�\�@��"></td>
						</tr>

<%
		m_Rs.MoveNext
	Loop
%>

	</table>

</body>

</html>

<%
    '---------- HTML END   ----------
End Sub
%>
